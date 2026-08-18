import os, json, re, random, math, io
from datetime import datetime, timezone, timedelta

from flask import render_template, request, redirect, url_for, flash, jsonify
from flask_login import login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from dotenv import load_dotenv
import requests

try:
    import PyPDF2
except Exception:
    PyPDF2 = None
try:
    from docx import Document
except Exception:
    Document = None

from .runtime import (
    app, db, login_manager, now_ist,
    BASE_DIR, IST,
    GROQ_API_KEY, GROQ_MODEL, GROQ_BASE, GROQ_TRANSCRIPTION_MODEL,
    ALLOWED_EXTENSIONS, SKILLS, CATEGORIES, CATEGORY_COUNTS
)
from .models import User, InterviewSession
from .services import *

@app.errorhandler(413)
def payload_too_large(_):
    flash('The uploaded file is too large. Please use a resume smaller than 8 MB.', 'error')
    return redirect(url_for('prepare')) if current_user.is_authenticated else redirect(url_for('index'))


@app.errorhandler(500)
def internal_error(_):
    db.session.rollback()
    flash('Something went wrong on the server. Please try again.', 'error')
    return redirect(url_for('dashboard')) if current_user.is_authenticated else redirect(url_for('index'))


@app.context_processor
def globals_ctx(): return {'year':datetime.now().year}

@app.route('/')
def index(): return render_template('index.html')

@app.route('/register',methods=['GET','POST'])
def register():
    if current_user.is_authenticated:return redirect(url_for('dashboard'))
    if request.method=='POST':
        name=request.form.get('name','').strip(); email=request.form.get('email','').strip().lower(); pwd=request.form.get('password','')
        if not name or not email or len(pwd)<6: flash('Enter valid details. Password must be at least 6 characters.','error')
        elif User.query.filter_by(email=email).first(): flash('Email already exists.','error')
        else:
            db.session.add(User(name=name,email=email,password=generate_password_hash(pwd)));db.session.commit();flash('Account created. Please login.','success');return redirect(url_for('login'))
    return render_template('register.html')

@app.route('/login',methods=['GET','POST'])
def login():
    if current_user.is_authenticated:return redirect(url_for('dashboard'))
    if request.method=='POST':
        u=User.query.filter_by(email=request.form.get('email','').strip().lower()).first()
        if u and check_password_hash(u.password,request.form.get('password','')):login_user(u);return redirect(url_for('dashboard'))
        flash('Invalid email or password.','error')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():logout_user();return redirect(url_for('index'))

@app.route('/dashboard')
@login_required
def dashboard():
    sessions=InterviewSession.query.filter_by(user_id=current_user.id).order_by(InterviewSession.created_at.desc()).limit(8).all()
    scores=[s.readiness_score for s in InterviewSession.query.filter_by(user_id=current_user.id).all() if s.readiness_score]
    latest = sessions[0] if sessions else None
    return render_template('dashboard.html', sessions=sessions, total=len(scores), avg=round(sum(scores)/len(scores),1) if scores else 0, latest=latest)

@app.route('/prepare',methods=['GET','POST'])
@login_required
def prepare():
    if request.method=='POST':
        role=request.form.get('job_role','').strip(); jd=request.form.get('job_desc','').strip(); f=request.files.get('resume')
        if not role or not jd or not f or not allowed_file(f.filename):flash('Provide target role, JD and a PDF/DOCX/TXT resume.','error');return render_template('prepare.html')
        fn=secure_filename(f'{current_user.id}_{int(datetime.now().timestamp())}_{f.filename}');path=os.path.join(app.config['UPLOAD_FOLDER'],fn);f.save(path)
        text=read_document(path)
        if len(normalize_text(text)) < 50:
            flash('Could not extract enough text from the uploaded resume. Use a text-based PDF, DOCX, or TXT file.', 'error')
            return render_template('prepare.html')
        try:
            profile = extract_profile(text)
            if not any(profile.get(k) for k in ('skills','projects','experience','certifications','education','achievements')):
                flash('The resume was readable, but no structured candidate evidence could be identified. Check the resume formatting and try again.', 'error')
                return render_template('prepare.html')
            analysis = analyze_resume_jd(profile, text, jd)
            questions = generate_questions(profile, jd)
            if len(questions) < 30:
                flash('The interview question set could not be completed safely from the resume evidence. Please try again.', 'error')
                return render_template('prepare.html')
            # Suggested answers are generated later from the Suggested Answers page.
            # Keeping them out of Analyze Resume removes dozens of extra Groq calls.
            for q in questions:
                q['suggested_answer'] = ''
                q['suggested_answer_available'] = True

            s = InterviewSession(
                user_id=current_user.id, job_role=role, job_desc=jd[:10000],
                resume_filename=fn, resume_text=text[:50000],
                profile=json.dumps(profile, ensure_ascii=False),
                jd_analysis=json.dumps(analysis, ensure_ascii=False),
                questions=json.dumps(questions, ensure_ascii=False)
            )
            db.session.add(s)
            db.session.commit()
            return redirect(url_for('result', sid=s.id))
        except Exception:
            db.session.rollback()
            app.logger.exception('Interview preparation failed')
            flash('We could not prepare this interview. Please check the resume and try again.', 'error')
            return render_template('prepare.html')
    return render_template('prepare.html')

@app.route('/results/<int:sid>')
@login_required
def results(sid):
    return redirect(url_for('result', sid=sid))

@app.route('/result/<int:sid>')
@login_required
def result(sid):
    s = owned(sid)
    if not s:
        return redirect(url_for('dashboard'))

    try:
        profile = json.loads(s.profile or '{}')
    except (TypeError, ValueError, json.JSONDecodeError):
        profile = {}

    try:
        questions = json.loads(s.questions or '[]')
    except (TypeError, ValueError, json.JSONDecodeError):
        questions = []

    try:
        analysis = json.loads(s.jd_analysis or '{}')
    except (TypeError, ValueError, json.JSONDecodeError):
        analysis = {}

    # Backward compatibility for sessions created before the schema fix.
    analysis.setdefault('match', analysis.get('match_score', 0))
    analysis.setdefault('jd_skills', find_skills(s.job_desc))
    analysis.setdefault('matched_skills', [])
    analysis.setdefault('missing_skills', [])
    analysis.setdefault('missing_keywords', [])
    analysis.setdefault('keyword_gaps', analysis.get('missing_keywords', []))
    analysis.setdefault('strengths', [])
    analysis.setdefault('priority_skills', analysis.get('missing_skills', []))

    # The result.html template expects resume_suggestions.
    # Older records may contain resume_improvements instead.
    suggestions = analysis.get('resume_suggestions')
    if not isinstance(suggestions, list):
        suggestions = []

    if not suggestions:
        old_suggestions = analysis.get('resume_improvements', [])
        if isinstance(old_suggestions, list):
            suggestions = old_suggestions

    if not suggestions:
        suggestions = [
            'Emphasize verified experience that directly matches the target JD.',
            'Add measurable outcomes to projects or experience where your resume supports them.',
            'Use JD keywords only when they truthfully describe your existing skills and experience.'
        ]

    analysis['resume_suggestions'] = suggestions
    analysis['resume_improvements'] = suggestions

    # Persist the repaired analysis so the same old interview does not
    # fail again on the next visit.
    s.jd_analysis = json.dumps(analysis)
    db.session.commit()

    return render_template(
        'result.html',
        s=s,
        profile=profile,
        skills=profile.get('skills', []),
        questions=questions,
        analysis=analysis
    )

@app.route('/api/transcribe', methods=['POST'])
@login_required
def transcribe_api():
    """Transcribe an uploaded microphone-only recording; no video is accepted."""
    audio = request.files.get('audio')
    if not audio:
        return jsonify({'ok': False, 'error': 'No microphone recording was received.'}), 400
    filename = secure_filename(audio.filename or 'answer.webm')
    content_type = audio.mimetype or 'audio/webm'
    if not content_type.startswith('audio/'):
        return jsonify({'ok': False, 'error': 'Only audio recordings are accepted.'}), 400
    data = audio.read()
    if not data:
        return jsonify({'ok': False, 'error': 'The microphone recording was empty.'}), 400
    if len(data) > 7 * 1024 * 1024:
        return jsonify({'ok': False, 'error': 'The recording is too large. Keep answers under a few minutes.'}), 413
    text = transcribe_audio_bytes(data, filename, content_type)
    if not text:
        return jsonify({'ok': False, 'error': 'Speech transcription failed. You can retry or type the answer manually.'}), 502
    return jsonify({'ok': True, 'text': text})


@app.route('/media-test')
@login_required
def media_test():
    return render_template('media_test.html')

@app.route('/interview/<int:sid>', methods=['GET', 'POST'])
@login_required
def interview(sid):
    s = owned(sid)

    if not s:
        return redirect(url_for('dashboard'))

    # -----------------------------
    # LOAD SAVED DATA
    # -----------------------------
    try:
        questions = json.loads(s.questions or '[]')
    except Exception:
        questions = []

    try:
        answers = json.loads(s.answers or '[]')
    except Exception:
        answers = []

    try:
        evals = json.loads(s.evaluations or '[]')
    except Exception:
        evals = []

    try:
        integrity = json.loads(s.integrity or '{}')
    except Exception:
        integrity = {}

    skipped = integrity.get('skipped', [])

    try:
        profile = json.loads(s.profile or '{}')
    except Exception:
        profile = {}

    # Make sure skipped is always a list
    if not isinstance(skipped, list):
        skipped = []

    # -----------------------------
    # NO QUESTIONS
    # -----------------------------
    if not questions:
        flash('No interview questions were generated for this resume.', 'error')
        return redirect(url_for('dashboard'))

    # We want exactly 30 questions in the interview.
    TOTAL_QUESTIONS = min(30, len(questions))

    # -----------------------------
    # POST
    # -----------------------------
    if request.method == 'POST':

        try:
            idx = int(request.form.get('index', '0'))
        except (ValueError, TypeError):
            idx = 0

        action = request.form.get('action', 'answer')

        # Never allow an invalid index
        if idx < 0 or idx >= TOTAL_QUESTIONS:
            return redirect(url_for('interview', sid=s.id))

        # -----------------------------
        # SKIP QUESTION
        # -----------------------------
        if action == 'skip':

            if idx not in skipped:
                skipped.append(idx)

            integrity['skipped'] = skipped
            s.integrity = json.dumps(integrity)

            db.session.commit()

            return redirect(url_for('interview', sid=s.id))

        # -----------------------------
        # NORMAL ANSWER
        # -----------------------------
        answer = request.form.get('answer', '').strip()

        # Also support typed_answer
        if not answer:
            answer = request.form.get('typed_answer', '').strip()

        if len(answer) < 3:
            flash(
                'Please record a clear answer or type an answer before continuing.',
                'error'
            )
            return redirect(url_for('interview', sid=s.id))

        q = questions[idx]

        # -----------------------------
        # PREVENT DUPLICATE SUBMISSION
        # -----------------------------
        already_answered = False

        for existing in answers:
            try:
                if int(existing.get('index', -1)) == idx:
                    already_answered = True
                    break
            except Exception:
                pass

        if not already_answered:

            # -----------------------------
            # EVALUATE ANSWER
            # -----------------------------
            try:
                ev = evaluate_answer(
                    q,
                    answer,
                    profile,
                    s.job_desc
                )
            except Exception as e:
                print("Evaluation error:", e)

                ev = {
                    "score": 0,
                    "feedback": "Answer recorded successfully.",
                    "strengths": [],
                    "improvements": []
                }

            # -----------------------------
            # SUGGESTED ANSWER
            # -----------------------------
            try:
                suggested = generate_suggested_answer(
                    q,
                    profile,
                    s.job_desc
                )
            except Exception as e:
                print("Suggested answer error:", e)
                suggested = ""

            answers.append({
                'index': idx,
                'question': q,
                'answer': answer,
                'suggested_answer': suggested,
                'created_at': now_ist().isoformat()
            })

            evals.append(ev)

            # Save immediately
            s.answers = json.dumps(answers)
            s.evaluations = json.dumps(evals)

            # -----------------------------
            # CALCULATE READINESS
            # -----------------------------
            try:
                overall, _ = compute_analytics(
                    questions,
                    answers,
                    evals
                )
                s.readiness_score = overall
            except Exception as e:
                print("Analytics calculation error:", e)

            db.session.commit()

        # -----------------------------
        # CHECK COMPLETION
        # -----------------------------
        answered_indexes = set()

        for item in answers:
            try:
                answered_indexes.add(int(item.get('index')))
            except Exception:
                pass

        completed_indexes = answered_indexes.union(
            set(int(x) for x in skipped if str(x).isdigit())
        )

        completed_count = len(
            [x for x in completed_indexes if 0 <= x < TOTAL_QUESTIONS]
        )

        if completed_count >= TOTAL_QUESTIONS:
            s.status = 'completed'
            db.session.commit()

            return redirect(
                url_for('analytics', sid=s.id)
            )

        # -----------------------------
        # IMPORTANT:
        # FIND NEXT UNANSWERED QUESTION
        # -----------------------------
        next_idx = None

        for i in range(TOTAL_QUESTIONS):
            if i not in completed_indexes:
                next_idx = i
                break

        if next_idx is None:
            s.status = 'completed'
            db.session.commit()

            return redirect(
                url_for('analytics', sid=s.id)
            )

        # Redirect with the next question explicitly.
        return redirect(
            url_for(
                'interview',
                sid=s.id,
                q=next_idx
            )
        )

    # =====================================================
    # GET
    # =====================================================

    # -----------------------------
    # BUILD COMPLETED INDEX SET
    # -----------------------------
    answered_indexes = set()

    for item in answers:
        try:
            answered_indexes.add(int(item.get('index')))
        except Exception:
            pass

    skipped_indexes = set()

    for item in skipped:
        try:
            skipped_indexes.add(int(item))
        except Exception:
            pass

    completed_indexes = answered_indexes.union(skipped_indexes)

    completed_indexes = {
        x for x in completed_indexes
        if 0 <= x < TOTAL_QUESTIONS
    }

    # -----------------------------
    # CURRENT QUESTION
    # -----------------------------
    requested_q = request.args.get('q')

    current_idx = None

    if requested_q is not None:
        try:
            requested_idx = int(requested_q)

            if (
                0 <= requested_idx < TOTAL_QUESTIONS
                and requested_idx not in completed_indexes
            ):
                current_idx = requested_idx

        except (ValueError, TypeError):
            pass

    # If no valid requested question,
    # find the first unanswered question.
    if current_idx is None:

        for i in range(TOTAL_QUESTIONS):

            if i not in completed_indexes:
                current_idx = i
                break

    # -----------------------------
    # COMPLETED
    # -----------------------------
    if current_idx is None:

        s.status = 'completed'
        db.session.commit()

        return redirect(
            url_for('analytics', sid=s.id)
        )

    # -----------------------------
    # PROGRESS
    # -----------------------------
    answered_count = len(completed_indexes)

    progress = round(
        (answered_count / TOTAL_QUESTIONS) * 100
    )

    progress = max(
        0,
        min(100, progress)
    )

    # -----------------------------
    # CURRENT QUESTION
    # -----------------------------
    current = questions[current_idx]

    # Support either string or dict questions
    if isinstance(current, dict):
        question_text = current.get(
            'question',
            current.get('text', '')
        )

        category = current.get(
            'category',
            current.get('sector', 'General')
        )

        difficulty = current.get(
            'difficulty',
            ''
        )

    else:
        question_text = str(current)
        category = 'General'
        difficulty = ''

    # -----------------------------
    # SUGGESTED ANSWER
    # -----------------------------
    suggested_answer = ''

    if isinstance(current, dict):
        suggested_answer = current.get(
            'suggested_answer',
            ''
        )

    if not suggested_answer:

        try:
            suggested_answer = generate_suggested_answer(
                current if isinstance(current, dict) else {'question': question_text, 'type': category, 'difficulty': difficulty, 'skill': ''},
                profile,
                s.job_desc
            )
        except Exception as e:
            print("Suggested answer error:", e)
            suggested_answer = ''

    # -----------------------------
    # READINESS
    # -----------------------------
    readiness = s.readiness_score or 0

    # -----------------------------
    # RENDER
    # -----------------------------
    return render_template(
        'interview.html',

        s=s,

        questions=questions,

        answers=answers,

        evals=evals,

        skipped=skipped,

        idx=current_idx,

        done=False,
        current=current,
        question_text=question_text,
        category=category,
        difficulty=difficulty,
        suggested_answer=suggested_answer,
        readiness=readiness,
        progress=progress,
        answered_count=answered_count,

        total_questions=TOTAL_QUESTIONS
    )
@app.route('/suggested-answers/<int:sid>')
@login_required
def suggested_answers(sid):
    s = owned(sid)
    if not s:
        return redirect(url_for('dashboard'))

    try:
        questions = json.loads(s.questions or '[]')
        answers = json.loads(s.answers or '[]')
        evals = json.loads(s.evaluations or '[]')
        profile = json.loads(s.profile or '{}')
    except (TypeError, ValueError, json.JSONDecodeError):
        flash('Could not load this interview data. Please try again.', 'error')
        return redirect(url_for('dashboard'))

    # Suggested answers are intentionally available BEFORE the live interview.
    # Generate missing answers dynamically from the current resume + exact question.
    missing = [i for i, q in enumerate(questions) if not q.get('suggested_answer')]
    if missing:
        batch = generate_suggested_answers_batch(
            [questions[i] for i in missing], profile, s.job_desc
        )
        for offset, idx in enumerate(missing):
            questions[idx]['suggested_answer'] = batch[offset]
            questions[idx]['suggested_answer_available'] = True
        s.questions = json.dumps(questions, ensure_ascii=False)

    # Preserve any answer/evaluation data from an ongoing or completed interview.
    answer_by_index = {a.get('index'): a for a in answers}
    eval_by_index = {a.get('index'): evals[i] for i, a in enumerate(answers) if i < len(evals)}
    items = []
    for i, question in enumerate(questions):
        answer_record = answer_by_index.get(i, {})
        items.append({
            'index': i,
            'question': question,
            'answer': answer_record.get('answer', ''),
            'suggested_answer': question.get('suggested_answer') or generate_suggested_answer(question, profile, s.job_desc),
            'evaluation': eval_by_index.get(i, {})
        })

    db.session.commit()
    return render_template(
        'suggested_answers.html',
        s=s, items=items,
        completed=(s.status == 'completed'),
        pre_interview=(len(answers) == 0)
    )

@app.route('/analytics/<int:sid>')
@login_required
def analytics(sid):
    s=owned(sid)
    if not s:return redirect(url_for('dashboard'))
    questions=json.loads(s.questions);answers=json.loads(s.answers);evals=json.loads(s.evaluations);overall,cat_scores=compute_analytics(questions,answers,evals)
    s.readiness_score=overall;db.session.commit();gaps=skill_gaps(s,evals)
    return render_template('analytics.html',s=s,evals=evals,answers=answers,questions=questions,cat_scores=cat_scores,weak=sorted(cat_scores.items(),key=lambda x:x[1])[:2],gaps=gaps)

@app.route('/skill-gaps/<int:sid>')
@app.route('/skill-gap/<int:sid>')
@login_required
def skill_gap(sid):
    s = owned(sid)
    if not s:
        return redirect(url_for('dashboard'))
    gaps = skill_gaps(s, json.loads(s.evaluations or '[]'))
    return render_template('skill_gap.html', s=s, gaps=gaps)

@app.route('/learning/<int:sid>')
@login_required
def learning(sid):
    s=owned(sid)
    if not s:return redirect(url_for('dashboard'))
    gaps=skill_gaps(s,json.loads(s.evaluations));priority=[g for g in gaps if g['priority'] in ('High','Medium')][:6]
    return render_template('learning.html',s=s,gaps=priority)

@app.route('/resume-improvement/<int:sid>')
@login_required
def resume_improvement(sid):
    s = owned(sid)
    if not s:
        return redirect(url_for('dashboard'))

    try:
        analysis = json.loads(s.jd_analysis or '{}')
    except (TypeError, ValueError, json.JSONDecodeError):
        analysis = {}

    analysis.setdefault('match', analysis.get('match_score', 0))
    analysis.setdefault('missing_skills', [])
    analysis.setdefault('missing_keywords', analysis.get('keyword_gaps', []))
    analysis.setdefault('keyword_gaps', analysis.get('missing_keywords', []))
    analysis.setdefault('matched_skills', [])
    analysis.setdefault('strengths', [])

    suggestions = analysis.get('resume_suggestions')
    if not isinstance(suggestions, list) or not suggestions:
        suggestions = analysis.get('resume_improvements', [])
    if not isinstance(suggestions, list) or not suggestions:
        suggestions = [
            'Emphasize verified experience relevant to this JD.',
            'Add measurable outcomes where supported by your actual work.',
            'Include relevant keywords only when they accurately describe your experience.'
        ]

    analysis['resume_suggestions'] = suggestions
    analysis['resume_improvements'] = suggestions
    s.jd_analysis = json.dumps(analysis)
    db.session.commit()

    return render_template(
        'resume_improvement.html',
        s=s,
        analysis=analysis
    )

@app.route('/history')
@login_required
def history():
    sessions = (
        InterviewSession.query
        .filter_by(user_id=current_user.id)
        .order_by(InterviewSession.created_at.desc())
        .all()
    )
    return render_template('history.html', sessions=sessions)


@app.post('/history/delete/<int:sid>')
@login_required
def delete_session(sid):
    s = owned(sid)
    if not s:
        flash('Interview not found.', 'error')
        return redirect(url_for('history'))
    db.session.delete(s)
    db.session.commit()
    flash('Interview removed from history.', 'success')
    return redirect(url_for('history'))

@app.route('/profile')
@login_required
def profile_home():
    s = InterviewSession.query.filter_by(user_id=current_user.id).order_by(InterviewSession.created_at.desc()).first()
    if not s:
        return redirect(url_for('prepare'))
    return redirect(url_for('profile', sid=s.id))

@app.route('/profile/<int:sid>')
@login_required
def profile(sid):
    s = owned(sid)

    if not s:
        return redirect(url_for('dashboard'))

    try:
        stored_profile = json.loads(
            s.profile or '{}'
        )
    except (TypeError, ValueError, json.JSONDecodeError):
        stored_profile = {}

    # Make sure all profile sections are lists
    profile_fields = [
        'skills',
        'projects',
        'education',
        'certifications',
        'achievements',
        'experience'
    ]

    for field in profile_fields:
        if not isinstance(stored_profile.get(field), list):
            stored_profile[field] = []

    return render_template(
        'profile.html',
        s=s,
        profile=stored_profile
    )
@app.route('/api/session/<int:sid>')
@login_required
def api_session(sid):
    s=owned(sid)
    if not s:return jsonify({'error':'forbidden'}),403
    return jsonify({'id':s.id,'questions':json.loads(s.questions),'answers':json.loads(s.answers),'evaluations':json.loads(s.evaluations),'integrity':json.loads(s.integrity or '{}'),'readiness':s.readiness_score})
