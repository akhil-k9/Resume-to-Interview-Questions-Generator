import os, json, re, random, math, io
from datetime import datetime, timezone, timedelta

from flask import request, redirect, url_for, flash
from flask_login import current_user
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
    GROQ_API_KEY, GROQ_MODEL, GROQ_BASE, GROQ_TRANSCRIPTION_MODEL,
    ALLOWED_EXTENSIONS, SKILLS, CATEGORIES, CATEGORY_COUNTS
)
from .models import User, InterviewSession

load_dotenv()

def load_user(uid):
    return db.session.get(User, int(uid))


def allowed_file(name):
    return '.' in name and name.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def read_document(path):
    ext = path.rsplit('.', 1)[-1].lower()
    try:
        if ext == 'txt':
            return open(path, encoding='utf-8', errors='ignore').read()
        if ext == 'pdf' and PyPDF2:
            with open(path, 'rb') as f:
                return '\n'.join((p.extract_text() or '') for p in PyPDF2.PdfReader(f).pages)
        if ext == 'docx' and Document:
            return '\n'.join(p.text for p in Document(path).paragraphs)
    except Exception:
        return ''
    return ''


def normalize_text(text):
    return re.sub(r'\s+', ' ', text or '').strip()


def find_skills(text):
    """Return only skills explicitly present in the supplied text.

    The caller decides which resume section to search. This keeps project
    descriptions from accidentally becoming the source of the profile skills.
    """
    lower = (text or '').lower()
    found = []
    for skill in SKILLS:
        if re.search(r'(?<![a-z0-9+#.])' + re.escape(skill.lower()) + r'(?![a-z0-9+#.])', lower):
            found.append(skill)
    return list(dict.fromkeys(found))


def _is_bullet_line(line):
    return bool(re.match(r'^(?:[•●▪◦*-]|\d+[.)])\s+', line or ''))


def _is_date_or_metric_line(line):
    value = normalize_text(line)
    if not value:
        return True
    if re.fullmatch(r'(?:19|20)\d{2}\s*[-–—/]\s*(?:(?:19|20)\d{2}|present|current)?', value, re.I):
        return True
    if re.fullmatch(r'(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+(?:19|20)\d{2}', value, re.I):
        return True
    if re.fullmatch(r'(?:cgpa|gpa|percentage|score|grade)\s*[:=-]?\s*[0-9.]+%?', value, re.I):
        return True
    return False


def _looks_like_project_title(line, next_line=''):
    """Conservative project-title helper for fallback extraction."""
    value = normalize_text(line).strip('•*-| ').strip()
    if not value or len(value) > 120:
        return False
    if _is_bullet_line(line) or _is_date_or_metric_line(value):
        return False

    lower = value.lower()
    description_starts = (
        'developed ', 'built ', 'designed ', 'implemented ', 'created ',
        'worked ', 'used ', 'utilized ', 'leveraged ', 'deployed ',
        'analyzed ', 'engineered ', 'integrated ', 'managed ',
        'responsible ', 'contributed ', 'performed ', 'conducted ',
        'developing ', 'using ', 'the project ', 'this project '
    )

    if lower.startswith(description_starts):
        return False
    if len(value.split()) > 12:
        return False
    if value.endswith(':'):
        return True
    if next_line:
        nxt = normalize_text(next_line).strip('•*-| ')
        if len(nxt) >= 40 and len(nxt) > len(value) * 1.2:
            return True
    return len(value.split()) <= 8 and not value.endswith(('.', ',', ';'))


def _normalize_section_heading(line):
    """Normalize a possible resume section heading."""
    value = normalize_text(line or '').lower()
    value = re.sub(r'^[\s•*|]+|[\s:|]+$', '', value)
    value = re.sub(r'\s*&\s*', ' and ', value)
    value = re.sub(r'\s+', ' ', value)
    return value.strip()


def _section_heading_kind(line):
    """Return the canonical section for a heading, or None for normal content."""
    key = _normalize_section_heading(line)
    if not key or len(key) > 80:
        return None

    exact = {
        'skills': 'skills',
        'technical skills': 'skills',
        'technical skill': 'skills',
        'technical expertise': 'skills',
        'technical competencies': 'skills',
        'core skills': 'skills',
        'key skills': 'skills',
        'skills and technologies': 'skills',
        'programming skills': 'skills',
        'tools and technologies': 'skills',
        'tech stack': 'skills',
        'technologies': 'skills',

        'projects': 'projects',
        'project': 'projects',
        'academic projects': 'projects',
        'academic project': 'projects',
        'personal projects': 'projects',
        'personal project': 'projects',
        'major projects': 'projects',
        'relevant projects': 'projects',
        'selected projects': 'projects',
        'key projects': 'projects',
        'project experience': 'projects',
        'project work': 'projects',

        'education': 'education',
        'educational background': 'education',
        'academic background': 'education',
        'academic qualifications': 'education',
        'educational qualifications': 'education',
        'educational details': 'education',
        'education details': 'education',
        'education and qualifications': 'education',
        'academic history': 'education',
        'academics': 'education',
        'qualifications': 'education',

        'experience': 'experience',
        'work experience': 'experience',
        'professional experience': 'experience',
        'professional background': 'experience',
        'employment': 'experience',
        'employment history': 'experience',
        'work history': 'experience',
        'career history': 'experience',
        'internship': 'experience',
        'internships': 'experience',
        'internship experience': 'experience',

        'certifications': 'certifications',
        'certification': 'certifications',
        'certificates': 'certifications',
        'certificate': 'certifications',
        'certifications and licenses': 'certifications',
        'licenses and certifications': 'certifications',
        'licenses certifications': 'certifications',
        'professional certifications': 'certifications',
        'certificates and training': 'certifications',
        'licenses': 'certifications',

        'achievements': 'achievements',
        'achievement': 'achievements',
        'awards': 'achievements',
        'award': 'achievements',
        'awards and achievements': 'achievements',
        'awards and honors': 'achievements',
        'honors': 'achievements',
        'honours': 'achievements',
        'honors and awards': 'achievements',
        'honours and awards': 'achievements',
        'accomplishments': 'achievements',
        'accomplishment': 'achievements'
    }

    if key in exact:
        return exact[key]

    words = set(re.findall(r'[a-z]+', key))
    if 'project' in words or 'projects' in words:
        if not any(w in words for w in ('experience', 'work', 'employment')):
            return 'projects'
    if any(w in words for w in ('certification', 'certifications', 'certificate', 'certificates')):
        return 'certifications'
    if any(w in words for w in ('achievement', 'achievements', 'award', 'awards', 'accomplishment', 'accomplishments')):
        return 'achievements'
    if any(w in words for w in ('education', 'educational', 'qualification', 'qualifications')):
        return 'education'
    if 'experience' in words and any(w in words for w in ('work', 'professional', 'employment', 'internship')):
        return 'experience'
    if 'skill' in words or 'skills' in words:
        return 'skills'
    return None


def _section_lines(text):
    """Split extracted resume text into bounded, evidence-only sections."""
    raw = text or ''
    sections = {
        'skills': [],
        'projects': [],
        'experience': [],
        'certifications': [],
        'education': [],
        'achievements': []
    }

    if not raw.strip():
        return [], sections

    aliases = {
        'TECHNICAL SKILLS': 'skills',
        'TECHNICAL EXPERTISE': 'skills',
        'TECHNICAL COMPETENCIES': 'skills',
        'SKILLS AND TECHNOLOGIES': 'skills',
        'SKILLS & TECHNOLOGIES': 'skills',
        'CORE SKILLS': 'skills',
        'KEY SKILLS': 'skills',
        'PROGRAMMING SKILLS': 'skills',
        'TOOLS AND TECHNOLOGIES': 'skills',
        'TOOLS & TECHNOLOGIES': 'skills',
        'TECH STACK': 'skills',
        'TECHNOLOGIES': 'skills',
        'SKILLS': 'skills',

        'ACADEMIC PROJECTS': 'projects',
        'ACADEMIC PROJECT': 'projects',
        'PERSONAL PROJECTS': 'projects',
        'PERSONAL PROJECT': 'projects',
        'MAJOR PROJECTS': 'projects',
        'RELEVANT PROJECTS': 'projects',
        'SELECTED PROJECTS': 'projects',
        'KEY PROJECTS': 'projects',
        'PROJECT EXPERIENCE': 'projects',
        'PROJECT WORK': 'projects',
        'PROJECTS': 'projects',
        'PROJECT': 'projects',

        'EDUCATIONAL QUALIFICATIONS': 'education',
        'ACADEMIC QUALIFICATIONS': 'education',
        'EDUCATIONAL BACKGROUND': 'education',
        'ACADEMIC BACKGROUND': 'education',
        'EDUCATIONAL DETAILS': 'education',
        'EDUCATION DETAILS': 'education',
        'EDUCATION AND QUALIFICATIONS': 'education',
        'EDUCATION & QUALIFICATIONS': 'education',
        'ACADEMIC HISTORY': 'education',
        'ACADEMICS': 'education',
        'QUALIFICATIONS': 'education',
        'EDUCATION': 'education',

        'PROFESSIONAL EXPERIENCE': 'experience',
        'PROFESSIONAL BACKGROUND': 'experience',
        'WORK EXPERIENCE': 'experience',
        'WORK HISTORY': 'experience',
        'EMPLOYMENT HISTORY': 'experience',
        'EMPLOYMENT': 'experience',
        'CAREER HISTORY': 'experience',
        'INTERNSHIP EXPERIENCE': 'experience',
        'INTERNSHIPS': 'experience',
        'INTERNSHIP': 'experience',
        'EXPERIENCE': 'experience',

        'CERTIFICATIONS AND LICENSES': 'certifications',
        'CERTIFICATIONS & LICENSES': 'certifications',
        'LICENSES AND CERTIFICATIONS': 'certifications',
        'LICENSES & CERTIFICATIONS': 'certifications',
        'PROFESSIONAL CERTIFICATIONS': 'certifications',
        'CERTIFICATES AND TRAINING': 'certifications',
        'CERTIFICATIONS': 'certifications',
        'CERTIFICATION': 'certifications',
        'CERTIFICATES': 'certifications',
        'CERTIFICATE': 'certifications',
        'LICENSES': 'certifications',

        'ACHIEVEMENTS & PARTICIPATION': 'achievements',
        'ACHIEVEMENT & PARTICIPATION': 'achievements',
        'AWARDS AND ACHIEVEMENTS': 'achievements',
        'AWARDS & ACHIEVEMENTS': 'achievements',
        'HONORS AND AWARDS': 'achievements',
        'HONOURS AND AWARDS': 'achievements',
        'ACCOMPLISHMENTS': 'achievements',
        'ACCOMPLISHMENT': 'achievements',
        'ACHIEVEMENTS': 'achievements',
        'ACHIEVEMENT': 'achievements',
        'AWARDS': 'achievements',
        'AWARD': 'achievements',
        'HONORS': 'achievements',
        'HONOURS': 'achievements'
    }

    aliases = sorted(aliases.items(), key=lambda x: len(x[0]), reverse=True)

    current = 'general'
    output_lines = []

    for raw_line in raw.splitlines():
        line = normalize_text(raw_line)
        if not line:
            continue

        # Remove only leading bullet decoration; preserve ':' because it is
        # useful for detecting project-title boundaries later.
        line = line.lstrip('•●▪◦*|- ').strip()
        if not line:
            continue

        exact = _section_heading_kind(line)
        if exact:
            current = exact
            continue

        upper = line.upper()
        handled = False

        for alias, section_name in aliases:
            if not upper.startswith(alias):
                continue

            # Do not split ordinary words that merely begin with a heading.
            if len(line) > len(alias):
                next_char = line[len(alias)]
                if next_char.isalnum() and alias not in {
                    'EDUCATION', 'PROJECTS', 'PROJECT', 'SKILLS',
                    'EXPERIENCE', 'CERTIFICATIONS', 'CERTIFICATION',
                    'CERTIFICATES', 'ACHIEVEMENTS', 'ACHIEVEMENT'
                }:
                    continue

            remainder = line[len(alias):].lstrip(' :|–—-•*')
            current = section_name

            if remainder:
                output_lines.append(remainder)
                sections[current].append(remainder)

            handled = True
            break

        if handled:
            continue

        output_lines.append(line)
        if current in sections:
            sections[current].append(line)

    for key, values in sections.items():
        cleaned = []
        seen = set()
        for value in values:
            value = normalize_text(value).strip('•●▪◦*|- ')
            if not value:
                continue
            # Preserve colon at the end of a project title.
            lookup = value.lower()
            if lookup in seen:
                continue
            seen.add(lookup)
            cleaned.append(value)
        sections[key] = cleaned

    return output_lines, sections


def _clean_profile_list(values, resume_text, limit=15, allowed_lines=None):
    """Keep concise evidence that is present in the resume and its section."""
    source = normalize_text(resume_text).lower()
    allowed_source = normalize_text(' '.join(allowed_lines or [])).lower() if allowed_lines is not None else source
    result = []
    seen = set()
    for value in values or []:
        if not isinstance(value, str):
            continue
        value = normalize_text(value).strip('•●▪◦*|- ')
        if not value or len(value) > 500:
            continue
        lower = value.lower()
        if lower in seen:
            continue
        if lower in source and lower in allowed_source:
            result.append(value)
            seen.add(lower)
    return result[:limit]


def _group_project_entries(lines, limit=15):
    """Reconstruct distinct project entries from compressed PDF text."""
    if not lines:
        return []

    text = ' '.join(normalize_text(x).strip('•●▪◦*|- ') for x in lines if normalize_text(x).strip('•●▪◦*|- '))
    text = normalize_text(text)
    if not text:
        return []

    # PDF extraction may return every word on its own line. Reconstruct titles
    # using the colon that commonly follows the project title. We also recognize
    # a title after a sentence boundary, so multiple projects can coexist on one line.
    title_pattern = re.compile(
        r'(?:^|[.!?]\s+|\s+-\s+)'
        r'([A-Z][A-Za-z0-9&()\[\]_/+\-–—’\' .]{2,110}?)'
        r':',
        re.M
    )

    candidates = []
    for match in title_pattern.finditer(text):
        title = normalize_text(match.group(1)).strip(' -–—|')
        if not title:
            continue
        if re.match(r'^(technologies|technology|tech stack|tools used|tools and technologies)$', title, re.I):
            continue
        if re.match(r'^(developed|built|designed|implemented|created|worked|used|utilized|leveraged|deployed|analyzed|engineered|integrated|managed|responsible|contributed|performed|conducted)\b', title, re.I):
            continue
        if len(title.split()) > 14 or len(title) > 120:
            continue
        candidates.append((match.start(1), match.end(), title))

    projects = []

    if candidates:
        for i, (_, end_pos, title) in enumerate(candidates):
            body_end = candidates[i + 1][0] if i + 1 < len(candidates) else len(text)
            body = text[end_pos:body_end].strip(' .|–—-')
            body = re.sub(r'\s+', ' ', body)

            # Remove a leading technology label only when it is directly attached;
            # retain the actual technology evidence in the project text.
            entry = title
            if body:
                entry += ' — ' + body
            projects.append(normalize_text(entry))
    else:
        # Fallback for resumes without colon-style project titles.
        current = None
        for i, line in enumerate(lines):
            value = normalize_text(line).strip('•●▪◦*|- ')
            if not value:
                continue
            next_line = normalize_text(lines[i + 1]) if i + 1 < len(lines) else ''
            if _looks_like_project_title(value, next_line):
                if current:
                    projects.append(normalize_text(current))
                current = value.rstrip(':').strip()
            elif current:
                current += ' ' + value
        if current:
            projects.append(normalize_text(current))

    result = []
    seen = set()
    for project in projects:
        project = normalize_text(project)
        if not project:
            continue
        key = project.lower()
        if key in seen:
            continue
        seen.add(key)
        result.append(project)
    return result[:limit]


def _group_education_entries(lines, limit=6):
    """Reconstruct education records from compressed PDF text."""
    if not lines:
        return []

    text = normalize_text(' '.join(normalize_text(x).strip('•●▪◦*|- ') for x in lines if normalize_text(x).strip('•●▪◦*|- ')))
    if not text:
        return []

    start_pattern = re.compile(
        r'(?<![A-Za-z0-9])('
        r'b\.?\s*tech|btech|b\.?\s*e\.?|be|'
        r'm\.?\s*tech|mtech|m\.?\s*e\.?|me|'
        r'mba|mca|bca|bsc|b\.?\s*sc|msc|m\.?\s*sc|'
        r'bba|bcom|mcom|'
        r'bachelor(?:\s+of)?(?:\'s)?|'
        r'master(?:\s+of)?(?:\'s)?|'
        r'doctorate|phd|diploma|associate degree|'
        r'intermediate(?:\([^)]*\))?|higher secondary|senior secondary|'
        r'secondary school|high school|ssc|hsc|'
        r'class\s*(?:10|11|12)|grade\s*(?:10|11|12)'
        r')(?=\b|\s)',
        re.I
    )

    matches = list(start_pattern.finditer(text))
    records = []

    if matches:
        for i, match in enumerate(matches):
            start = match.start()
            end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
            record = normalize_text(text[start:end]).strip('•●▪◦*|- ')
            if record:
                records.append(record)
    else:
        # Fallback for resumes that list the institution first.
        institution_pattern = re.compile(
            r'(?i)\b(?:[A-Z][A-Za-z.&-]+\s+){0,5}(?:university|college|institute|school|academy|polytechnic)\b[^.]*'
        )
        records = [normalize_text(m.group(0)) for m in institution_pattern.finditer(text)]

    # Clean records but keep evidence; never invent missing values.
    result = []
    seen = set()
    for record in records:
        record = normalize_text(record)
        if not record:
            continue
        # Collapse obvious duplicated punctuation/spaces created by PDF extraction.
        record = re.sub(r'\s*\|\s*', ' | ', record)
        key = record.lower()
        if key in seen:
            continue
        seen.add(key)
        result.append(record)
    return result[:limit]


def extract_profile(text):
    """Build a resume-grounded profile locally, with no extra Groq call."""
    lines, sections = _section_lines(text)

    email = re.search(r'[\w.+-]+@[\w-]+\.[\w.-]+', text or '')
    phone = re.search(r'(?:\+?\d[\d\s().-]{8,}\d)', text or '')

    skill_source = '\n'.join(sections.get('skills', []))
    skills = find_skills(skill_source) if skill_source.strip() else find_skills(text)

    projects = _group_project_entries(sections.get('projects', []))
    education = _group_education_entries(sections.get('education', []))

    def unique_section(name, limit):
        result = []
        seen = set()
        for value in sections.get(name, []):
            value = normalize_text(value).strip('•●▪◦*|- ')
            if not value:
                continue
            key = value.lower()
            if key in seen:
                continue
            seen.add(key)
            result.append(value)
        return result[:limit]

    return {
        'name': lines[0][:100] if lines else '',
        'email': email.group(0) if email else '',
        'phone': phone.group(0).strip() if phone else '',
        'skills': skills[:30],
        'projects': projects[:15],
        'certifications': unique_section('certifications', 12),
        'experience': unique_section('experience', 15),
        'education': education[:12],
        'achievements': unique_section('achievements', 12),
        'evidence': lines[:80]
    }

def chunk_text(text, size=850, overlap=120):
    words = normalize_text(text).split()
    if not words:
        return []
    chunks = []
    step = max(1, size - overlap)
    for i in range(0, len(words), step):
        c = ' '.join(words[i:i+size])
        if c:
            chunks.append(c)
    return chunks


def lexical_rag(query, documents, k=4):
    """Small local lexical retriever used to select resume evidence."""
    qtokens = set(re.findall(r'[a-zA-Z0-9+#.]{2,}', (query or '').lower()))
    scored = []
    for doc in documents:
        tokens = set(re.findall(r'[a-zA-Z0-9+#.]{2,}', doc.lower()))
        score = len(qtokens & tokens) / max(1, math.sqrt(len(qtokens) * len(tokens)))
        scored.append((score, doc))
    return [d for s, d in sorted(scored, reverse=True)[:k] if s > 0] or documents[:k]


def call_llm(system, user, temperature=0.2, max_tokens=1800, json_mode=False, retries=0):
    """Call Groq only. Secrets come exclusively from GROQ_API_KEY."""
    if not GROQ_API_KEY:
        return None
    payload = {
        'model': GROQ_MODEL,
        'messages': [
            {'role': 'system', 'content': system},
            {'role': 'user', 'content': user}
        ],
        'temperature': max(0.0, min(float(temperature), 1.0)),
        'max_tokens': max_tokens
    }
    if json_mode:
        payload['response_format'] = {'type': 'json_object'}

    for attempt in range(retries + 1):
        try:
            r = requests.post(
                GROQ_BASE.rstrip('/') + '/chat/completions',
                headers={
                    'Authorization': f'Bearer {GROQ_API_KEY}',
                    'Content-Type': 'application/json'
                },
                json=payload, timeout=35
            )
            if r.status_code == 429 and attempt < retries:
                import time
                time.sleep(1.5 * (attempt + 1))
                continue
            r.raise_for_status()
            data = r.json()
            return data['choices'][0]['message']['content']
        except Exception:
            if attempt >= retries:
                return None
    return None

def json_from_llm(raw):
    if not raw:
        return None
    raw = re.sub(r'^```(?:json)?\s*|\s*```$', '', raw.strip(), flags=re.I)
    try:
        return json.loads(raw)
    except Exception:
        a, b = raw.find('{'), raw.rfind('}')
        if a >= 0 and b > a:
            try: return json.loads(raw[a:b+1])
            except Exception: pass
        a, b = raw.find('['), raw.rfind(']')
        if a >= 0 and b > a:
            try: return json.loads(raw[a:b+1])
            except Exception: pass
    return None


def transcribe_audio_bytes(audio_bytes, filename='answer.webm', content_type='audio/webm'):
    """Transcribe microphone audio with Groq Whisper. Returns text or None."""
    if not GROQ_API_KEY or not audio_bytes:
        return None
    models = [GROQ_TRANSCRIPTION_MODEL]
    if GROQ_TRANSCRIPTION_MODEL != 'whisper-large-v3':
        models.append('whisper-large-v3')
    for model in models:
        try:
            response = requests.post(
                GROQ_BASE.rstrip('/') + '/audio/transcriptions',
                headers={'Authorization': f'Bearer {GROQ_API_KEY}'},
                files={'file': (filename, io.BytesIO(audio_bytes), content_type)},
                data={'model': model, 'language': 'en', 'response_format': 'json', 'temperature': '0'},
                timeout=90
            )
            response.raise_for_status()
            data = response.json()
            text = normalize_text(data.get('text', '')) if isinstance(data, dict) else ''
            if text:
                return text
        except Exception:
            app.logger.exception('Groq audio transcription failed with model %s', model)
    return None


def analyze_resume_jd(profile, resume_text, job_desc):
    """Fast local JD analysis; no Groq call is needed for simple matching."""
    jd_skills = find_skills(job_desc)
    resume_skills = set(profile.get('skills', []))
    missing = [s for s in jd_skills if s not in resume_skills]
    matched = [s for s in jd_skills if s in resume_skills]
    match = round((len(matched) / len(jd_skills)) * 100, 1) if jd_skills else 70.0

    strongest = matched[:3]
    strengths = ([f'{s} is explicitly present in the resume and requested by the JD.' for s in strongest]
                 if strongest else ['The resume contains explicit skills and project evidence that can be discussed in the interview.'])
    suggestions = [
        (f'Emphasize verified experience with {matched[0]}.' if matched else 'Emphasize the strongest skills already demonstrated in the resume.'),
        'Add measurable project outcomes only where the resume already supports them.',
        'Use job-description keywords only when they truthfully describe existing skills or experience.'
    ]
    return {
        'match': match,
        'match_score': match,
        'jd_skills': jd_skills,
        'matched_skills': matched,
        'missing_skills': missing,
        'missing_keywords': missing[:8],
        'keyword_gaps': missing[:8],
        'priority_skills': missing[:8],
        'strengths': strengths,
        'resume_suggestions': suggestions,
        'resume_improvements': suggestions,
        'retrieved_context': []
    }

def _available_counts(profile):
    """Adapt the preferred 30-question distribution to available resume evidence."""
    counts = CATEGORY_COUNTS.copy()
    available = []
    if profile.get('skills'): available.append('Technical')
    if profile.get('projects'): available.append('Projects')
    # HR/Behavioural are always possible, but their prompts must use resume evidence.
    available += ['HR', 'Behavioural']
    if profile.get('certifications'): available.append('Certifications')

    for category in ('Projects', 'Certifications'):
        if not profile.get('projects' if category == 'Projects' else 'certifications'):
            counts[category] = 0

    missing = 30 - sum(counts.values())
    recipients = [c for c in ('Technical','Projects','HR','Behavioural')
                  if counts[c] > 0 and c in available]
    if not profile.get('skills'):
        if counts['Technical']:
            counts['Technical'] = 0
            missing += CATEGORY_COUNTS['Technical']
        recipients = [c for c in recipients if c != 'Technical']
    i = 0
    while missing > 0 and recipients:
        counts[recipients[i % len(recipients)]] += 1
        missing -= 1
        i += 1
    return counts


def _resume_evidence(profile):
    return {
        'skills': profile.get('skills', []),
        'projects': profile.get('projects', []),
        'certifications': profile.get('certifications', []),
        'experience': profile.get('experience', []),
        'education': profile.get('education', []),
        'achievements': profile.get('achievements', [])
    }


def _question_supported(q, profile):
    """Reject questions that introduce resume facts not in the profile."""
    if not isinstance(q, dict):
        return False
    category = q.get('type')
    question = normalize_text(q.get('question', ''))
    if not question or category not in CATEGORIES:
        return False

    skills = [s.lower() for s in profile.get('skills', [])]
    projects = [normalize_text(x).lower() for x in profile.get('projects', [])]
    certs = [normalize_text(x).lower() for x in profile.get('certifications', [])]
    experience = [normalize_text(x).lower() for x in profile.get('experience', [])]
    ql = question.lower()

    # Detect known technology names in the question and require resume support.
    mentioned_known = [s for s in SKILLS if re.search(
        r'(?<![a-z0-9+#.])' + re.escape(s.lower()) + r'(?![a-z0-9+#.])', ql)]
    if mentioned_known and any(s not in skills for s in mentioned_known):
        return False

    if category == 'Projects' and projects:
        # At least one explicit project phrase/name must appear.
        if not any(p and p in ql for p in projects):
            # Also accept a distinctive first line token if exact full line is too long.
            project_tokens = []
            for p in projects:
                project_tokens.extend(re.findall(r'[a-zA-Z][a-zA-Z0-9_-]{3,}', p)[:4])
            if not any(t.lower() in ql for t in project_tokens):
                return False
    if category == 'Certifications' and certs:
        if not any(c in ql for c in certs):
            cert_tokens = []
            for c in certs:
                cert_tokens.extend(re.findall(r'[a-zA-Z][a-zA-Z0-9_-]{3,}', c)[:5])
            if not any(t.lower() in ql for t in cert_tokens):
                return False

    # Explicit experience/company facts must be traceable to experience evidence.
    known_companies = []
    for item in experience:
        known_companies += re.findall(r'\b[A-Z][A-Za-z0-9&.-]{2,}\b', item)
    for token in known_companies:
        if token.lower() in ql and token.lower() not in ql:
            return False

    # Prevent unsupported "my experience" claims when there is no experience.
    if category in ('HR', 'Behavioural') and not (profile.get('projects') or profile.get('experience') or profile.get('education') or profile.get('achievements')):
        # Only resume-independent self-reflection is safe when the resume has no evidence.
        allowed = ('skills', 'resume', 'learning', 'technology', 'role', 'strength', 'improve')
        if not any(w in ql for w in allowed):
            return False
    return True


def _dedupe_questions(items, profile, counts):
    final, seen = [], set()
    current = {c: 0 for c in CATEGORIES}
    for q in items or []:
        q = dict(q) if isinstance(q, dict) else {}
        q['question'] = normalize_text(q.get('question', ''))
        qtype = q.get('type')
        if current.get(qtype, 0) >= counts.get(qtype, 0):
            continue
        key = q['question'].lower()
        if not key or key in seen or not _question_supported(q, profile):
            continue
        seen.add(key)
        q.setdefault('difficulty', 'Intermediate')
        q.setdefault('skill', '')
        q['reason'] = 'Grounded in the uploaded resume evidence.'
        current[qtype] += 1
        final.append(q)
    return final, current, seen


def local_question_bank(profile, jd):
    """Dynamic offline question generation. It never contains candidate facts."""
    counts = _available_counts(profile)
    bank = []
    skills = profile.get('skills', [])
    projects = profile.get('projects', [])
    certs = profile.get('certifications', [])
    exp = profile.get('experience', [])
    education = profile.get('education', [])
    achievements = profile.get('achievements', [])

    variants = [
        'Focus on the decision you made.',
        'Focus on how you would verify the result.',
        'Focus on one limitation you noticed.',
        'Focus on how you would explain it to an interviewer.',
        'Focus on what you learned from the work.',
        'Focus on a practical trade-off.',
        'Focus on how you would improve the approach.',
        'Focus on how the evidence in your resume supports your explanation.',
        'Focus on one implementation detail you can defend.',
        'Focus on what you would do differently next time.',
        'Focus on the reasoning behind your approach.',
        'Focus on how you tested or validated the work.',
        'Focus on reliability and maintainability.',
        'Focus on what you can demonstrate confidently.',
        'Focus on a concrete example from the supplied resume.'
    ]
    technical_templates = [
        'How did you use {s} in the resume evidence you provided, and what design decision did you make?',
        'What practical problem did you solve using {s}, based on your resume?',
        'How would you test or validate the {s} work described on your resume?',
        'What trade-off did you consider when using {s} in your documented work?',
        'How would you troubleshoot a failure in the {s} work you describe on your resume?',
        'What part of your {s} work would you improve for a production setting?',
        'How does {s} interact with another technology explicitly listed on your resume?',
        'Explain one important concept of {s} that you can connect to your resume evidence.',
        'What implementation decision did you make while working with {s}?',
        'What did you learn from the {s} work shown on your resume?'
    ]
    for i in range(counts['Technical']):
        if skills:
            s = skills[i % len(skills)]
            q = technical_templates[i % len(technical_templates)].format(s=s)
            if i >= len(technical_templates):
                q += ' ' + variants[(i - len(technical_templates)) % len(variants)]
            bank.append({'type':'Technical','question':q,
                         'difficulty':'Intermediate','skill':s})
    project_templates = [
        'Walk me through {p}. What problem did it address and what does your resume show about your contribution?',
        'What technical challenge can you discuss from {p}, based on the evidence in your resume?',
        'Which technologies listed with {p} did you use, and why were they appropriate?',
        'How did you test, validate, or demonstrate {p} according to your resume?',
        'What would you improve in {p} without changing the facts documented on your resume?',
        'What did you learn from building or contributing to {p}?'
    ]
    for i in range(counts['Projects']):
        if projects:
            p = projects[i % len(projects)]
            q = project_templates[i % 6].format(p=p)
            if i >= 6:
                q += ' ' + variants[(i - 6) % len(variants)]
            bank.append({'type':'Projects','question':q,
                         'difficulty':'Intermediate','skill':''})
    cert_templates = [
        'What did you learn from {c} that you can explain confidently?',
        'How could you apply knowledge from {c} to a practical task?',
        'Which concept associated with {c} would you like to strengthen further?',
        'How has completing {c} influenced the way you approach technical learning?'
    ]
    for i in range(counts['Certifications']):
        if certs:
            c = certs[i % len(certs)]
            q = cert_templates[i % 4].format(c=c)
            if i >= 4:
                q += ' ' + variants[(i - 4) % len(variants)]
            bank.append({'type':'Certifications','question':q,
                         'difficulty':'Intermediate','skill':''})

    evidence_anchor = projects[0] if projects else (exp[0] if exp else (education[0] if education else 'the evidence in your resume'))
    hr_templates = [
        f'Tell me about yourself using the strongest evidence shown in your resume, especially {evidence_anchor}.',
        'Which part of your resume best demonstrates that you are ready for this role, and why?',
        'What is one skill or area shown on your resume that you want to strengthen next?',
        'How do the projects, education, skills, or achievements on your resume reflect your learning approach?',
        'Which item on your resume would you most like an interviewer to ask about, and why?'
    ]
    for i in range(counts['HR']):
        q = hr_templates[i % len(hr_templates)]
        if i >= len(hr_templates):
            q += ' ' + variants[(i - len(hr_templates)) % len(variants)]
        bank.append({'type':'HR','question':q,
                     'difficulty':'Intermediate','skill':''})

    beh_templates = [
        f'Describe a challenge you can support with evidence from {evidence_anchor}, and explain how you approached it.',
        'Tell me about a time you had to learn something reflected in your resume. What was your approach?',
        'Describe a problem from your documented project, education, internship, or achievement work and how you handled it.',
        'Tell me about a piece of resume evidence that shows how you respond when something does not work as expected.',
        'What did you learn from one experience or project explicitly documented on your resume?'
    ]
    for i in range(counts['Behavioural']):
        q = beh_templates[i % len(beh_templates)]
        if i >= len(beh_templates):
            q += ' ' + variants[(i - len(beh_templates)) % len(variants)]
        bank.append({'type':'Behavioural','question':q,
                     'difficulty':'Intermediate','skill':''})
    return bank


def generate_questions(profile, jd):
    counts = _available_counts(profile)
    evidence = _resume_evidence(profile)
    chunks = chunk_text(json.dumps(evidence, ensure_ascii=False)) + chunk_text(jd)
    prompt = f"""Generate interview questions using ONLY the candidate resume evidence below.
Preferred categories are Technical 10, Projects 6, HR 5, Behavioural 5, Certifications 4,
but the requested counts below have already been adapted to what actually exists in the resume:
{json.dumps(counts)}

Do not invent skills, projects, certifications, employers, responsibilities, achievements,
metrics or technologies. Technical questions must reference a skill explicitly present.
Project questions must reference an explicit project/evidence item. Certification questions
must reference an explicit certification. HR and Behavioural questions must be answerable from
resume evidence or honest self-reflection without assuming an experience not shown.
Return a JSON object with key "questions" containing an array. Each item must have:
type, question, difficulty, skill, reason.

RESUME EVIDENCE:
{json.dumps(evidence, ensure_ascii=False)[:12000]}

RETRIEVED EVIDENCE:
{json.dumps(lexical_rag('skills projects certifications experience education achievements', chunks, 8), ensure_ascii=False)[:9000]}

TARGET JOB DESCRIPTION:
{jd[:6000]}"""
    parsed = json_from_llm(call_llm(
        'You are a strict adaptive interviewer. Return JSON only and never fabricate resume facts.',
        prompt, 0.15, 5000, json_mode=True
    ))
    llm_items = parsed.get('questions', []) if isinstance(parsed, dict) else []
    final, current, seen = _dedupe_questions(llm_items, profile, counts)

    # Deterministic fallback fills any missing category without candidate-specific data.
    fallback, _, _ = _dedupe_questions(local_question_bank(profile, jd), profile, counts)
    for q in fallback:
        if len(final) >= 30:
            break
        typ = q['type']
        if current[typ] < counts[typ] and q['question'].lower() not in seen:
            final.append(q); current[typ] += 1; seen.add(q['question'].lower())

    # If a generated batch still lacks questions, use more dynamic templates based
    # on the available evidence; never pad with invented candidate facts.
    if len(final) < 30:
        extra = local_question_bank(profile, jd)
        for q in extra:
            if len(final) >= 30: break
            key = q['question'].lower()
            if key not in seen and current.get(q['type'], 0) < counts.get(q['type'], 0):
                final.append(q); current[q['type']] += 1; seen.add(key)

    random.shuffle(final)
    for q in final:
        q['suggested_answer'] = ''
        q['suggested_answer_available'] = True
    return final[:30]


def _answer_prompt(question, profile, jd):
    qtype = question.get('type', 'Interview')
    relevant = _resume_evidence(profile)
    # Retrieve only evidence lexically related to the exact current question.
    docs = []
    for field in ('skills','projects','certifications','experience','education','achievements'):
        docs += [str(x) for x in relevant.get(field, [])]
    retrieved = lexical_rag(question.get('question', ''), docs, 6)
    return f"""Write a suggested interview answer for the EXACT current question.
Use ONLY the supplied resume evidence. Never invent a result, metric, employer,
technology, responsibility, project detail, certification or achievement.
If a detail is missing, say so naturally and answer using only what can be supported.
For HR/Behavioural use a concise STAR-like structure when the evidence supports it.
For technical/project questions, directly address the exact question and reference the
relevant resume evidence. Sound like a natural student/fresher, not a generic template.
Return only the answer text, about 70-130 words.

QUESTION CATEGORY: {qtype}
EXACT QUESTION: {question.get('question','')}
QUESTION SKILL: {question.get('skill','')}

RESUME PROFILE:
{json.dumps(relevant, ensure_ascii=False)[:12000]}

MOST RELEVANT RESUME EVIDENCE:
{json.dumps(retrieved, ensure_ascii=False)[:6000]}

TARGET JOB:
{jd[:4000]}"""


def _safe_fallback_answer(question, profile):
    q = question.get('question', '')
    qtype = question.get('type', 'Interview')
    skills = profile.get('skills', [])
    projects = profile.get('projects', [])
    certs = profile.get('certifications', [])
    exp = profile.get('experience', [])
    edu = profile.get('education', [])
    anchor = projects[0] if projects else (exp[0] if exp else (edu[0] if edu else 'the experience documented on my resume'))
    if qtype == 'Technical' and skills:
        return (f"For this question, I would explain how {question.get('skill') or skills[0]} is represented in my resume "
                f"and focus on the practical work I can demonstrate. I would describe the approach I used, why I chose it, "
                f"how I tested it, and what I learned. I would avoid claiming a detail that is not documented and be clear "
                f"about what I know versus what I would need to learn.")
    if qtype == 'Projects' and projects:
        selected = next((p for p in projects if any(
            token.lower() in q.lower() for token in re.findall(r'[a-zA-Z][a-zA-Z0-9_-]{3,}', normalize_text(p))[:6]
        )), projects[0])
        return (f"I would start with {selected}, which is explicitly listed on my resume. I would explain the problem it addressed, "
                f"the contribution and technologies that are actually documented, and then discuss a challenge I can support "
                f"with evidence. I would finish with what I learned and what I would improve next, without inventing metrics "
                f"or functionality that is not on my resume.")
    if qtype == 'Certifications' and certs:
        selected = next((c for c in certs if any(
            token.lower() in q.lower() for token in re.findall(r'[a-zA-Z][a-zA-Z0-9_-]{3,}', normalize_text(c))[:6]
        )), certs[0])
        return (f"I would focus on {selected} and explain the concepts I can genuinely demonstrate from that certification. "
                f"If the question asks about practical application, I would connect it only to work that appears on my resume. "
                f"Where I have not yet applied a concept in a real project, I would say that honestly and explain how I would "
                f"approach applying it.")
    return (f"I would answer this by connecting it to the evidence on my resume, especially {anchor}. "
            f"I would explain the situation or context, the specific action I can genuinely support, and what I learned. "
            f"If the resume does not contain a specific result or metric, I would not invent one; I would describe the outcome "
            f"at the level supported by the evidence.")

def _question_terms(question):
    return set(re.findall(r'[a-zA-Z0-9+#.]{3,}', normalize_text(question).lower()))


def _answer_similarity(a, b):
    """Token Jaccard similarity used to reject near-duplicate suggestions."""
    ta = set(re.findall(r'[a-zA-Z0-9+#.]{3,}', normalize_text(a).lower()))
    tb = set(re.findall(r'[a-zA-Z0-9+#.]{3,}', normalize_text(b).lower()))
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / max(1, len(ta | tb))


def _answer_matches_question(question, answer, profile, existing_answers=None):
    answer = normalize_text(answer)
    if len(answer) < 45:
        return False
    lower = answer.lower()
    ql = normalize_text(question.get('question', '')).lower()
    supported_skills = {normalize_text(s).lower() for s in profile.get('skills', [])}

    # Never allow a generated answer to introduce a known technology absent from the resume.
    for skill in SKILLS:
        if re.search(r'(?<![a-z0-9+#.])' + re.escape(skill.lower()) + r'(?![a-z0-9+#.])', lower):
            if skill.lower() not in supported_skills:
                return False

    qtype = question.get('type')
    if qtype == 'Technical' and question.get('skill'):
        skill = normalize_text(question['skill']).lower()
        if skill not in lower:
            return False
    elif qtype == 'Projects' and profile.get('projects'):
        matched = False
        for project in profile['projects']:
            tokens = re.findall(r'[a-zA-Z][a-zA-Z0-9_-]{3,}', normalize_text(project))[:8]
            if any(token.lower() in ql and token.lower() in lower for token in tokens):
                matched = True
                break
        if not matched:
            return False
    elif qtype == 'Certifications' and profile.get('certifications'):
        matched = False
        for cert in profile['certifications']:
            tokens = re.findall(r'[a-zA-Z][a-zA-Z0-9_-]{3,}', normalize_text(cert))[:8]
            if any(token.lower() in ql and token.lower() in lower for token in tokens):
                matched = True
                break
        if not matched:
            return False

    # The answer must share meaningful language with the exact question/evidence.
    qterms = {t for t in _question_terms(ql) if t not in {'what','which','how','why','your','tell','about','from','the','and','you','did','would','can','could'}}
    aterms = set(re.findall(r'[a-zA-Z0-9+#.]{3,}', lower))
    evidence_terms = set()
    for field in ('skills','projects','certifications','experience','education','achievements'):
        for item in profile.get(field, []) if isinstance(profile.get(field, []), list) else []:
            evidence_terms |= set(re.findall(r'[a-zA-Z0-9+#.]{3,}', normalize_text(item).lower()))
    if qterms and not ((qterms & aterms) or (evidence_terms & aterms)):
        return False

    for old in existing_answers or []:
        if _answer_similarity(answer, old) >= 0.62:
            return False
    return True


def _answer_prompt(question, profile, jd, variation_hint=''):
    qtype = question.get('type', 'Interview')
    relevant = _resume_evidence(profile)
    docs = []
    for field in ('skills','projects','certifications','experience','education','achievements'):
        docs += [str(x) for x in relevant.get(field, [])]
    retrieved = lexical_rag(question.get('question', ''), docs, 7)
    return f"""Write a suggested interview answer for the EXACT current question.\n\nHARD RULES:\n- Use ONLY supplied resume evidence. Never invent a result, metric, employer, technology, responsibility, project detail, certification, achievement, or event.\n- Answer THIS question directly; do not give a generic interview introduction.\n- Mention the specific resume entity named or implied by the question whenever the evidence supports it.\n- If the resume does not contain enough detail for a requested fact, explicitly say that the resume does not specify that detail and then explain only what is supported.\n- Do not reuse stock phrases or another question's answer. Start from the question's actual subject.\n- For technical questions: explain the practical use/decision/validation of the named skill, not a textbook definition unless the question asks for one.\n- For project questions: focus on the named project and only its documented evidence.\n- For certification questions: focus on the named certification and supported learning/application.\n- For HR/Behavioural: use a concise STAR-like structure only when the resume contains enough evidence; otherwise be honest about what is and is not documented.\n- Sound like a natural student/fresher.\n- Return ONLY the answer text, about 80-140 words.\n\nQUESTION CATEGORY: {qtype}\nEXACT QUESTION: {question.get('question','')}\nQUESTION SKILL: {question.get('skill','')}\nVARIATION REQUIREMENT: {variation_hint}\n\nRESUME PROFILE:\n{json.dumps(relevant, ensure_ascii=False)[:14000]}\n\nMOST RELEVANT RESUME EVIDENCE:\n{json.dumps(retrieved, ensure_ascii=False)[:7000]}\n\nTARGET JOB:\n{jd[:4000]}"""


def _safe_fallback_answer(question, profile, variation_index=0):
    """Deterministic, candidate-specific fallback; every question gets different evidence/angle."""
    q = normalize_text(question.get('question', ''))
    qtype = question.get('type', 'Interview')
    skills = profile.get('skills', []) if isinstance(profile.get('skills', []), list) else []
    projects = profile.get('projects', []) if isinstance(profile.get('projects', []), list) else []
    certs = profile.get('certifications', []) if isinstance(profile.get('certifications', []), list) else []
    exp = profile.get('experience', []) if isinstance(profile.get('experience', []), list) else []
    edu = profile.get('education', []) if isinstance(profile.get('education', []), list) else []
    achievements = profile.get('achievements', []) if isinstance(profile.get('achievements', []), list) else []
    evidence = projects + exp + certs + achievements + edu
    if not evidence:
        evidence = ['the information documented in my uploaded resume']
    anchor = evidence[variation_index % len(evidence)]

    if qtype == 'Technical' and skills:
        skill = question.get('skill') or skills[variation_index % len(skills)]
        angles = [
            f'practical use of {skill}', f'design reasoning around {skill}',
            f'testing or validation of {skill}', f'a limitation of the {skill} work',
            f'how the {skill} work could be improved'
        ]
        angle = angles[variation_index % len(angles)]
        return (f"For this question, I would focus on {angle} as it appears in my resume. "
                f"I would explain the part of the work I can directly support from my resume, the approach I took, "
                f"and the reasoning behind it. I would avoid adding a tool, metric, or responsibility that is not documented. "
                f"If the interviewer asks for a detail that my resume does not specify, I would be transparent about that "
                f"and then explain what I understand and how I would verify or improve the approach.")
    if qtype == 'Projects' and projects:
        project = next((p for p in projects if any(t.lower() in q.lower() for t in re.findall(r'[a-zA-Z][a-zA-Z0-9_-]{3,}', normalize_text(p))[:8])), projects[variation_index % len(projects)])
        angles = ['the problem and purpose', 'my documented contribution', 'the technologies explicitly associated with it', 'testing or validation', 'a challenge supported by the resume', 'what I learned from it']
        return (f"I would answer this by focusing on {project} and the {angles[variation_index % len(angles)]}. "
                f"I would first explain only the project details that are stated in my resume, then connect those details "
                f"to the exact point in the question. I would distinguish clearly between what I actually documented and "
                f"what I would investigate further rather than inventing a result or implementation detail. This keeps my answer "
                f"specific to the project while remaining accurate to my resume.")
    if qtype == 'Certifications' and certs:
        cert = next((c for c in certs if any(t.lower() in q.lower() for t in re.findall(r'[a-zA-Z][a-zA-Z0-9_-]{3,}', normalize_text(c))[:8])), certs[variation_index % len(certs)])
        angles = ['the main learning', 'practical application', 'a concept I want to strengthen', 'how it changed my learning approach']
        return (f"I would connect this question to {cert}, which is listed in my resume. I would focus on {angles[variation_index % len(angles)]} "
                f"and explain the knowledge I can genuinely support. If the question asks about a real-world application that is not "
                f"documented in my resume, I would say that directly instead of claiming an experience I do not have. "
                f"I would then explain how I would apply or strengthen the relevant knowledge in a practical setting.")
    return (f"I would answer this by connecting it to {anchor}, because that is evidence actually present in my resume. "
            f"The key point I would address is the specific question: {q}. I would explain the context and the part I can "
            f"personally support from my documented experience, then describe what I learned or would improve where the resume "
            f"supports that statement. I would not add a company, technology, responsibility, metric, or achievement that is not listed.")


def generate_suggested_answer(question, profile, jd, existing_answers=None, variation_index=0):
    hints = [
        'Use the named resume entity first and focus on practical details.',
        'Emphasize the decision/reasoning behind the documented work.',
        'Emphasize validation/testing or how the work was checked.',
        'Emphasize a supported limitation, challenge, or trade-off.',
        'Emphasize learning and a realistic improvement, without inventing outcomes.'
    ]
    for attempt in range(3):
        raw = call_llm(
            'You are a precise interview coach. The exact question is authoritative. Never fabricate resume facts. Every answer must be materially different from the supplied existing answers.',
            _answer_prompt(question, profile, jd, hints[(variation_index + attempt) % len(hints)]),
            0.25 if attempt else 0.18, 1000, False, 1
        )
        if raw:
            cleaned = re.sub(r'^```(?:text|json)?\s*|\s*```$', '', raw.strip(), flags=re.I)
            if _answer_matches_question(question, cleaned, profile, existing_answers):
                return cleaned
    return _safe_fallback_answer(question, profile, variation_index)


def generate_suggested_answers_batch(questions, profile, jd):
    """Generate answers one question at a time so one generic batch response cannot contaminate all 30 answers."""
    answers = []
    for i, q in enumerate(questions):
        answer = generate_suggested_answer(q, profile, jd, existing_answers=answers, variation_index=i)
        # A final deterministic uniqueness guard, including fallback answers.
        if any(_answer_similarity(answer, old) >= 0.62 for old in answers):
            answer = _safe_fallback_answer(q, profile, i + len(answers))
        if any(_answer_similarity(answer, old) >= 0.62 for old in answers):
            # Add a question-specific closing based on the exact question terms.
            anchor_terms = [t for t in _question_terms(q.get('question','')) if len(t) > 4][:4]
            answer = normalize_text(answer) + ' I would keep the explanation focused on ' + ', '.join(anchor_terms or ['the exact point asked']) + '.'
        answers.append(answer)
    return answers


def evaluate_answer(q, answer, profile, jd):
    answer=normalize_text(answer); words=re.findall(r'\b\w+\b',answer); length=len(words)
    qterms=set(re.findall(r'[a-zA-Z0-9+#.]{3,}',q.get('question','').lower())); aterms=set(w.lower() for w in words)
    overlap=len(qterms & aterms); specificity=min(10,4+overlap*0.7+(2 if length>=60 else 0))
    completeness=min(10,2.5+length/20); clarity=9 if length>=70 else 8 if length>=40 else 6 if length>=20 else 3
    relevance=min(10,4+overlap*0.8+(1 if any(s.lower() in answer.lower() for s in profile.get('skills',[])) else 0))
    overall=round((specificity+completeness+clarity+relevance)/4,1)
    llm = call_llm('Evaluate the answer only against the question and supplied evidence. Do not invent facts. Return JSON with overall, relevance, correctness, completeness, clarity, feedback, missing_points.',
                   f'QUESTION: {q.get("question")}\nANSWER: {answer}\nRESUME PROFILE: {json.dumps(profile)[:5000]}\nJD: {jd[:4000]}',0.15,900)
    parsed=json_from_llm(llm)
    if isinstance(parsed,dict):
        for k in ['overall','relevance','correctness','completeness','clarity']:
            if k in parsed:
                try: parsed[k]=float(parsed[k])
                except: pass
        parsed.setdefault('feedback','Good effort. Add concrete evidence and outcomes.'); parsed.setdefault('missing_points','Add one specific example, reasoning and measurable result where available.'); return parsed
    return {'overall':overall,'relevance':round(relevance,1),'correctness':round(specificity,1),'completeness':round(completeness,1),'clarity':round(clarity,1),
            'feedback':'Strong structure.' if overall>=8 else 'Good start. Add concrete evidence, reasoning and outcomes.' if overall>=6 else 'Needs more depth. Explain your action and result.',
            'missing_points':'Add a concrete example, technical reasoning and result.' if overall<8 else 'Consider a deeper trade-off or measurable outcome.'}


def adaptive_next(questions, answers, evaluations, skipped):
    used={a.get('index') for a in answers} | set(skipped)
    candidates=[(i,q) for i,q in enumerate(questions) if i not in used]
    if not candidates: return len(questions)
    if not evaluations: return random.choice(candidates)[0]
    last=float(evaluations[-1].get('overall',0)); prev=answers[-1].get('index') if answers else None
    prevq=questions[prev] if prev is not None and prev < len(questions) else {}
    if last>=8:
        related=[x for x in candidates if x[1].get('skill') and x[1].get('skill')==prevq.get('skill')]
        if related: return random.choice(related)[0]
    if last<5.5:
        easier=[x for x in candidates if x[1].get('type') in ('Technical','Projects') and x[1].get('difficulty') in ('Beginner','Easy','Intermediate')]
        if easier: return random.choice(easier)[0]
    return random.choice(candidates)[0]


def compute_analytics(questions, answers, evaluations):
    answered_by={a.get('index'):a for a in answers}; eval_by={a.get('index'):e for a,e in zip(answers,evaluations)}
    cats={c:[] for c in CATEGORIES}
    for idx,e in eval_by.items():
        if idx is not None and idx < len(questions): cats.setdefault(questions[idx].get('type','Other'),[]).append(float(e.get('overall',0))*10)
    cat_scores={c:round(sum(v)/len(v),1) if v else 0 for c,v in cats.items()}
    overall=round(sum(e.get('overall',0) for e in evaluations)/len(evaluations)*10,1) if evaluations else 0
    return overall,cat_scores


def skill_gaps(s, evaluations):
    profile=json.loads(s.profile); jd=json.loads(s.jd_analysis or '{}'); resume=set(profile.get('skills',[])); required=jd.get('jd_skills',[])
    # interview evidence per skill
    skill_scores={}
    questions=json.loads(s.questions); answers=json.loads(s.answers)
    for a,e in zip(answers,evaluations):
        q=questions[a.get('index',0)] if a.get('index',0)<len(questions) else {}
        if q.get('skill'): skill_scores[q['skill']]=max(skill_scores.get(q['skill'],0),float(e.get('overall',0))*10)
    rows=[]
    for skill in required:
        if skill not in resume: status='Missing'; score=0
        elif skill_scores.get(skill,100)>=80: status='Strong'; score=skill_scores.get(skill,100)
        elif skill_scores.get(skill,0)>=60: status='Moderate'; score=skill_scores.get(skill,60)
        else: status='Needs Improvement'; score=skill_scores.get(skill,45)
        priority='High' if skill not in resume else ('Medium' if score<75 else 'Low')
        rows.append({'skill':skill,'status':status,'score':round(score,1),'priority':priority,'why':f'{skill} is relevant to the target JD.'})
    for skill in sorted(set(profile.get('skills',[])) - set(required)):
        if skill_scores.get(skill,100)<60:
            rows.append({'skill':skill,'status':'Needs Improvement','score':skill_scores.get(skill,50),'priority':'Medium','why':'Interview evidence suggests additional practice would help.'})
    return sorted(rows,key=lambda x:({'High':0,'Medium':1,'Low':2}[x['priority']],x['score']))


def owned(sid):
    s=db.session.get(InterviewSession,sid)
    return s if s and s.user_id==current_user.id else None
