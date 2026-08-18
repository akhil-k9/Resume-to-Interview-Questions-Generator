from flask_login import UserMixin
from .runtime import db, now_ist

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(180), unique=True, nullable=False)
    password = db.Column(db.String(300), nullable=False)
    created_at = db.Column(db.DateTime, default=now_ist)
    sessions = db.relationship('InterviewSession', backref='user', lazy=True, cascade='all,delete-orphan')

class InterviewSession(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    job_role = db.Column(db.String(200), nullable=False)
    job_desc = db.Column(db.Text, default='')
    resume_filename = db.Column(db.String(300), default='')
    resume_text = db.Column(db.Text, default='')
    profile = db.Column(db.Text, default='{}')
    jd_analysis = db.Column(db.Text, default='{}')
    questions = db.Column(db.Text, default='[]')
    answers = db.Column(db.Text, default='[]')
    evaluations = db.Column(db.Text, default='[]')
    integrity = db.Column(db.Text, default='{}')
    status = db.Column(db.String(40), default='prepared')
    readiness_score = db.Column(db.Float, default=0)
    created_at = db.Column(db.DateTime, default=now_ist)
