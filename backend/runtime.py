import os, json, re, random, math, io
from datetime import datetime, timezone, timedelta

from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager
from dotenv import load_dotenv

load_dotenv()

# Project root is one level above this package.
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
IST = timezone(timedelta(hours=5, minutes=30))

def now_ist():
    return datetime.now(IST).replace(tzinfo=None)

app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, 'templates'),
    static_folder=os.path.join(BASE_DIR, 'static')
)
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY') or os.urandom(32).hex()
app.config['MAX_CONTENT_LENGTH'] = 8 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = os.path.join(BASE_DIR, 'static', 'uploads')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(os.path.join(BASE_DIR, 'instance'), exist_ok=True)
app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('DATABASE_URL', 'sqlite:///' + os.path.join(BASE_DIR, 'instance', 'resume2interview.db'))
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'
app.jinja_env.globals['enumerate'] = enumerate
app.jinja_env.globals['zip'] = zip

GROQ_API_KEY = os.getenv('GROQ_API_KEY', '').strip()
GROQ_MODEL = os.getenv('GROQ_MODEL', 'llama-3.3-70b-versatile').strip()
GROQ_BASE = os.getenv('GROQ_BASE_URL', 'https://api.groq.com/openai/v1').strip()
GROQ_TRANSCRIPTION_MODEL = os.getenv('GROQ_TRANSCRIPTION_MODEL', 'whisper-large-v3-turbo').strip()
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'txt'}

SKILLS = [
    'python','java','javascript','typescript','c++','c#','react','angular','vue','node.js','nodejs','django','flask','fastapi','spring',
    'html','css','sql','mysql','postgresql','mongodb','redis','docker','kubernetes','git','github','aws','azure','gcp','linux',
    'machine learning','deep learning','tensorflow','pytorch','scikit-learn','pandas','numpy','rest api','graphql','agile','scrum',
    'tableau','data science','nlp','devops','ci/cd','kotlin','swift','flutter','dart','express.js','express','next.js','tailwind','bootstrap',
    'firebase','hadoop','spark','dbms','data structures','oops','object oriented programming','operating systems','cloud computing','api'
]

CATEGORIES = ['Technical','HR','Behavioural','Projects','Certifications']
CATEGORY_COUNTS = {'Technical': 10, 'HR': 5, 'Behavioural': 5, 'Projects': 6, 'Certifications': 4}
