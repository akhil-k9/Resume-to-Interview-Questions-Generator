# Backend structure

The original `app.py` was split without changing the application's routes, templates,
database schema, prompts, AI calls, or frontend files.

- `app.py` - compatibility entry point. `python app.py` still works.
- `backend/runtime.py` - Flask app, configuration, database/login extensions, constants.
- `backend/models.py` - SQLAlchemy models.
- `backend/services.py` - resume parsing, profile extraction, AI/Groq helpers,
  question generation, answer generation/evaluation, analytics, and shared helpers.
- `backend/routes.py` - the existing Flask error handlers, context processor, and routes.
- `templates/` and `static/` remain unchanged.

Run exactly as before:
    python app.py

The existing `run.bat` is left unchanged.
