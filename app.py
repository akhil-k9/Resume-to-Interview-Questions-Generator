import os
from backend.runtime import app, db, login_manager
from backend.models import User, InterviewSession
from backend.services import load_user

login_manager.user_loader(load_user)

# Importing routes registers all existing Flask routes and handlers.
from backend import routes  # noqa: F401

with app.app_context():
    db.create_all()

if __name__ == '__main__':
    app.run(
        debug=os.getenv('FLASK_DEBUG', '0') == '1',
        port=int(os.getenv('PORT', 5000)),
        host=os.getenv('HOST', '127.0.0.1')
    )
