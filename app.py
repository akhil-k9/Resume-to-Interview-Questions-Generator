
import os

from backend.runtime import app, db, login_manager
from backend.models import User, InterviewSession
from backend.services import load_user



login_manager.user_loader(load_user)

from backend import routes  


with app.app_context():
    db.create_all()


if __name__ == '__main__':
    app.run(
        debug=os.getenv('FLASK_DEBUG', '0') == '1',
        host='0.0.0.0',
        port=int(os.getenv('PORT', '10000'))
    )