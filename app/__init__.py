from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import text

db = SQLAlchemy()

def _ensure_runtime_schema():
    stmts = [
        "ALTER TABLE users ADD COLUMN full_name VARCHAR(150) NULL",
        """
        CREATE TABLE IF NOT EXISTS user_pt_accesses (
            id INT NOT NULL AUTO_INCREMENT,
            user_id INT NOT NULL,
            pt_name VARCHAR(150) NOT NULL,
            created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (id),
            UNIQUE KEY uq_user_pt_accesses_user_pt (user_id, pt_name),
            KEY idx_user_pt_accesses_user_id (user_id),
            KEY idx_user_pt_accesses_pt_name (pt_name),
            CONSTRAINT fk_user_pt_accesses_user FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
        """,
    ]
    for stmt in stmts:
        try:
            db.session.execute(text(stmt))
            db.session.commit()
        except Exception:
            db.session.rollback()


def create_app():
    app = Flask(__name__)

    app.config["SECRET_KEY"] = "change-me"
    app.config["SQLALCHEMY_DATABASE_URI"] = "mysql+pymysql://root:Fleet%40123@localhost/fleet_sinar_group"
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

    db.init_app(app)

    from .routes import bp
    app.register_blueprint(bp)

    with app.app_context():
        db.create_all()
        _ensure_runtime_schema()

    return app
