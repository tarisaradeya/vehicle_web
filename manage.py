import sys
from app import create_app, db

app = create_app()

def init_db():
    db.create_all()
    print("Database berhasil dibuat.")

def reset_db():
    db.drop_all()
    db.create_all()
    print("Database berhasil di-reset.")

def main():
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python manage.py initdb")
        print("  python manage.py resetdb")
        sys.exit(1)

    cmd = sys.argv[1].lower()

    with app.app_context():
        if cmd == "initdb":
            init_db()
        elif cmd == "resetdb":
            reset_db()
        else:
            print("Unknown command:", cmd)
            sys.exit(1)

if __name__ == "__main__":
    main()