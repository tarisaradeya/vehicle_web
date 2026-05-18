import os
from app import create_app

app = create_app()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("Fleet & Asset Vehicle System")
    print("Database sudah tersedia, sistem siap dijalankan")
    print(f"Server berjalan di http://0.0.0.0:{port}")

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False,
        use_reloader=False
    )