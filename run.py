from app import create_app

app = create_app()

if __name__ == "__main__":
    print("Fleet & Asset Vehicle System")
    print("Database sudah tersedia, sistem siap dijalankan")
    print("Server berjalan di http://127.0.0.1:5000")

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True,
        use_reloader=False
    )