from app import app
from waitress import serve
import webbrowser
from threading import Timer
import socket

def find_free_port(default_port=5001):
    """Cari port kosong, mulai dari default_port."""
    port = default_port
    while True:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("0.0.0.0", port))
                return port
            except OSError:
                port += 1  # Jika sudah dipakai, lanjut ke port berikutnya

if __name__ == "__main__":
    port = find_free_port(5001)
    print(f"Server berjalan di http://localhost:{port}")
    
    # Buka browser otomatis 1 detik setelah exe dijalankan
    Timer(1, lambda: webbrowser.open(f"http://localhost:{port}")).start()
    
    # Jalankan server dengan Waitress
    serve(app, host="0.0.0.0", port=port)
