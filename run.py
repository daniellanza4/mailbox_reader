
import webbrowser

from app import create_app
from app.services.mail_service import get_access_token

app = create_app()

if __name__ == "__main__":
    print("Autenticando con Microsoft...", flush=True)

    try:
        get_access_token()
        print("Autenticación completada.", flush=True)

        webbrowser.open("http://127.0.0.1:5000/")
        app.run(debug=False)

    except Exception as exc:
        print(f"Error al iniciar Mailbox Reader: {exc}", flush=True)
