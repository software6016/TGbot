"""
Запусти этот файл ОДИН РАЗ чтобы получить REFRESH_TOKEN для бота.
После запуска скопируй токен в bot.py или в переменные окружения сервера.

Запуск: python setup_auth.py
"""

import webbrowser
import urllib.parse
import requests
from http.server import HTTPServer, BaseHTTPRequestHandler

# ── Вставь сюда свои данные из Azure ──────────────────────────────────────────
CLIENT_ID     = "YOUR_CLIENT_ID"      # из Azure Portal
CLIENT_SECRET = "YOUR_CLIENT_SECRET"  # из Azure Portal
REDIRECT_URI  = "http://localhost:8080/callback"
SCOPE         = "https://graph.microsoft.com/Files.Read offline_access"
# ──────────────────────────────────────────────────────────────────────────────

auth_code_holder = {"code": None}


class CallbackHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed.query)
        if "code" in params:
            auth_code_holder["code"] = params["code"][0]
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"<h2>OK! Можно закрыть это окно и вернуться в терминал.</h2>")
        else:
            self.send_response(400)
            self.end_headers()
            self.wfile.write(b"<h2>Ошибка. Попробуй снова.</h2>")

    def log_message(self, *args):
        pass  # отключаем лишние логи


def main():
    # Шаг 1 — открываем браузер для входа в Microsoft
    auth_url = (
        "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
        f"?client_id={CLIENT_ID}"
        f"&response_type=code"
        f"&redirect_uri={urllib.parse.quote(REDIRECT_URI)}"
        f"&scope={urllib.parse.quote(SCOPE)}"
        f"&response_mode=query"
    )
    print("\n🌐 Открываю браузер для входа в Microsoft...")
    webbrowser.open(auth_url)

    # Шаг 2 — ждём callback с кодом
    print("⏳ Ожидаю авторизацию (порт 8080)...")
    server = HTTPServer(("localhost", 8080), CallbackHandler)
    server.handle_request()

    if not auth_code_holder["code"]:
        print("❌ Не получил код авторизации. Попробуй снова.")
        return

    # Шаг 3 — обмениваем код на токены
    token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    data = {
        "grant_type":    "authorization_code",
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "code":          auth_code_holder["code"],
        "redirect_uri":  REDIRECT_URI,
        "scope":         SCOPE,
    }
    r = requests.post(token_url, data=data, timeout=15)
    tokens = r.json()

    if "refresh_token" not in tokens:
        print("❌ Ошибка получения токена:", tokens)
        return

    refresh_token = tokens["refresh_token"]
    print("\n" + "="*60)
    print("✅ УСПЕХ! Скопируй этот REFRESH_TOKEN и вставь в bot.py")
    print("="*60)
    print(f"\nREFRESH_TOKEN = \"{refresh_token}\"\n")
    print("Или задай переменную окружения на сервере:")
    print(f"  ONEDRIVE_REFRESH_TOKEN={refresh_token}\n")


if __name__ == "__main__":
    main()
