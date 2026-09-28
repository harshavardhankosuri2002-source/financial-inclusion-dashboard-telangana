import sys
import os
import urllib.parse

# Ensure root directory is on Python path so all modules and Excel file are accessible
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

from app import app, dash_app

class VercelPathFixMiddleware:
    """
    Vercel Serverless Function WSGI middleware.
    Restores the real client request path from the __dash_path parameter
    forwarded by vercel.json rewrite rules.
    """
    def __init__(self, wsgi_app):
        self.wsgi_app = wsgi_app

    def __call__(self, environ, start_response):
        qs = environ.get("QUERY_STRING", "")
        if "__dash_path=" in qs:
            params = urllib.parse.parse_qs(qs)
            if "__dash_path" in params:
                real_path = params["__dash_path"][0]
                if not real_path.startswith("/"):
                    real_path = "/" + real_path
                environ["PATH_INFO"] = real_path
                del params["__dash_path"]
                environ["QUERY_STRING"] = urllib.parse.urlencode(params, doseq=True)

        return self.wsgi_app(environ, start_response)

app.wsgi_app = VercelPathFixMiddleware(app.wsgi_app)

server = app
handler = app
application = app

if __name__ == "__main__":
    dash_app.run(debug=False, host="127.0.0.1", port=8050)
