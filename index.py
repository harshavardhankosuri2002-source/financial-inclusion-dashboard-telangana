import sys
import os

from app import app, dash_app

# Official Vercel entrypoint: Vercel expects a Flask instance named `app`
server = app
handler = app
application = app

if __name__ == "__main__":
    dash_app.run(debug=False, host="127.0.0.1", port=8050)
