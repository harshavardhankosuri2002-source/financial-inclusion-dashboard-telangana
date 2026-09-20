import sys
import os

# Ensure root directory is on Python path so all modules and Excel file are accessible
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

from app import app, dash_app

# Official Vercel entrypoint: Vercel expects a Flask instance named `app`
server = app
handler = app
application = app

if __name__ == "__main__":
    dash_app.run(debug=False, host="127.0.0.1", port=8050)
