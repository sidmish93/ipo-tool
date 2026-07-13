"""
Production entrypoint. Runs the Flask app under waitress (a real WSGI server).

  PORT     port to bind (default 8080; hosts like Render inject this)
  HOST     interface to bind (default 0.0.0.0 so it's reachable externally)

Local dev without a WSGI server:  python app.py
Production:                        python serve.py
"""

import os
from waitress import serve

from app import app

if __name__ == "__main__":
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "8080"))
    # 2 threads is plenty: generation is serialized to one browser job anyway,
    # the extra thread keeps the UI/status responsive during a run.
    print(f"Serving on http://{host}:{port}")
    serve(app, host=host, port=port, threads=4)
