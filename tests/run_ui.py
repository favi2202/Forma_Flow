"""Run DOM integration against a temporary real server; no browser installation."""
import os
from pathlib import Path
import socket
import subprocess
import sys
import tempfile
import time
from urllib.request import urlopen

root = Path(__file__).resolve().parents[1]
with socket.socket() as sock:
    sock.bind(('127.0.0.1', 0))
    port = sock.getsockname()[1]
base = f'http://127.0.0.1:{port}'
with tempfile.TemporaryFile() as log:
    server = subprocess.Popen([sys.executable, '-m', 'uvicorn', 'app:app', '--host', '127.0.0.1', '--port', str(port)], cwd=root, stdout=log, stderr=log)
    try:
        for _ in range(100):
            try:
                with urlopen(base + '/health', timeout=1) as response:
                    if response.status == 200:
                        break
            except OSError:
                time.sleep(0.1)
        else:
            raise RuntimeError('Test server did not start')
        env = {**os.environ, 'FORMAFLOW_TEST_URL': base}
        result = subprocess.run(['node', 'tests/ui_roster.cjs'], cwd=root, env=env)
        sys.exit(result.returncode)
    finally:
        server.terminate()
        server.wait(timeout=10)
