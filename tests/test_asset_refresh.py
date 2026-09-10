"""Guard against new HTML loading cached code without its new controls."""
import hashlib
import re
from pathlib import Path

from fastapi.testclient import TestClient

import app as module


client = TestClient(module.app)


def test_home_references_the_current_assets_and_disables_page_caching():
    response = client.get('/')
    assert response.status_code == 200
    assert response.headers['cache-control'] == 'no-store'
    for filename in ('app.js', 'styles.css'):
        content = (module.STATIC_DIR / filename).read_bytes()
        version = hashlib.sha256(content).hexdigest()[:16]
        url = f'/static/{filename}?v={version}'
        assert f'"{url}"' in response.text
        asset = client.get(url)
        assert asset.status_code == 200
        assert asset.content == content
        assert asset.headers['cache-control'] == 'no-cache, must-revalidate'
        # Existing unversioned URLs must also revalidate their cached responses.
        old_url = client.get('/static/' + filename)
        assert old_url.headers['cache-control'] == 'no-cache, must-revalidate'


def test_script_or_style_change_produces_a_new_cache_key(tmp_path, monkeypatch):
    for filename in ('index.html', 'app.js', 'styles.css'):
        (tmp_path / filename).write_bytes((module.STATIC_DIR / filename).read_bytes())
    monkeypatch.setattr(module, 'STATIC_DIR', tmp_path)
    before = client.get('/').text
    for filename in ('app.js', 'styles.css'):
        path = tmp_path / filename
        path.write_bytes(path.read_bytes() + b'\n/* new release */\n')
    after = client.get('/').text
    for filename in ('app.js', 'styles.css'):
        pattern = rf'/static/{re.escape(filename)}\?v=[a-f0-9]{{16}}'
        assert re.search(pattern, before).group() != re.search(pattern, after).group()


def test_manual_panel_does_not_expose_uninitialized_english_controls():
    html = client.get('/').text
    assert re.search(r'id="manualPanel"[^>]*class="[^"]*\bhidden\b', html)
    assert re.search(r'id="addStudent"[^>]*\bdisabled', html)
