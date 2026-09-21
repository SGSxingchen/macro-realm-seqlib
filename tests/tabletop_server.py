"""Test-only host: production FastAPI app + built SPA, no mocked tabletop API."""
from pathlib import Path
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from web.backend.app.main import app

DIST = Path(__file__).resolve().parents[1] / 'web' / 'frontend' / 'dist'
app.mount('/assets', StaticFiles(directory=DIST / 'assets'), name='test-assets')
@app.get('/')

def frontend():
    return FileResponse(DIST / 'index.html')

@app.get('/tabletop.html')
def tabletop_frontend():
    return FileResponse(DIST / 'tabletop.html')
