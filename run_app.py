
import sys
import os
import uvicorn
import multiprocessing
import webview
import threading
import time
import logging
import base64
import requests
from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.concurrency import run_in_threadpool

# Setup file logging for debugging frozen app
log_file = os.path.join(os.path.expanduser("~"), "mis_portal_debug.log")
logging.basicConfig(
    filename=log_file,
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger("MIS_PORTAL")

# Fix for "AttributeError: 'NoneType' object has no attribute 'isatty'"
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

# Essential for PyInstaller with multiprocessing (if uvicorn uses it)
multiprocessing.freeze_support()

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
        logger.info(f"Running in frozen mode. MEIPASS: {base_path}")
    except Exception:
        base_path = os.path.abspath(".")
        logger.info(f"Running in dev mode. Base path: {base_path}")

    return os.path.join(base_path, relative_path)

app = FastAPI()

# Proxy configuration
TARGET_URL = "https://mis-portal-production.up.railway.app"

@app.api_route("/api/{path_name:path}", methods=["GET", "POST", "PUT", "DELETE", "OPTIONS", "HEAD", "PATCH"])
async def proxy_api(path_name: str, request: Request):
    """
    Proxy /api requests to the production backend to avoid CORS issues in the frozen app.
    """
    url = f"{TARGET_URL}/api/{path_name}"
    
    # Read body
    body = await request.body()
    
    # Filter headers (host and content-length are handled by the library/network)
    headers = {key: value for key, value in request.headers.items() 
               if key.lower() not in ['host', 'content-length']}
    
    def make_request():
        return requests.request(
            method=request.method,
            url=url,
            headers=headers,
            data=body,
            params=request.query_params,
            stream=True
        )
    
    try:
        # Run synchronous request in threadpool to avoid blocking event loop
        r = await run_in_threadpool(make_request)
        
        return StreamingResponse(
            r.iter_content(chunk_size=4096),
            status_code=r.status_code,
            headers=dict(r.headers),
            media_type=r.headers.get("content-type")
        )
    except Exception as e:
        logger.error(f"Proxy error: {e}")
        return {"error": str(e)}

# 1. Determine Build Directory
build_dir = resource_path('frontend_build')
logger.info(f"Build directory determined as: {build_dir}")

if not os.path.exists(build_dir):
    logger.warning(f"Build directory not found at {build_dir}. Trying fallback...")
    build_dir = os.path.join(os.path.abspath("."), "frontend", "build")
    logger.info(f"Fallback build directory: {build_dir}")

if os.path.exists(build_dir):
    logger.info(f"Build directory exists: {build_dir}")
    logger.info(f"Contents of build dir: {os.listdir(build_dir)}")
else:
    logger.error(f"CRITICAL: Build directory does not exist!")

# 2. Mount Static Files
static_dir = os.path.join(build_dir, "static")
if os.path.exists(static_dir):
    app.mount("/static", StaticFiles(directory=static_dir), name="static")
    logger.info(f"Mounted static files from {static_dir}")
else:
    logger.warning(f"Static directory not found at {static_dir}")

# 3. Serve React App
@app.get("/")
async def serve_index():
    """Explicitly serve index.html at root to avoid 404"""
    logger.info("Serving root route /")
    index_file = os.path.join(build_dir, "index.html")
    if os.path.exists(index_file):
        logger.info(f"Serving index.html from {index_file}")
        return FileResponse(index_file)
    
    logger.error(f"Index file missing at {index_file}")
    return HTMLResponse(content=f"<h1>Error: index.html not found</h1><p>Checked: {index_file}</p>", status_code=404)

@app.get("/{full_path:path}")
async def serve_react_app(full_path: str):
    """Serve React App or Static files"""
    logger.info(f"Request for: {full_path}")
    
    # Try to find the file in the build directory
    target_file = os.path.join(build_dir, full_path)
    
    if full_path and os.path.exists(target_file) and os.path.isfile(target_file):
        logger.info(f"Serving file: {target_file}")
        return FileResponse(target_file)
    
    # If not found, serve index.html (for client-side routing)
    index_file = os.path.join(build_dir, "index.html")
    if os.path.exists(index_file):
        logger.info(f"Serving index.html (SPA fallback) for {full_path}")
        return FileResponse(index_file)
    
    logger.error(f"404 Not Found: {full_path}")
    return HTMLResponse(content=f"<h1>404 Not Found</h1><p>Path: {full_path}</p><p>Build Dir: {build_dir}</p>", status_code=404)


class Api:
    def save_file(self, filename, content_base64):
        try:
            # Use pywebview's save dialog
            file_types = ('Excel Files (*.xlsx)', 'All files (*.*)')
            save_path = webview.windows[0].create_file_dialog(
                webview.SAVE_DIALOG, 
                directory='', 
                save_filename=filename,
                file_types=file_types
            )
            
            if save_path:
                if isinstance(save_path, (list, tuple)):
                    if not save_path: return {"success": False, "reason": "cancelled"}
                    save_path = save_path[0]
                
                if save_path:
                    # content_base64 might have header "data:application/vnd...;base64,"
                    if ',' in content_base64:
                        content_base64 = content_base64.split(',')[1]
                    
                    data = base64.b64decode(content_base64)
                    with open(save_path, 'wb') as f:
                        f.write(data)
                    return {"success": True, "path": save_path}
            return {"success": False, "reason": "cancelled"}
        except Exception as e:
            logger.error(f"Save file failed: {e}")
            return {"success": False, "reason": str(e)}

def start_server():
    # Run Uvicorn on a specific port
    logger.info("Starting Uvicorn server...")
    try:
        uvicorn.run(app, host="127.0.0.1", port=8000, log_level="debug")
    except Exception as e:
        logger.error(f"Uvicorn failed: {e}")

if __name__ == "__main__":
    logger.info("Application starting...")
    # Start the server in a separate thread
    t = threading.Thread(target=start_server)
    t.daemon = True
    t.start()

    # Give the server a moment to start
    time.sleep(2)

    # Start the webview
    logger.info("Starting Webview...")
    api = Api()
    try:
        webview.create_window(
            'MIS PORTAL', 
            'http://127.0.0.1:8000',
            width=1200,
            height=800,
            resizable=True,
            js_api=api
        )
        webview.start()
    except Exception as e:
        logger.error(f"Webview failed: {e}")


