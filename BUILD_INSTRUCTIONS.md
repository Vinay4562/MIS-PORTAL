
# Build Instructions for Windows Installer (.exe)

This project is set up to generate a standalone Windows executable that serves the React frontend as a desktop application, along with a Setup installer.

## Prerequisites

*   **Python 3.8+**:
    *   **Important**: To generate an executable compatible with **both 32-bit and 64-bit systems**, you must use a **32-bit version of Python**.
    *   If you use 64-bit Python, the output `.exe` will only work on 64-bit Windows.
*   **Node.js** and **npm**
*   **Python Dependencies**:
    ```bash
    pip install pyinstaller pillow uvicorn fastapi pywebview pywin32 winshell
    ```

## Steps to Build

1.  **Build the Frontend**:
    Navigate to the frontend directory and build the React application.
    ```bash
    cd frontend
    npm install
    npm run build
    cd ..
    ```
    *Note: Ensure `frontend/.env` contains the correct `REACT_APP_BACKEND_URL` (e.g., your production URL) before building.*

2.  **Ensure Icon Exists**:
    The build script expects `frontend/public/favicon.ico`. If it's missing, you can convert the PNG version:
    ```python
    from PIL import Image
    Image.open('frontend/public/favicon.png').save('frontend/public/favicon.ico')
    ```

3.  **Build the Main Application**:
    Run PyInstaller to bundle the application logic (Server + Frontend + Webview).
    
    ```bash
    pyinstaller --noconfirm --onefile --windowed --icon "frontend/public/favicon.ico" --add-data "frontend/build;frontend_build" --name "MIS_PORTAL" run_app.py
    ```

    *   `--onefile`: Create a single `.exe` file.
    *   `--windowed`: Run without a console window (GUI mode).
    *   `--add-data`: Bundle the frontend build folder.
    *   **Output**: `dist/MIS_PORTAL.exe`

4.  **Build the Installer (Setup)**:
    Run PyInstaller to bundle the `MIS_PORTAL.exe` into a setup wizard script.
    
    ```bash
    pyinstaller --noconfirm --onefile --console --name "MIS_PORTAL_Setup" --add-data "dist/MIS_PORTAL.exe;." install_script.py
    ```
    
    *   This creates `dist/MIS_PORTAL_Setup.exe`.
    *   When run, this Setup file will:
        1.  Extract the inner `MIS_PORTAL.exe`.
        2.  Copy it to `%LOCALAPPDATA%\MIS_PORTAL\`.
        3.  Create a Desktop Shortcut "MIS_PORTAL".

## How to Distribute
*   Share **`dist/MIS_PORTAL_Setup.exe`** with users.
*   Users simply run this file to install the app and get a desktop shortcut.
*   The installed app runs as a standalone window (not in the browser) and connects to the backend URL configured in step 1.

## Debugging
*   The application creates a log file at `C:\Users\<User>\mis_portal_debug.log`.
*   Check this log if the application fails to start or load content.
