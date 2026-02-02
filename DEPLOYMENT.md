# MIS PORTAL Deployment Guide

This guide outlines the steps to deploy the MIS PORTAL application. The backend is configured for [Railway](https://railway.app/) and the frontend for [Vercel](https://vercel.com/).

## Prerequisites

- GitHub account with the repository pushed
- Railway account
- Vercel account
- MongoDB Atlas connection string (or other MongoDB provider)

## 1. Backend Deployment (Railway)

1.  **Login to Railway**: Go to [railway.app](https://railway.app/) and login.
2.  **New Project**: Click "New Project" > "Deploy from GitHub repo".
3.  **Select Repository**: Choose the `MIS-PORTAL` repository.
4.  **Configure Variables**:
    - Click on the newly created service.
    - Go to the "Variables" tab.
    - Add the following environment variables (using values from your `backend/.env`):
        - `MONGO_URL`: Your MongoDB connection string.
        - `DB_NAME`: `mis_portal`
        - `JWT_SECRET_KEY`: A strong secret key.
5.  **Settings**:
    - Go to the "Settings" tab.
    - Under "Root Directory", set it to `/backend`.
    - Under "Build", Railway usually auto-detects Python. Ensure the "Build Command" is empty (unless needed) and "Start Command" is:
      ```bash
      uvicorn server:app --host 0.0.0.0 --port $PORT
      ```
      (Railway automatically sets `$PORT`).
6.  **Deploy**: Railway will trigger a deployment. Once active, go to "Settings" > "Networking" and generate a domain (e.g., `mis-portal-backend.up.railway.app`).
7.  **Copy URL**: Copy this URL. You will need it for the frontend.

## 2. Frontend Deployment (Vercel)

1.  **Login to Vercel**: Go to [vercel.com](https://vercel.com/) and login.
2.  **Add New**: Click "Add New..." > "Project".
3.  **Import Git Repository**: Find `MIS-PORTAL` and click "Import".
4.  **Configure Project**:
    - **Framework Preset**: Create React App (should be auto-detected).
    - **Root Directory**: Click "Edit" and select `frontend`.
5.  **Environment Variables**:
    - Expand "Environment Variables".
    - Add:
        - `REACT_APP_BACKEND_URL`: The URL of your deployed Railway backend (e.g., `https://mis-portal-production.up.railway.app`). **Important**: Do not include `/api` at the end if your code appends it, but based on current code `const API = ${process.env.REACT_APP_BACKEND_URL}/api`, so just the base URL is needed.
        - `ENABLE_HEALTH_CHECK`: `false` (or `true` if you want it).
6.  **Deploy**: Click "Deploy".
7.  **Verify**: Once deployed, visit the Vercel URL.

## 3. Post-Deployment Checks

- **Backend**: Visit `https://<your-backend-url>/docs` to see the Swagger UI and verify the server is running.
- **Frontend**: Log in to the application and verify data is loading from the backend.

## Troubleshooting

- **CORS Issues**: If the frontend cannot talk to the backend, check `backend/server.py` to ensure `CORSMiddleware` includes your Vercel domain.
  - Currently it is set to `allow_origins=["*"]`, which is open but you may want to restrict it to your specific Vercel domain in production.
- **Build Errors**: Check the build logs in Railway/Vercel for missing dependencies.
- **Vercel Build Fails on Warnings**: If the build fails due to warnings (e.g., peer dependency warnings or ESLint warnings), add an Environment Variable in Vercel:
  - Key: `CI`
  - Value: `false`
  This prevents the build from failing when warnings are present.
