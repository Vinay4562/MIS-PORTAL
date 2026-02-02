# Deployment Guide

This project is configured for deployment with **Railway** (Backend) and **Vercel** (Frontend).

## 1. Backend Deployment (Railway)

The backend is a FastAPI application using MongoDB.

### Steps:
1.  **Push to GitHub**: Ensure your latest code is pushed to your GitHub repository.
2.  **Login to Railway**: Go to [railway.app](https://railway.app/) and login with GitHub.
3.  **New Project**: Click "New Project" -> "Deploy from GitHub repo" -> Select this repository.
4.  **Configure Service**:
    *   Railway should automatically detect the `backend` directory if you select it as the Root Directory, or you can configure it manually.
    *   Go to **Settings** -> **Root Directory**: Set to `backend`.
5.  **Environment Variables**:
    *   Go to the **Variables** tab.
    *   Add the following variables:
        *   `MONGO_URL`: Your MongoDB connection string (e.g., from MongoDB Atlas).
        *   `DB_NAME`: `mis_portal` (or your preferred DB name).
        *   `JWT_SECRET_KEY`: A strong random string for security.
        *   `CORS_ORIGINS`: Initially `*`, or update to your Vercel frontend URL once deployed (e.g., `https://your-frontend.vercel.app`).
6.  **Deploy**: Railway will automatically build and deploy.
    *   It uses the `Procfile` included in the `backend` folder: `web: uvicorn server:app --host 0.0.0.0 --port $PORT`.
7.  **Get URL**: Once deployed, copy the **Public Networking URL** (e.g., `https://backend-production.up.railway.app`).

## 2. Frontend Deployment (Vercel)

The frontend is a React application using CRA (via Craco).

### Steps:
1.  **Login to Vercel**: Go to [vercel.com](https://vercel.com/) and login with GitHub.
2.  **Add New Project**: Click "Add New..." -> "Project" -> Import this repository.
3.  **Configure Project**:
    *   **Root Directory**: Click "Edit" and select `frontend`.
    *   **Framework Preset**: Vercel should auto-detect "Create React App".
    *   **Build Command**: `yarn build` (default).
    *   **Output Directory**: `build` (default).
4.  **Environment Variables**:
    *   Expand the **Environment Variables** section.
    *   Add `REACT_APP_BACKEND_URL`.
    *   **Value**: Paste the Railway Backend URL from step 1 (e.g., `https://backend-production.up.railway.app`). **Important**: Do not add a trailing slash `/`.
5.  **Deploy**: Click "Deploy".
6.  **Verify**: Vercel will build and assign a domain (e.g., `your-project.vercel.app`).

## 3. Post-Deployment

1.  **Update CORS**:
    *   Go back to Railway -> Variables.
    *   Update `CORS_ORIGINS` to your new Vercel URL (e.g., `https://your-project.vercel.app`).
    *   Railway will automatically redeploy.

2.  **Test**:
    *   Open your Vercel URL.
    *   Try logging in and fetching data to ensure the frontend can talk to the backend.

## Troubleshooting

*   **Frontend 404s on Refresh**: A `vercel.json` has been added to `frontend/` to handle client-side routing rewrites.
*   **Backend Connection Error**: Check the Console in your browser's Developer Tools.
    *   If you see "CORS error", check `CORS_ORIGINS` in Railway.
    *   If you see "404 Not Found" for API calls, check `REACT_APP_BACKEND_URL` in Vercel (ensure no trailing slash).
