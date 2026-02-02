# MIS PORTAL

## Project Overview

MIS PORTAL is a comprehensive management information system designed to track and analyze energy consumption and line losses. It features a modern React frontend and a robust FastAPI backend, providing real-time data visualization and management capabilities.

### Key Features
- **Line Losses Tracking**: Monitor and analyze transmission losses across various feeders.
- **Energy Consumption**: Track energy usage with detailed monthly and yearly reports.
- **Interactive Dashboard**: User-friendly interface with dark/light mode support.
- **Secure Authentication**: JWT-based authentication system.

## Tech Stack
- **Frontend**: React, Tailwind CSS, Shadcn UI
- **Backend**: Python, FastAPI, MongoDB
- **Deployment**: Vercel (Frontend), Railway (Backend)

## Running Locally

### Prerequisites
- Node.js (v18+)
- Python (v3.9+)
- MongoDB connection string

### 1. Backend Setup

1.  Navigate to the backend directory:
    ```bash
    cd backend
    ```
2.  Create a virtual environment:
    ```bash
    python -m venv .venv
    ```
3.  Activate the virtual environment:
    - Windows: `.venv\Scripts\activate`
    - Mac/Linux: `source .venv/bin/activate`
4.  Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```
5.  Create a `.env` file based on `.env.example` and add your MongoDB URL:
    ```env
    MONGO_URL=your_mongodb_connection_string
    DB_NAME=mis_portal
    JWT_SECRET_KEY=your_secret_key
    ```
6.  Start the server:
    ```bash
    uvicorn server:app --reload
    ```
    The API will be available at `http://localhost:8000`.

### 2. Frontend Setup

1.  Navigate to the frontend directory:
    ```bash
    cd frontend
    ```
2.  Install dependencies:
    ```bash
    yarn install
    ```
3.  Create a `.env` file:
    ```env
    REACT_APP_BACKEND_URL=http://localhost:8000
    ENABLE_HEALTH_CHECK=false
    ```
4.  Start the development server:
    ```bash
    yarn start
    ```
    The application will run at `http://localhost:3000`.

## Deployment

Refer to [DEPLOYMENT.md](DEPLOYMENT.md) for detailed deployment instructions on Railway and Vercel.

## License
© 2026 VinTech Solutions. All rights reserved.
