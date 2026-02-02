import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom';
import { ThemeProvider } from '@/components/ThemeProvider';
import { Toaster } from '@/components/ui/sonner';
import { AuthProvider, useAuth } from '@/contexts/AuthContext';
import LoginPage from '@/pages/LoginPage';
import DashboardHome from '@/pages/DashboardHome';
import LineLosses from '@/pages/LineLosses';
import EnergyConsumption from '@/pages/EnergyConsumption';
import DashboardLayout from '@/components/DashboardLayout';
import '@/App.css';

function ProtectedRoute({ children }) {
  const { user, loading } = useAuth();
  
  if (loading) {
    return <div className="flex items-center justify-center h-screen">Loading...</div>;
  }
  
  if (!user) {
    return <Navigate to="/login" replace />;
  }
  
  return children;
}

function App() {
  return (
    <ThemeProvider defaultTheme="light" storageKey="mis-portal-theme">
      <AuthProvider>
        <BrowserRouter>
          <Routes>
            <Route path="/login" element={<LoginPage />} />
            <Route path="/" element={
              <ProtectedRoute>
                <DashboardLayout>
                  <DashboardHome />
                </DashboardLayout>
              </ProtectedRoute>
            } />
            <Route path="/line-losses" element={
              <ProtectedRoute>
                <DashboardLayout>
                  <LineLosses />
                </DashboardLayout>
              </ProtectedRoute>
            } />
            <Route path="/energy-consumption" element={
              <ProtectedRoute>
                <DashboardLayout>
                  <EnergyConsumption />
                </DashboardLayout>
              </ProtectedRoute>
            } />
          </Routes>
          <Toaster richColors position="top-right" />
        </BrowserRouter>
      </AuthProvider>
    </ThemeProvider>
  );
}

export default App;
