import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom';
import { ThemeProvider } from '@/components/ThemeProvider';
import { Toaster } from '@/components/ui/sonner';
import { AuthProvider, useAuth } from '@/contexts/AuthContext';
import { FullPageLoader } from '@/components/ui/loader';
import LoginPage from '@/pages/LoginPage';
import ForgotPassword from '@/pages/ForgotPassword';
import DashboardHome from '@/pages/DashboardHome';
import LineLosses from '@/pages/LineLosses';
import EnergyConsumption from '@/pages/EnergyConsumption';
import MaxMinData from '@/pages/MaxMinData';
import Reports from '@/pages/Reports';
import DashboardLayout from '@/components/DashboardLayout';
import '@/App.css';

function ProtectedRoute({ children }) {
  const { user, loading } = useAuth();
  
  if (loading) {
    return <FullPageLoader text="Authenticating..." />;
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
            <Route path="/forgot-password" element={<ForgotPassword />} />
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
            <Route path="/max-min-data" element={
              <ProtectedRoute>
                <DashboardLayout>
                  <MaxMinData />
                </DashboardLayout>
              </ProtectedRoute>
            } />
            <Route path="/reports" element={
              <ProtectedRoute>
                <DashboardLayout>
                  <Reports />
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
