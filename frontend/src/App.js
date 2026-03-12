import { useEffect, useState } from 'react';
import axios from 'axios';
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
import Interruptions from '@/pages/Interruptions';
import StationLoads from '@/pages/StationLoads';
import Reports from '@/pages/Reports';
import DashboardLayout from '@/components/DashboardLayout';
import AdminLayout from '@/components/admin/AdminLayout';
import AdminLoginPage from '@/pages/admin/AdminLoginPage';
import AdminDashboardHome from '@/pages/admin/AdminDashboardHome';
import AdminEnergyAnalytics from '@/pages/admin/AdminEnergyAnalytics';
import AdminLineLossesAnalytics from '@/pages/admin/AdminLineLossesAnalytics';
import AdminMaxMinAnalytics from '@/pages/admin/AdminMaxMinAnalytics';
import AdminStationLoadAnalytics from '@/pages/admin/AdminStationLoadAnalytics';
import AdminInterruptionsAnalytics from '@/pages/admin/AdminInterruptionsAnalytics';
import AdminBulkImport from '@/pages/admin/AdminBulkImport';
import '@/App.css';

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

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

function AdminProtectedRoute({ children }) {
  const { user, loading } = useAuth();
  const [checking, setChecking] = useState(true);
  const [isAdmin, setIsAdmin] = useState(false);

  useEffect(() => {
    const verifyAdmin = async () => {
      if (loading) return;

      if (!user) {
        setChecking(false);
        setIsAdmin(false);
        return;
      }

      try {
        const token = localStorage.getItem('token');
        await axios.get(`${API}/admin/me`, {
          headers: {
            Authorization: `Bearer ${token}`,
          },
        });
        setIsAdmin(true);
      } catch (error) {
        console.error('Admin access check failed:', error);
        setIsAdmin(false);
      } finally {
        setChecking(false);
      }
    };

    verifyAdmin();
  }, [loading, user]);

  if (loading || checking) {
    return <FullPageLoader text="Verifying admin access..." />;
  }

  if (!user) {
    return <Navigate to="/admin/login" replace />;
  }

  if (!isAdmin) {
    return <Navigate to="/" replace />;
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
            <Route path="/admin/login" element={<AdminLoginPage />} />
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
            <Route path="/interruptions" element={
              <ProtectedRoute>
                <DashboardLayout>
                  <Interruptions />
                </DashboardLayout>
              </ProtectedRoute>
            } />
            <Route path="/station-loads" element={
              <ProtectedRoute>
                <DashboardLayout>
                  <StationLoads />
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
            <Route path="/admin" element={
              <AdminProtectedRoute>
                <AdminLayout>
                  <AdminDashboardHome />
                </AdminLayout>
              </AdminProtectedRoute>
            } />
            <Route path="/admin/energy" element={
              <AdminProtectedRoute>
                <AdminLayout>
                  <AdminEnergyAnalytics />
                </AdminLayout>
              </AdminProtectedRoute>
            } />
            <Route path="/admin/line-losses" element={
              <AdminProtectedRoute>
                <AdminLayout>
                  <AdminLineLossesAnalytics />
                </AdminLayout>
              </AdminProtectedRoute>
            } />
            <Route path="/admin/max-min" element={
              <AdminProtectedRoute>
                <AdminLayout>
                  <AdminMaxMinAnalytics />
                </AdminLayout>
              </AdminProtectedRoute>
            } />
            <Route path="/admin/station-load" element={
              <AdminProtectedRoute>
                <AdminLayout>
                  <AdminStationLoadAnalytics />
                </AdminLayout>
              </AdminProtectedRoute>
            } />
            <Route path="/admin/interruptions" element={
              <AdminProtectedRoute>
                <AdminLayout>
                  <AdminInterruptionsAnalytics />
                </AdminLayout>
              </AdminProtectedRoute>
            } />
            <Route path="/admin/import" element={
              <AdminProtectedRoute>
                <AdminLayout>
                  <AdminBulkImport />
                </AdminLayout>
              </AdminProtectedRoute>
            } />
          </Routes>
          <Toaster richColors position="top-right" />
        </BrowserRouter>
      </AuthProvider>
    </ThemeProvider>
  );
}

export default App;
