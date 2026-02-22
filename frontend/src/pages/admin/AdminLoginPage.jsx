import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import axios from 'axios';
import { useAuth } from '@/contexts/AuthContext';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Loader } from '@/components/ui/loader';
import { toast } from 'sonner';
import { Zap, ShieldCheck, Eye, EyeOff } from 'lucide-react';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter } from '@/components/ui/dialog';

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

export default function AdminLoginPage() {
  const navigate = useNavigate();
  const { adminLogin } = useAuth();
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [loading, setLoading] = useState(false);
  const [forgotOpen, setForgotOpen] = useState(false);
  const [forgotStep, setForgotStep] = useState(1);
  const [forgotEmail, setForgotEmail] = useState('');
  const [otp, setOtp] = useState('');
  const [newPassword, setNewPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [resetToken, setResetToken] = useState('');
  const [forgotSubmitting, setForgotSubmitting] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setLoading(true);
    try {
      await adminLogin(email, password);
      await axios.get(`${API}/admin/me`);
      toast.success('Admin login successful');
      navigate('/admin');
    } catch (error) {
      const status = error.response?.status;
      if (status === 403) {
        toast.error('This account is not authorised for Admin access');
      } else if (status === 401) {
        toast.error('Invalid credentials');
      } else {
        toast.error(error.response?.data?.detail || 'Failed to authenticate admin');
      }
    } finally {
      setLoading(false);
    }
  };

  const resetForgotState = () => {
    setForgotStep(1);
    setForgotEmail('');
    setOtp('');
    setNewPassword('');
    setConfirmPassword('');
    setResetToken('');
  };

  const openForgot = () => {
    setForgotEmail(email || '');
    setForgotOpen(true);
  };

  const handleRequestOtp = async () => {
    if (!forgotEmail.trim()) {
      toast.error('Enter admin email to request OTP');
      return;
    }
    setForgotSubmitting(true);
    try {
      await axios.post(`${API}/admin/auth/forgot-password`, {
        email: forgotEmail.trim(),
      });
      toast.success('If this admin account exists, an OTP has been sent');
      setForgotStep(2);
    } catch (error) {
      const detail = error.response?.data?.detail;
      if (detail) {
        toast.error(detail);
      } else {
        toast.error('Failed to request OTP');
      }
    } finally {
      setForgotSubmitting(false);
    }
  };

  const handleVerifyOtp = async () => {
    if (!otp.trim()) {
      toast.error('Enter the OTP sent to your email');
      return;
    }
    setForgotSubmitting(true);
    try {
      const resp = await axios.post(`${API}/admin/auth/verify-otp`, {
        email: forgotEmail.trim(),
        otp: otp.trim(),
      });
      const token = resp.data?.reset_token;
      if (!token) {
        toast.error('Failed to verify OTP');
        return;
      }
      setResetToken(token);
      toast.success('OTP verified. You can now set a new password');
      setForgotStep(3);
    } catch (error) {
      const detail = error.response?.data?.detail;
      if (detail) {
        toast.error(detail);
      } else {
        toast.error('Failed to verify OTP');
      }
    } finally {
      setForgotSubmitting(false);
    }
  };

  const handleResetPassword = async () => {
    if (!newPassword || !confirmPassword) {
      toast.error('Enter and confirm the new password');
      return;
    }
    if (newPassword !== confirmPassword) {
      toast.error('New password and confirmation do not match');
      return;
    }
    if (newPassword.length < 8) {
      toast.error('New password must be at least 8 characters long');
      return;
    }
    setForgotSubmitting(true);
    try {
      await axios.post(`${API}/admin/auth/reset-password`, {
        email: forgotEmail.trim(),
        reset_token: resetToken,
        new_password: newPassword,
      });
      toast.success('Admin password reset successful. You can now sign in');
      setForgotOpen(false);
      resetForgotState();
      setEmail(forgotEmail.trim());
      setPassword('');
    } catch (error) {
      const detail = error.response?.data?.detail;
      if (detail) {
        toast.error(detail);
      } else {
        toast.error('Failed to reset password');
      }
    } finally {
      setForgotSubmitting(false);
    }
  };

  return (
    <div className="min-h-screen flex flex-col items-center justify-center bg-gradient-to-br from-slate-950 via-slate-900 to-slate-800 px-4">
      <div className="max-w-md w-full space-y-8">
        <div className="text-center space-y-4">
          <div className="inline-flex items-center justify-center p-3 rounded-2xl bg-slate-900 border border-slate-700 shadow-xl">
            <Zap className="w-6 h-6 text-yellow-400" />
          </div>
          <div className="space-y-1">
            <h1 className="text-2xl font-heading font-bold text-slate-50 tracking-tight">
              Admin Control Access
            </h1>
            <p className="text-sm text-slate-400">
              Sign in with predefined administrator credentials to access the executive console.
            </p>
          </div>
        </div>
        <Card className="border-slate-800 bg-slate-950/80 backdrop-blur">
          <CardHeader className="space-y-1">
            <CardTitle className="text-lg flex items-center gap-2 text-slate-50">
              <ShieldCheck className="w-5 h-5 text-emerald-400" />
              Secure Admin Login
            </CardTitle>
            <CardDescription className="text-slate-400">
              Only authorised admin accounts can access this panel.
            </CardDescription>
          </CardHeader>
          <CardContent>
            <form onSubmit={handleSubmit} className="space-y-6">
              <div className="space-y-2">
                <Label htmlFor="email" className="text-slate-200">Admin Email</Label>
                <Input
                  id="email"
                  type="email"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  placeholder="admin@example.com"
                  required
                  className="bg-slate-900 border-slate-700 text-slate-50 placeholder:text-slate-500"
                />
              </div>
              <div className="space-y-2">
                <Label htmlFor="password" className="text-slate-200">Password</Label>
                <div className="relative">
                  <Input
                    id="password"
                    type={showPassword ? 'text' : 'password'}
                    value={password}
                    onChange={(e) => setPassword(e.target.value)}
                    required
                    className="bg-slate-900 border-slate-700 text-slate-50 placeholder:text-slate-500 pr-10"
                  />
                  <button
                    type="button"
                    onClick={() => setShowPassword(!showPassword)}
                    className="absolute inset-y-0 right-0 flex items-center pr-3 text-slate-500 hover:text-slate-300"
                  >
                    {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                  </button>
                </div>
              </div>
              <Button
                type="submit"
                disabled={loading}
                className="w-full bg-slate-50 text-slate-900 hover:bg-slate-200"
              >
                {loading ? (
                  <span className="inline-flex items-center gap-2">
                    <Loader />
                    Verifying admin access...
                  </span>
                ) : (
                  <span className="inline-flex items-center gap-2">
                    <ShieldCheck className="w-4 h-4" />
                    Sign in as Admin
                  </span>
                )}
              </Button>
              <div className="flex flex-col items-center gap-1">
                <p className="text-xs text-center text-slate-500">
                  Use your assigned administrator credentials. Regular user accounts cannot access this panel.
                </p>
                <button
                  type="button"
                  onClick={openForgot}
                  className="text-[11px] text-slate-300 hover:text-slate-100 underline-offset-2 hover:underline"
                >
                  Forgot password?
                </button>
              </div>
            </form>
          </CardContent>
        </Card>
        <p className="text-center text-[11px] text-slate-500">
          © 2026 VinTech Solutions. All rights reserved.
        </p>
      </div>
      <Dialog
        open={forgotOpen}
        onOpenChange={(open) => {
          setForgotOpen(open);
          if (!open) {
            resetForgotState();
          }
        }}
      >
        <DialogContent className="max-w-md">
          <DialogHeader>
            <DialogTitle>Admin password reset</DialogTitle>
          </DialogHeader>
          {forgotStep === 1 && (
            <div className="space-y-4">
              <p className="text-xs text-slate-500">
                Enter your registered admin email to receive a one-time password (OTP) for resetting your password.
                The OTP is valid for 10 minutes. You can request up to 5 OTPs per hour, with a short cooldown between requests.
              </p>
              <div className="space-y-2">
                <Label htmlFor="forgot-email" className="text-xs text-slate-700">
                  Admin email
                </Label>
                <Input
                  id="forgot-email"
                  type="email"
                  value={forgotEmail}
                  onChange={(e) => setForgotEmail(e.target.value)}
                  placeholder="admin@example.com"
                  className="text-sm"
                />
              </div>
              <DialogFooter>
                <Button
                  type="button"
                  onClick={handleRequestOtp}
                  disabled={forgotSubmitting}
                >
                  {forgotSubmitting ? (
                    <span className="inline-flex items-center gap-2">
                      <Loader />
                      Sending OTP...
                    </span>
                  ) : (
                    'Send OTP'
                  )}
                </Button>
              </DialogFooter>
            </div>
          )}
          {forgotStep === 2 && (
            <div className="space-y-4">
              <p className="text-xs text-slate-500">
                Enter the 6-digit OTP sent to your admin email to continue. You have up to 5 verification attempts before you need to request a new OTP.
              </p>
              <div className="space-y-2">
                <Label htmlFor="otp" className="text-xs text-slate-700">
                  OTP
                </Label>
                <Input
                  id="otp"
                  type="text"
                  inputMode="numeric"
                  maxLength={6}
                  value={otp}
                  onChange={(e) => setOtp(e.target.value.replace(/\D/g, '').slice(0, 6))}
                  placeholder="Enter 6-digit OTP"
                  className="text-sm tracking-[0.3em]"
                />
              </div>
              <DialogFooter className="flex justify-between">
                <Button
                  type="button"
                  variant="outline"
                  onClick={() => setForgotStep(1)}
                  disabled={forgotSubmitting}
                >
                  Back
                </Button>
                <Button
                  type="button"
                  onClick={handleVerifyOtp}
                  disabled={forgotSubmitting}
                >
                  {forgotSubmitting ? (
                    <span className="inline-flex items-center gap-2">
                      <Loader />
                      Verifying...
                    </span>
                  ) : (
                    'Verify OTP'
                  )}
                </Button>
              </DialogFooter>
            </div>
          )}
          {forgotStep === 3 && (
            <div className="space-y-4">
              <p className="text-xs text-slate-500">
                Set a new password for your admin account after successful OTP verification.
              </p>
              <div className="space-y-2">
                <Label htmlFor="new-password" className="text-xs text-slate-700">
                  New password
                </Label>
                <Input
                  id="new-password"
                  type="password"
                  value={newPassword}
                  onChange={(e) => setNewPassword(e.target.value)}
                  className="text-sm"
                />
              </div>
              <div className="space-y-2">
                <Label htmlFor="confirm-password" className="text-xs text-slate-700">
                  Confirm new password
                </Label>
                <Input
                  id="confirm-password"
                  type="password"
                  value={confirmPassword}
                  onChange={(e) => setConfirmPassword(e.target.value)}
                  className="text-sm"
                />
              </div>
              <DialogFooter className="flex justify-between">
                <Button
                  type="button"
                  variant="outline"
                  onClick={() => setForgotStep(2)}
                  disabled={forgotSubmitting}
                >
                  Back
                </Button>
                <Button
                  type="button"
                  onClick={handleResetPassword}
                  disabled={forgotSubmitting}
                >
                  {forgotSubmitting ? (
                    <span className="inline-flex items-center gap-2">
                      <Loader />
                      Updating...
                    </span>
                  ) : (
                    'Update password'
                  )}
                </Button>
              </DialogFooter>
            </div>
          )}
        </DialogContent>
      </Dialog>
    </div>
  );
}
