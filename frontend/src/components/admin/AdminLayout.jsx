import { useState, useEffect } from 'react';
import { useTheme } from '@/components/ThemeProvider';
import { Button } from '@/components/ui/button';
import { Sheet, SheetContent, SheetTrigger } from '@/components/ui/sheet';
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu';
import { Menu, Zap, LogOut, Sun, Moon, LayoutDashboard, Activity, TrendingDown, BarChart2, Database, KeyRound } from 'lucide-react';
import { useAuth } from '@/contexts/AuthContext';
import { useNavigate, useLocation } from 'react-router-dom';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter } from '@/components/ui/dialog';
import { Input } from '@/components/ui/input';
import { toast } from 'sonner';
import axios from 'axios';

export default function AdminLayout({ children }) {
  const { user, logout } = useAuth();
  const { theme, setTheme } = useTheme();
  const navigate = useNavigate();
  const location = useLocation();
  const [open, setOpen] = useState(false);
  const [changePasswordOpen, setChangePasswordOpen] = useState(false);
  const [currentPassword, setCurrentPassword] = useState('');
  const [newPassword, setNewPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [changingPassword, setChangingPassword] = useState(false);
  const [now, setNow] = useState(new Date());

  const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

  useEffect(() => {
    const timer = setInterval(() => {
      setNow(new Date());
    }, 1000);
    return () => {
      clearInterval(timer);
    };
  }, []);

  const formattedDate = now.toLocaleDateString(undefined, {
    weekday: 'short',
    year: 'numeric',
    month: 'short',
    day: '2-digit',
  });

  const formattedTime = now.toLocaleTimeString(undefined, {
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  });

  const toggleTheme = () => {
    setTheme(theme === 'light' ? 'dark' : 'light');
  };

  const handleLogout = () => {
    logout();
    navigate('/login');
  };

  const resetChangePasswordState = () => {
    setCurrentPassword('');
    setNewPassword('');
    setConfirmPassword('');
  };

  const handleChangePassword = async () => {
    if (!currentPassword || !newPassword || !confirmPassword) {
      toast.error('Fill all password fields');
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
    setChangingPassword(true);
    try {
      const token = localStorage.getItem('token');
      await axios.post(
        `${API}/admin/auth/change-password`,
        {
          current_password: currentPassword,
          new_password: newPassword,
        },
        {
          headers: {
            Authorization: `Bearer ${token}`,
          },
        },
      );
      toast.success('Password updated successfully');
      setChangePasswordOpen(false);
      resetChangePasswordState();
    } catch (error) {
      const detail = error.response?.data?.detail;
      if (detail) {
        toast.error(detail);
      } else {
        toast.error('Failed to update password');
      }
    } finally {
      setChangingPassword(false);
    }
  };

  const SidebarContent = () => (
    <div className="flex flex-col h-full">
      <div className="border-b border-slate-200 dark:border-slate-700 p-6">
        <div className="flex items-center gap-3">
          <div className="p-2 bg-slate-900 rounded-lg">
            <Zap className="w-6 h-6 text-yellow-400" />
          </div>
          <div>
            <h1 className="text-lg font-heading font-bold text-slate-900 dark:text-slate-100">
              Admin Panel
            </h1>
            <p className="text-xs text-slate-500 dark:text-slate-400">Executive Analytics</p>
          </div>
        </div>
      </div>

      <nav className="flex-1 p-4 space-y-2 overflow-y-auto">
        <button
          onClick={() => {
            navigate('/admin');
            setOpen(false);
          }}
          className={`w-full flex items-center gap-3 px-4 py-3 rounded-lg transition-colors text-left ${
            location.pathname === '/admin'
              ? 'bg-slate-900 text-white'
              : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'
          }`}
        >
          <LayoutDashboard className="w-5 h-5" />
          <span className="font-medium">Overview</span>
        </button>
        <button
          onClick={() => {
            navigate('/admin/energy');
            setOpen(false);
          }}
          className={`w-full flex items-center gap-3 px-4 py-3 rounded-lg transition-colors text-left ${
            location.pathname === '/admin/energy'
              ? 'bg-slate-900 text-white'
              : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'
          }`}
        >
          <Activity className="w-5 h-5" />
          <span className="font-medium">Energy Analytics</span>
        </button>
        <button
          onClick={() => {
            navigate('/admin/line-losses');
            setOpen(false);
          }}
          className={`w-full flex items-center gap-3 px-4 py-3 rounded-lg transition-colors text-left ${
            location.pathname === '/admin/line-losses'
              ? 'bg-slate-900 text-white'
              : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'
          }`}
        >
          <TrendingDown className="w-5 h-5" />
          <span className="font-medium">Line Losses</span>
        </button>
        <button
          onClick={() => {
            navigate('/admin/max-min');
            setOpen(false);
          }}
          className={`w-full flex items-center gap-3 px-4 py-3 rounded-lg transition-colors text-left ${
            location.pathname === '/admin/max-min'
              ? 'bg-slate-900 text-white'
              : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'
          }`}
        >
          <BarChart2 className="w-5 h-5" />
          <span className="font-medium">Max–Min Data</span>
        </button>
        <button
          onClick={() => {
            navigate('/admin/station-load');
            setOpen(false);
          }}
          className={`w-full flex items-center gap-3 px-4 py-3 rounded-lg transition-colors text-left ${
            location.pathname === '/admin/station-load'
              ? 'bg-slate-900 text-white'
              : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'
          }`}
        >
          <BarChart2 className="w-5 h-5" />
          <span className="font-medium">Station Load</span>
        </button>
        <button
          onClick={() => {
            navigate('/admin/interruptions');
            setOpen(false);
          }}
          className={`w-full flex items-center gap-3 px-4 py-3 rounded-lg transition-colors text-left ${
            location.pathname === '/admin/interruptions'
              ? 'bg-slate-900 text-white'
              : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'
          }`}
        >
          <Activity className="w-5 h-5" />
          <span className="font-medium">Interruptions</span>
        </button>
        <button
          onClick={() => {
            navigate('/admin/import');
            setOpen(false);
          }}
          className={`w-full flex items-center gap-3 px-4 py-3 rounded-lg transition-colors text-left ${
            location.pathname === '/admin/import'
              ? 'bg-slate-900 text-white'
              : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'
          }`}
        >
          <Database className="w-5 h-5" />
          <span className="font-medium">Bulk Import</span>
        </button>
      </nav>

      <div className="border-t border-slate-200 dark:border-slate-700 p-4">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-full bg-slate-900 flex items-center justify-center text-yellow-400 font-bold">
            {user?.full_name?.charAt(0) || 'A'}
          </div>
          <div className="overflow-hidden">
            <p className="text-sm font-medium text-slate-900 dark:text-slate-100 truncate">
              {user?.full_name}
            </p>
            <p className="text-xs text-slate-500 dark:text-slate-400 truncate">
              {user?.email}
            </p>
          </div>
        </div>
      </div>
    </div>
  );

  return (
    <div className="flex h-screen overflow-hidden bg-slate-50 dark:bg-slate-950">
      <aside className="hidden md:flex w-72 bg-white dark:bg-slate-900 border-r border-slate-200 dark:border-slate-800 flex-col">
        <SidebarContent />
      </aside>

      <div className="flex-1 flex flex-col overflow-hidden">
        <header className="bg-white/90 dark:bg-slate-900/90 backdrop-blur-xl border-b border-slate-200/70 dark:border-slate-800/70 px-4 md:px-8 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <Sheet open={open} onOpenChange={setOpen}>
              <SheetTrigger asChild>
                <button className="md:hidden p-2 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors">
                  <Menu className="w-6 h-6" />
                </button>
              </SheetTrigger>
              <SheetContent side="left" className="p-0 w-80 bg-white dark:bg-slate-900 border-r border-slate-200 dark:border-slate-800">
                <SidebarContent />
              </SheetContent>
            </Sheet>
            <div>
              <p className="text-xs uppercase tracking-[0.25em] text-slate-500 dark:text-slate-400">
                Executive Console
              </p>
              <h2 className="text-lg font-semibold text-slate-900 dark:text-slate-50">
                Grid Reliability Insights
              </h2>
            </div>
          </div>
          <div className="flex items-center gap-3">
            <div className="flex flex-col items-center">
              <DropdownMenu>
                <DropdownMenuTrigger asChild>
                  <Button variant="outline" className="rounded-full border-slate-200 dark:border-slate-700">
                    <div className="flex items-center gap-2">
                      <div className="w-8 h-8 rounded-full bg-slate-900 flex items-center justify-center text-yellow-400 text-sm font-semibold">
                        {user?.full_name?.charAt(0) || 'A'}
                      </div>
                      <div className="hidden md:flex flex-col items-start">
                        <span className="text-xs text-slate-500 dark:text-slate-400">Signed in as</span>
                        <span className="text-sm font-medium text-slate-900 dark:text-slate-50">
                          {user?.full_name}
                        </span>
                      </div>
                    </div>
                  </Button>
                </DropdownMenuTrigger>
                <DropdownMenuContent align="end">
                  <DropdownMenuLabel>Admin Account</DropdownMenuLabel>
                  <DropdownMenuSeparator />
                  <DropdownMenuItem onClick={toggleTheme} className="cursor-pointer">
                    {theme === 'light' ? <Moon className="w-4 h-4 mr-2" /> : <Sun className="w-4 h-4 mr-2" />}
                    <span>Toggle Theme</span>
                  </DropdownMenuItem>
                  <DropdownMenuItem
                    onClick={() => setChangePasswordOpen(true)}
                    className="cursor-pointer"
                  >
                    <KeyRound className="w-4 h-4 mr-2" />
                    <span>Change Password</span>
                  </DropdownMenuItem>
                  <DropdownMenuItem onClick={handleLogout} className="cursor-pointer text-red-600 focus:text-red-600">
                    <LogOut className="w-4 h-4 mr-2" />
                    <span>Logout</span>
                  </DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
              <span className="hidden md:block mt-1 text-[11px] text-slate-500 dark:text-slate-400 animate-pulse">
                {formattedDate} • {formattedTime}
              </span>
            </div>
          </div>
        </header>

        <main className="flex-1 overflow-y-auto px-4 md:px-8 py-6 bg-gradient-to-b from-slate-50/80 to-slate-100/80 dark:from-slate-950 dark:to-slate-900 flex flex-col">
          <div className="max-w-7xl mx-auto space-y-6 flex-1 w-full">
            {children}
          </div>
          <footer className="mt-8 pt-4 border-t border-slate-200/70 dark:border-slate-800/70">
            <p className="text-center text-xs sm:text-sm text-slate-500 dark:text-slate-400">
              © 2026 VinTech Solutions · Admin Panel
            </p>
          </footer>
        </main>
        <Dialog
          open={changePasswordOpen}
          onOpenChange={(open) => {
            setChangePasswordOpen(open);
            if (!open) {
              resetChangePasswordState();
            }
          }}
        >
          <DialogContent className="max-w-md">
            <DialogHeader>
              <DialogTitle>Change admin password</DialogTitle>
            </DialogHeader>
            <div className="space-y-4">
              <div className="space-y-2">
                <p className="text-xs text-slate-500 dark:text-slate-400">
                  Update your admin panel password. You will need to provide your current password for verification.
                </p>
              </div>
              <div className="space-y-2">
                <label className="text-xs text-slate-600 dark:text-slate-300">Current password</label>
                <Input
                  type="password"
                  value={currentPassword}
                  onChange={(e) => setCurrentPassword(e.target.value)}
                  className="text-sm"
                />
              </div>
              <div className="space-y-2">
                <label className="text-xs text-slate-600 dark:text-slate-300">New password</label>
                <Input
                  type="password"
                  value={newPassword}
                  onChange={(e) => setNewPassword(e.target.value)}
                  className="text-sm"
                />
              </div>
              <div className="space-y-2">
                <label className="text-xs text-slate-600 dark:text-slate-300">Confirm new password</label>
                <Input
                  type="password"
                  value={confirmPassword}
                  onChange={(e) => setConfirmPassword(e.target.value)}
                  className="text-sm"
                />
              </div>
            </div>
            <DialogFooter className="flex justify-end gap-2">
              <Button
                type="button"
                variant="outline"
                onClick={() => setChangePasswordOpen(false)}
                disabled={changingPassword}
              >
                Cancel
              </Button>
              <Button
                type="button"
                onClick={handleChangePassword}
                disabled={changingPassword}
              >
                {changingPassword ? 'Saving...' : 'Save password'}
              </Button>
            </DialogFooter>
          </DialogContent>
        </Dialog>
      </div>
    </div>
  );
}
