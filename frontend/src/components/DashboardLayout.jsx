import { useState } from 'react';
import { useAuth } from '@/contexts/AuthContext';
import { useTheme } from '@/components/ThemeProvider';
import { Button } from '@/components/ui/button';
import { Sheet, SheetContent, SheetTrigger } from "@/components/ui/sheet";
import { Menu, Zap, LogOut, Sun, Moon, TrendingDown, Activity, Maximize2, BarChart2 } from 'lucide-react';
import { useNavigate, useLocation } from 'react-router-dom';

export default function DashboardLayout({ children }) {
  const { user, logout } = useAuth();
  const { theme, setTheme } = useTheme();
  const navigate = useNavigate();
  const location = useLocation();
  const [open, setOpen] = useState(false);

  const toggleTheme = () => {
    setTheme(theme === 'light' ? 'dark' : 'light');
  };

  const handleLogout = () => {
    logout();
    navigate('/login');
  };

  const SidebarContent = () => (
    <div className="flex flex-col h-full">
      <div className="p-6 border-b border-slate-200 dark:border-slate-700">
        <div className="flex items-center gap-3">
          <div className="p-2 bg-blue-600 rounded-lg">
            <Zap className="w-6 h-6 text-white" />
          </div>
          <div>
            <h1 className="text-lg font-heading font-bold text-slate-900 dark:text-slate-100">
              MIS PORTAL
            </h1>
            <p className="text-xs text-slate-500 dark:text-slate-400">Dashboard</p>
          </div>
        </div>
      </div>

      <nav className="flex-1 p-4 space-y-2 overflow-y-auto">
        <button
          onClick={() => {
            navigate('/line-losses');
            setOpen(false);
          }}
          className={`w-full flex items-center gap-3 px-4 py-3 rounded-lg transition-colors text-left ${
            location.pathname === '/line-losses'
              ? 'bg-blue-600 text-white'
              : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'
          }`}
          data-testid="line-losses-menu"
        >
          <TrendingDown className="w-5 h-5" />
          <span className="font-medium">LINE LOSSES</span>
        </button>

        <button
          onClick={() => {
            navigate('/energy-consumption');
            setOpen(false);
          }}
          className={`w-full flex items-center gap-3 px-4 py-3 rounded-lg transition-colors text-left ${
            location.pathname === '/energy-consumption'
              ? 'bg-blue-600 text-white'
              : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'
          }`}
          data-testid="energy-consumption-menu"
        >
          <Activity className="w-5 h-5" />
          <span className="font-medium">ENERGY CONSUMPTION</span>
        </button>

        <button
          onClick={() => {
            navigate('/max-min-data');
            setOpen(false);
          }}
          className={`w-full flex items-center gap-3 px-4 py-3 rounded-lg transition-colors text-left ${
            location.pathname === '/max-min-data'
              ? 'bg-blue-600 text-white'
              : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'
          }`}
          data-testid="max-min-data-menu"
        >
          <BarChart2 className="w-5 h-5" />
          <span className="font-medium">MAX-MIN DATA</span>
        </button>
      </nav>

      <div className="p-4 border-t border-slate-200 dark:border-slate-700 space-y-2">
        <div className="px-4 py-2 text-sm text-slate-600 dark:text-slate-400">
          <p className="font-medium truncate">{user?.full_name}</p>
          <p className="text-xs truncate">{user?.email}</p>
        </div>
      </div>
    </div>
  );

  return (
    <div className="flex h-screen overflow-hidden bg-slate-50 dark:bg-slate-900">
      {/* Desktop Sidebar */}
      <aside className="hidden md:flex w-64 bg-white dark:bg-slate-800 border-r border-slate-200 dark:border-slate-700 flex-col transition-all duration-300">
        <SidebarContent />
      </aside>

      {/* Main Content */}
      <div className="flex-1 flex flex-col overflow-hidden">
        {/* Header */}
        <header className="bg-white/80 dark:bg-slate-800/80 backdrop-blur-md border-b border-slate-200/50 dark:border-slate-700/50 px-4 md:px-6 py-4">
          <div className="flex items-center justify-between">
            {/* Mobile Sidebar Trigger */}
            <Sheet open={open} onOpenChange={setOpen}>
              <SheetTrigger asChild>
                <button
                  className="md:hidden p-2 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors"
                  data-testid="sidebar-toggle"
                >
                  <Menu className="w-6 h-6" />
                </button>
              </SheetTrigger>
              <SheetContent side="left" className="p-0 w-72 bg-white dark:bg-slate-800 border-r border-slate-200 dark:border-slate-700">
                <SidebarContent />
              </SheetContent>
            </Sheet>

            <div className="flex items-center gap-3 ml-auto">
              <Button
                variant="ghost"
                size="icon"
                onClick={toggleTheme}
                data-testid="theme-toggle"
              >
                {theme === 'light' ? <Moon className="w-5 h-5" /> : <Sun className="w-5 h-5" />}
              </Button>
              
              <Button
                variant="ghost"
                size="sm"
                onClick={handleLogout}
                data-testid="logout-button"
                className="gap-2"
              >
                <LogOut className="w-4 h-4" />
                Logout
              </Button>
            </div>
          </div>
        </header>

        {/* Page Content */}
        <main className="flex-1 overflow-y-auto p-4 md:p-6 flex flex-col">
          <div className="flex-1">
            {children}
          </div>
          
          {/* Footer */}
          <footer className="mt-auto pt-8 pb-2">
            <p className="text-center text-sm font-medium text-slate-500 dark:text-slate-400">
              © 2026 VinTech Solutions. All rights reserved.
            </p>
          </footer>
        </main>
      </div>
    </div>
  );
}
