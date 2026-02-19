import { useState, useEffect } from 'react';
import { useAuth } from '@/contexts/AuthContext';
import { useTheme } from '@/components/ThemeProvider';
import { Button } from '@/components/ui/button';
import { Sheet, SheetContent, SheetTrigger } from "@/components/ui/sheet";
import UnifiedReminderBanner from './UnifiedReminderBanner';
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { Menu, Zap, LogOut, Sun, Moon, TrendingDown, Activity, Maximize2, BarChart2, CircleUser, User, FileText } from 'lucide-react';
import { useNavigate, useLocation } from 'react-router-dom';

export default function DashboardLayout({ children }) {
  const { user, logout } = useAuth();
  const { theme, setTheme } = useTheme();
  const navigate = useNavigate();
  const location = useLocation();
  const [open, setOpen] = useState(false);
  const [isCollapsed, setIsCollapsed] = useState(false);
  const [now, setNow] = useState(new Date());

  useEffect(() => {
    const handler = () => {
      setIsCollapsed(true);
      setOpen(false);
    };
    const timer = setInterval(() => {
      setNow(new Date());
    }, 1000);
    window.addEventListener('collapse-sidebar', handler);
    return () => {
      window.removeEventListener('collapse-sidebar', handler);
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

  const SidebarContent = ({ collapsed }) => (
    <div className="flex flex-col h-full">
      <div className={`border-b border-slate-200 dark:border-slate-700 ${collapsed ? 'p-4 flex justify-center' : 'p-6'}`}>
        <div className={`flex items-center ${collapsed ? 'justify-center' : 'gap-3'}`}>
          <div className="p-2 bg-blue-600 rounded-lg">
            <Zap className="w-6 h-6 text-white" />
          </div>
          {!collapsed && (
            <div>
              <h1 className="text-lg font-heading font-bold text-slate-900 dark:text-slate-100">
                MIS PORTAL
              </h1>
              <p className="text-xs text-slate-500 dark:text-slate-400">Dashboard</p>
            </div>
          )}
        </div>
      </div>

      <nav className="flex-1 p-4 space-y-2 overflow-y-auto">
        <button
          onClick={() => {
            navigate('/line-losses');
            setOpen(false);
          }}
          className={`w-full flex items-center ${collapsed ? 'justify-center' : 'gap-3'} px-4 py-3 rounded-lg transition-colors text-left ${
            location.pathname === '/line-losses'
              ? 'bg-blue-600 text-white'
              : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'
          }`}
          data-testid="line-losses-menu"
          title={collapsed ? "LINE LOSSES" : undefined}
        >
          <TrendingDown className="w-5 h-5" />
          {!collapsed && <span className="font-medium">LINE LOSSES</span>}
        </button>

        <button
          onClick={() => {
            navigate('/energy-consumption');
            setOpen(false);
          }}
          className={`w-full flex items-center ${collapsed ? 'justify-center' : 'gap-3'} px-4 py-3 rounded-lg transition-colors text-left ${
            location.pathname === '/energy-consumption'
              ? 'bg-blue-600 text-white'
              : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'
          }`}
          data-testid="energy-consumption-menu"
          title={collapsed ? "ENERGY CONSUMPTION" : undefined}
        >
          <Activity className="w-5 h-5" />
          {!collapsed && <span className="font-medium">ENERGY CONSUMPTION</span>}
        </button>

        <button
          onClick={() => {
            navigate('/max-min-data');
            setOpen(false);
          }}
          className={`w-full flex items-center ${collapsed ? 'justify-center' : 'gap-3'} px-4 py-3 rounded-lg transition-colors text-left ${
            location.pathname === '/max-min-data'
              ? 'bg-blue-600 text-white'
              : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'
          }`}
          data-testid="max-min-data-menu"
          title={collapsed ? "MAX-MIN DATA" : undefined}
        >
          <BarChart2 className="w-5 h-5" />
          {!collapsed && <span className="font-medium">MAX-MIN DATA</span>}
        </button>

        <button
          onClick={() => {
            navigate('/interruptions');
            setOpen(false);
          }}
          className={`w-full flex items-center ${collapsed ? 'justify-center' : 'gap-3'} px-4 py-3 rounded-lg transition-colors text-left ${
            location.pathname === '/interruptions'
              ? 'bg-blue-600 text-white'
              : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'
          }`}
          data-testid="interruptions-menu"
          title={collapsed ? "INTERRUPTIONS" : undefined}
        >
          <Activity className="w-5 h-5" />
          {!collapsed && <span className="font-medium">INTERRUPTIONS</span>}
        </button>

        <button
          onClick={() => {
            navigate('/reports');
            setOpen(false);
          }}
          className={`w-full flex items-center ${collapsed ? 'justify-center' : 'gap-3'} px-4 py-3 rounded-lg transition-colors text-left ${
            location.pathname === '/reports'
              ? 'bg-blue-600 text-white'
              : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700'
          }`}
          data-testid="reports-menu"
          title={collapsed ? "REPORTS" : undefined}
        >
          <FileText className="w-5 h-5" />
          {!collapsed && <span className="font-medium">REPORTS</span>}
        </button>
      </nav>

      <div className={`border-t border-slate-200 dark:border-slate-700 ${collapsed ? 'p-2 flex justify-center' : 'p-4'}`}>
        <div className={`flex items-center ${collapsed ? 'justify-center' : 'gap-3'}`}>
          <div className="w-10 h-10 rounded-full bg-indigo-100 flex items-center justify-center text-indigo-600 font-bold">
            {user?.full_name?.charAt(0) || 'U'}
          </div>
          {!collapsed && (
            <div className="overflow-hidden">
              <p className="text-sm font-medium text-slate-900 dark:text-slate-100 truncate">
                {user?.full_name}
              </p>
              <p className="text-xs text-slate-500 dark:text-slate-400 truncate">
                {user?.email}
              </p>
            </div>
          )}
        </div>
      </div>
    </div>
  );

  return (
    <div className="flex h-screen overflow-hidden bg-slate-50 dark:bg-slate-900">
      {/* Desktop Sidebar */}
      <aside
        onMouseEnter={() => setIsCollapsed(false)}
        onMouseLeave={() => setIsCollapsed(true)}
        className={`hidden md:flex ${isCollapsed ? 'w-20' : 'w-64'} bg-white dark:bg-slate-800 border-r border-slate-200 dark:border-slate-700 flex-col transition-all duration-300`}
      >
        <SidebarContent collapsed={isCollapsed} />
      </aside>

      {/* Main Content */}
      <div className="flex-1 flex flex-col overflow-hidden">
        {/* Header */}
        <header className="bg-white/80 dark:bg-slate-800/80 backdrop-blur-md border-b border-slate-200/50 dark:border-slate-700/50 px-4 md:px-6 py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-4">
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
                  <SidebarContent collapsed={false} />
                </SheetContent>
              </Sheet>

              {/* Desktop Sidebar Trigger */}
              <button
                onClick={() => setIsCollapsed(!isCollapsed)}
                className="hidden md:block p-2 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors"
                title={isCollapsed ? "Expand Sidebar" : "Collapse Sidebar"}
              >
                <Menu className="w-6 h-6" />
              </button>
            </div>

            <div className="flex-1 flex items-center overflow-hidden">
              <UnifiedReminderBanner />
            </div>

            <div className="flex items-center gap-3 ml-auto">
              <div className="hidden md:flex flex-col items-end leading-tight mr-1">
                <span className="text-base font-semibold text-slate-700 dark:text-slate-200">
                  {user?.full_name}
                </span>
                <span className="text-[11px] text-slate-500 dark:text-slate-400 animate-pulse">
                  {formattedDate} • {formattedTime}
                </span>
              </div>
              <DropdownMenu>
                <DropdownMenuTrigger asChild>
                  <Button variant="ghost" size="icon" className="rounded-full">
                    <CircleUser className="w-6 h-6 text-slate-700 dark:text-slate-200" />
                  </Button>
                </DropdownMenuTrigger>
                <DropdownMenuContent align="end">
                  <DropdownMenuLabel>My Account</DropdownMenuLabel>
                  <DropdownMenuSeparator />
                  <DropdownMenuItem onClick={toggleTheme} className="cursor-pointer">
                    {theme === 'light' ? <Moon className="w-4 h-4 mr-2" /> : <Sun className="w-4 h-4 mr-2" />}
                    <span>Toggle Theme</span>
                  </DropdownMenuItem>
                  <DropdownMenuItem onClick={handleLogout} className="cursor-pointer text-red-600 focus:text-red-600">
                    <LogOut className="w-4 h-4 mr-2" />
                    <span>Logout</span>
                  </DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
            </div>
          </div>
        </header>

        {/* Page Content */}
        <main className="flex-1 overflow-y-auto px-4 md:px-6 pb-4 md:pb-6 pt-0 flex flex-col">
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
