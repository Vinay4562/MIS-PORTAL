import { Zap } from 'lucide-react';
import { Card, CardContent } from '@/components/ui/card';

export default function DashboardHome() {
  return (
    <div className="flex flex-col items-center justify-center" style={{ minHeight: 'calc(100vh - 200px)' }}>
      <Card className="max-w-2xl w-full shadow-lg border-2">
        <CardContent className="p-12 text-center">
          <div className="flex justify-center mb-6">
            <div className="p-4 bg-blue-600 rounded-xl">
              <Zap className="w-16 h-16 text-white" />
            </div>
          </div>
          <h1 className="text-3xl md:text-5xl font-heading font-bold mb-4 text-slate-900 dark:text-slate-100">
            Welcome to MIS PORTAL
          </h1>
          <p className="text-base md:text-lg text-slate-600 dark:text-slate-400 mb-6">
            Integrated Substation Data Management System
          </p>
          <p className="text-sm text-slate-500 dark:text-slate-500">
            Select "LINE LOSSES" from the sidebar menu to get started
          </p>
        </CardContent>
      </Card>
    </div>
  );
}
