import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Activity, TrendingDown, BarChart2, Zap } from 'lucide-react';

export default function AdminDashboardHome() {
  return (
    <div className="space-y-6">
      <div className="grid gap-4 md:grid-cols-4">
        <Card className="col-span-2 bg-gradient-to-br from-slate-900 to-slate-800 text-white shadow-xl">
          <CardHeader>
            <CardTitle className="flex items-center justify-between">
              <span>System Synopsis</span>
              <Zap className="w-5 h-5 text-yellow-400" />
            </CardTitle>
          </CardHeader>
          <CardContent>
            <p className="text-sm text-slate-200">
              Executive overview of energy, losses, max–min performance and interruption health across all feeders.
            </p>
          </CardContent>
        </Card>
        <Card className="bg-white/80 dark:bg-slate-900/80 backdrop-blur">
          <CardHeader>
            <CardTitle className="flex items-center gap-2 text-sm">
              <Activity className="w-4 h-4 text-emerald-500" />
              Energy Analytics
            </CardTitle>
          </CardHeader>
          <CardContent>
            <p className="text-xs text-slate-500 dark:text-slate-400">
              Monitor consumption profiles and boundary meter performance at feeder and portfolio level.
            </p>
          </CardContent>
        </Card>
        <Card className="bg-white/80 dark:bg-slate-900/80 backdrop-blur">
          <CardHeader>
            <CardTitle className="flex items-center gap-2 text-sm">
              <TrendingDown className="w-4 h-4 text-red-500" />
              Line Losses
            </CardTitle>
          </CardHeader>
          <CardContent>
            <p className="text-xs text-slate-500 dark:text-slate-400">
              Track technical losses and identify outlier feeders over any custom period.
            </p>
          </CardContent>
        </Card>
      </div>
      <div className="grid gap-4 md:grid-cols-3">
        <Card className="bg-white/80 dark:bg-slate-900/80 backdrop-blur">
          <CardHeader>
            <CardTitle className="flex items-center gap-2 text-sm">
              <BarChart2 className="w-4 h-4 text-blue-500" />
              Max–Min Performance
            </CardTitle>
          </CardHeader>
          <CardContent>
            <p className="text-xs text-slate-500 dark:text-slate-400">
              Visualise daily max–min envelopes for critical feeders and ICTs.
            </p>
          </CardContent>
        </Card>
        <Card className="bg-white/80 dark:bg-slate-900/80 backdrop-blur">
          <CardHeader>
            <CardTitle className="flex items-center gap-2 text-sm">
              <Activity className="w-4 h-4 text-amber-500" />
              Interruption Insights
            </CardTitle>
          </CardHeader>
          <CardContent>
            <p className="text-xs text-slate-500 dark:text-slate-400">
              Analyse interruption density, duration and cause-of-interruption by feeder cluster.
            </p>
          </CardContent>
        </Card>
        <Card className="bg-white/80 dark:bg-slate-900/80 backdrop-blur">
          <CardHeader>
            <CardTitle className="flex items-center gap-2 text-sm">
              <Activity className="w-4 h-4 text-sky-500" />
              Bulk Import Readiness
            </CardTitle>
          </CardHeader>
          <CardContent>
            <p className="text-xs text-slate-500 dark:text-slate-400">
              Centralised controls for month-wise and year-wise imports across all modules.
            </p>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}

