import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import { TrendingDown, TrendingUp, Activity } from 'lucide-react';
import { formatDate } from '@/lib/utils';

export default function AnalyticsCharts({ entries, feeder }) {
  const chartData = entries.map(entry => ({
    date: formatDate(entry.date),
    lossPercent: parseFloat(entry.loss_percent.toFixed(2)),
    end1Import: entry.end1_import_consumption,
    end1Export: entry.end1_export_consumption,
    end2Import: entry.end2_import_consumption,
    end2Export: entry.end2_export_consumption,
    totalImport: entry.end1_import_consumption + entry.end2_import_consumption,
    totalExport: entry.end1_export_consumption + entry.end2_export_consumption
  }));

  const avgLoss = entries.reduce((sum, e) => sum + e.loss_percent, 0) / entries.length;
  const maxLoss = Math.max(...entries.map(e => e.loss_percent));
  const minLoss = Math.min(...entries.map(e => e.loss_percent));
  const totalImport = entries.reduce((sum, e) => sum + e.end1_import_consumption + e.end2_import_consumption, 0);
  const totalExport = entries.reduce((sum, e) => sum + e.end1_export_consumption + e.end2_export_consumption, 0);

  return (
    <div className="space-y-6">
      {/* Summary Cards */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
        <Card>
          <CardContent className="p-6">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-sm text-slate-500 dark:text-slate-400">Average Loss</p>
                <p className="text-2xl font-bold font-mono-data mt-1">{avgLoss.toFixed(2)}%</p>
              </div>
              <Activity className="w-8 h-8 text-blue-600" />
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardContent className="p-6">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-sm text-slate-500 dark:text-slate-400">Max Loss</p>
                <p className="text-2xl font-bold font-mono-data mt-1 text-red-600">{maxLoss.toFixed(2)}%</p>
              </div>
              <TrendingUp className="w-8 h-8 text-red-600" />
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardContent className="p-6">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-sm text-slate-500 dark:text-slate-400">Min Loss</p>
                <p className="text-2xl font-bold font-mono-data mt-1 text-green-600">{minLoss.toFixed(2)}%</p>
              </div>
              <TrendingDown className="w-8 h-8 text-green-600" />
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardContent className="p-6">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-sm text-slate-500 dark:text-slate-400">Total Import</p>
                <p className="text-2xl font-bold font-mono-data mt-1">{totalImport.toFixed(0)}</p>
                <p className="text-xs text-slate-400">MWH</p>
              </div>
            </div>
          </CardContent>
        </Card>
      </div>

      {/* Charts */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Loss Trend */}
        <Card>
          <CardHeader>
            <CardTitle className="text-lg font-heading">Loss Percentage Trend</CardTitle>
          </CardHeader>
          <CardContent>
            <ResponsiveContainer width="100%" height={300}>
              <LineChart data={chartData}>
                <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                <XAxis 
                  dataKey="date" 
                  stroke="hsl(var(--foreground))"
                  fontSize={12}
                  tickFormatter={(val) => val.split('-')[0]}
                />
                <YAxis 
                  stroke="hsl(var(--foreground))"
                  fontSize={12}
                  label={{ value: '% Loss', angle: -90, position: 'insideLeft' }}
                />
                <Tooltip 
                  contentStyle={{
                    backgroundColor: 'hsl(var(--card))',
                    border: '1px solid hsl(var(--border))',
                    borderRadius: '0.5rem'
                  }}
                />
                <Line 
                  type="monotone" 
                  dataKey="lossPercent" 
                  stroke="hsl(var(--primary))" 
                  strokeWidth={2}
                  dot={{ fill: 'hsl(var(--primary))' }}
                  name="Loss %"
                />
              </LineChart>
            </ResponsiveContainer>
          </CardContent>
        </Card>

        {/* Import vs Export */}
        <Card>
          <CardHeader>
            <CardTitle className="text-lg font-heading">Import vs Export Consumption</CardTitle>
          </CardHeader>
          <CardContent>
            <ResponsiveContainer width="100%" height={300}>
              <BarChart data={chartData}>
                <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                <XAxis 
                  dataKey="date" 
                  stroke="hsl(var(--foreground))"
                  fontSize={12}
                  tickFormatter={(val) => val.split('-')[0]}
                />
                <YAxis 
                  stroke="hsl(var(--foreground))"
                  fontSize={12}
                  label={{ value: 'Consumption (MWH)', angle: -90, position: 'insideLeft' }}
                />
                <Tooltip 
                  contentStyle={{
                    backgroundColor: 'hsl(var(--card))',
                    border: '1px solid hsl(var(--border))',
                    borderRadius: '0.5rem'
                  }}
                />
                <Legend />
                <Bar dataKey="totalImport" fill="hsl(var(--primary))" name="Total Import" />
                <Bar dataKey="totalExport" fill="hsl(var(--secondary))" name="Total Export" />
              </BarChart>
            </ResponsiveContainer>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
