import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import { Activity, TrendingUp, TrendingDown } from 'lucide-react';
import { formatDate } from '@/lib/utils';

export default function EnergyAnalytics({ entries, sheet }) {
  const meterMap = {};
  (sheet.meters || []).forEach(m => { meterMap[m.id] = m.name; });
  const meterNames = (sheet.meters || []).map(m => m.name);
  const chartData = entries.map(e => {
    const row = { date: formatDate(e.date), total: e.total_consumption };
    meterNames.forEach(n => { row[n] = 0; });
    e.readings.forEach(r => {
      const name = meterMap[r.meter_id];
      if (name) row[name] = r.consumption;
    });
    return row;
  });
  const avgTotal = entries.length ? (entries.reduce((s, e) => s + e.total_consumption, 0) / entries.length) : 0;
  const maxEntry = entries.reduce((acc, e) => e.total_consumption > acc.total_consumption ? e : acc, entries[0] || { total_consumption: 0, date: '-' });
  const minEntry = entries.reduce((acc, e) => e.total_consumption < acc.total_consumption ? e : acc, entries[0] || { total_consumption: 0, date: '-' });
  const sumTotal = entries.reduce((s, e) => s + e.total_consumption, 0);
  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
        <Card>
          <CardContent className="p-6">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-sm text-slate-500 dark:text-slate-400">Average Daily Consumption</p>
                <p className="text-2xl font-bold font-mono-data mt-1">{avgTotal.toFixed(2)}</p>
              </div>
              <Activity className="w-8 h-8 text-blue-600" />
            </div>
          </CardContent>
        </Card>
        <Card>
          <CardContent className="p-6">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-sm text-slate-500 dark:text-slate-400">Max Daily Consumption</p>
                <p className="text-2xl font-bold font-mono-data mt-1 text-red-600">{(maxEntry?.total_consumption || 0).toFixed(2)}</p>
              </div>
              <TrendingUp className="w-8 h-8 text-red-600" />
            </div>
          </CardContent>
        </Card>
        <Card>
          <CardContent className="p-6">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-sm text-slate-500 dark:text-slate-400">Min Daily Consumption</p>
                <p className="text-2xl font-bold font-mono-data mt-1 text-green-600">{(minEntry?.total_consumption || 0).toFixed(2)}</p>
              </div>
              <TrendingDown className="w-8 h-8 text-green-600" />
            </div>
          </CardContent>
        </Card>
        <Card>
          <CardContent className="p-6">
            <div>
              <p className="text-sm text-slate-500 dark:text-slate-400">Total Consumption</p>
              <p className="text-2xl font-bold font-mono-data mt-1">{sumTotal.toFixed(2)}</p>
            </div>
          </CardContent>
        </Card>
      </div>
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <Card>
          <CardHeader>
            <CardTitle className="text-lg font-heading">Total Consumption Trend</CardTitle>
          </CardHeader>
          <CardContent>
            <ResponsiveContainer width="100%" height={300}>
              <LineChart data={chartData}>
                <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                <XAxis dataKey="date" stroke="hsl(var(--foreground))" fontSize={12} tickFormatter={(v) => v.split('-')[0]} />
                <YAxis stroke="hsl(var(--foreground))" fontSize={12} />
                <Tooltip contentStyle={{ backgroundColor: 'hsl(var(--card))', border: '1px solid hsl(var(--border))', borderRadius: '0.5rem' }} />
                <Line type="monotone" dataKey="total" stroke="hsl(var(--primary))" strokeWidth={2} dot={{ fill: 'hsl(var(--primary))' }} name="Total" />
              </LineChart>
            </ResponsiveContainer>
          </CardContent>
        </Card>
        <Card>
          <CardHeader>
            <CardTitle className="text-lg font-heading">Consumption by Meter</CardTitle>
          </CardHeader>
          <CardContent>
            <ResponsiveContainer width="100%" height={300}>
              <BarChart data={chartData}>
                <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                <XAxis dataKey="date" stroke="hsl(var(--foreground))" fontSize={12} tickFormatter={(v) => v.split('-')[0]} />
                <YAxis stroke="hsl(var(--foreground))" fontSize={12} />
                <Tooltip contentStyle={{ backgroundColor: 'hsl(var(--card))', border: '1px solid hsl(var(--border))', borderRadius: '0.5rem' }} />
                <Legend />
                {meterNames.map((n, idx) => (
                  <Bar key={n} dataKey={n} stackId="m" fill={`hsl(var(--primary))`} name={n} />
                ))}
              </BarChart>
            </ResponsiveContainer>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
