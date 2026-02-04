import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, BarChart, Bar } from 'recharts';
import { Activity, TrendingUp, TrendingDown } from 'lucide-react';
import { formatDate } from '@/lib/utils';

function toFloat(v) {
  const n = parseFloat(v);
  if (isNaN(n)) return null;
  return n;
}

export default function MaxMinAnalytics({ entries, selectedFeeder }) {
  const isBus = selectedFeeder.type === 'bus_station';
  if (isBus) {
    const chartData = entries.map(e => {
      const d = e.data || {};
      const row = {
        date: formatDate(e.date),
        max400: toFloat(d.max_bus_voltage_400kv?.value) ?? null,
        min400: toFloat(d.min_bus_voltage_400kv?.value) ?? null,
        max220: toFloat(d.max_bus_voltage_220kv?.value) ?? null,
        min220: toFloat(d.min_bus_voltage_220kv?.value) ?? null,
        stationLoad: toFloat(d.station_load?.max_mw) ?? null
      };
      return row;
    });
    const validMax400 = chartData.map(r => r.max400).filter(v => v !== null);
    const validMin400 = chartData.map(r => r.min400).filter(v => v !== null);
    const validMax220 = chartData.map(r => r.max220).filter(v => v !== null);
    const validMin220 = chartData.map(r => r.min220).filter(v => v !== null);
    const avg400 = validMax400.length ? validMax400.reduce((s, v) => s + v, 0) / validMax400.length : 0;
    const avg220 = validMax220.length ? validMax220.reduce((s, v) => s + v, 0) / validMax220.length : 0;
    const maxStation = Math.max(...chartData.map(r => r.stationLoad ?? -Infinity));
    const minStation = Math.min(...chartData.map(r => r.stationLoad ?? Infinity));
    return (
      <div className="space-y-6">
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
          <Card><CardContent className="p-6"><div className="flex items-center justify-between"><div><p className="text-sm text-slate-500 dark:text-slate-400">Avg 400KV Max</p><p className="text-2xl font-bold font-mono-data mt-1">{avg400.toFixed(2)}</p></div><Activity className="w-8 h-8 text-blue-600" /></div></CardContent></Card>
          <Card><CardContent className="p-6"><div className="flex items-center justify-between"><div><p className="text-sm text-slate-500 dark:text-slate-400">Avg 220KV Max</p><p className="text-2xl font-bold font-mono-data mt-1">{avg220.toFixed(2)}</p></div><Activity className="w-8 h-8 text-blue-600" /></div></CardContent></Card>
          <Card><CardContent className="p-6"><div className="flex items-center justify-between"><div><p className="text-sm text-slate-500 dark:text-slate-400">Max Station Load MW</p><p className="text-2xl font-bold font-mono-data mt-1 text-red-600">{(isFinite(maxStation) ? maxStation : 0).toFixed(2)}</p></div><TrendingUp className="w-8 h-8 text-red-600" /></div></CardContent></Card>
          <Card><CardContent className="p-6"><div><p className="text-sm text-slate-500 dark:text-slate-400">Min Station Load MW</p><p className="text-2xl font-bold font-mono-data mt-1 text-green-600">{(isFinite(minStation) ? minStation : 0).toFixed(2)}</p></div></CardContent></Card>
        </div>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <Card>
            <CardHeader><CardTitle className="text-lg font-heading">Bus Voltages Trend</CardTitle></CardHeader>
            <CardContent>
              <ResponsiveContainer width="100%" height={300}>
                <LineChart data={chartData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                  <XAxis dataKey="date" stroke="hsl(var(--foreground))" fontSize={12} tickFormatter={(v) => v.split('-')[0]} />
                  <YAxis stroke="hsl(var(--foreground))" fontSize={12} />
                  <Tooltip contentStyle={{ backgroundColor: 'hsl(var(--card))', border: '1px solid hsl(var(--border))', borderRadius: '0.5rem' }} />
                  <Legend />
                  <Line type="monotone" dataKey="max400" stroke="hsl(var(--primary))" strokeWidth={2} name="Max 400KV" />
                  <Line type="monotone" dataKey="max220" stroke="hsl(var(--secondary))" strokeWidth={2} name="Max 220KV" />
                </LineChart>
              </ResponsiveContainer>
            </CardContent>
          </Card>
          <Card>
            <CardHeader><CardTitle className="text-lg font-heading">Station Load Max MW</CardTitle></CardHeader>
            <CardContent>
              <ResponsiveContainer width="100%" height={300}>
                <BarChart data={chartData}>
                  <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                  <XAxis dataKey="date" stroke="hsl(var(--foreground))" fontSize={12} tickFormatter={(v) => v.split('-')[0]} />
                  <YAxis stroke="hsl(var(--foreground))" fontSize={12} />
                  <Tooltip contentStyle={{ backgroundColor: 'hsl(var(--card))', border: '1px solid hsl(var(--border))', borderRadius: '0.5rem' }} />
                  <Bar dataKey="stationLoad" fill="hsl(var(--primary))" name="Max MW" />
                </BarChart>
              </ResponsiveContainer>
            </CardContent>
          </Card>
        </div>
      </div>
    );
  }
  const chartData = entries.map(e => {
    const d = e.data || {};
    const row = {
      date: formatDate(e.date),
      maxMW: toFloat(d.max?.mw) ?? null,
      minMW: toFloat(d.min?.mw) ?? null,
      avgMW: toFloat(d.avg?.mw) ?? null,
      maxAmps: toFloat(d.max?.amps) ?? null,
      minAmps: toFloat(d.min?.amps) ?? null,
      avgAmps: toFloat(d.avg?.amps) ?? null
    };
    return row;
  });
  const validMaxMW = chartData.map(r => r.maxMW).filter(v => v !== null);
  const validMinMW = chartData.map(r => r.minMW).filter(v => v !== null);
  const avgMaxMW = validMaxMW.length ? validMaxMW.reduce((s, v) => s + v, 0) / validMaxMW.length : 0;
  const avgMinMW = validMinMW.length ? validMinMW.reduce((s, v) => s + v, 0) / validMinMW.length : 0;
  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
        <Card><CardContent className="p-6"><div className="flex items-center justify-between"><div><p className="text-sm text-slate-500 dark:text-slate-400">Avg Max MW</p><p className="text-2xl font-bold font-mono-data mt-1">{avgMaxMW.toFixed(2)}</p></div><Activity className="w-8 h-8 text-blue-600" /></div></CardContent></Card>
        <Card><CardContent className="p-6"><div className="flex items-center justify-between"><div><p className="text-sm text-slate-500 dark:text-slate-400">Avg Min MW</p><p className="text-2xl font-bold font-mono-data mt-1">{avgMinMW.toFixed(2)}</p></div><Activity className="w-8 h-8 text-blue-600" /></div></CardContent></Card>
        <Card><CardContent className="p-6"><div className="flex items-center justify-between"><div><p className="text-sm text-slate-500 dark:text-slate-400">Max Amps</p><p className="text-2xl font-bold font-mono-data mt-1 text-red-600">{(Math.max(...chartData.map(r => r.maxAmps ?? -Infinity)) || 0).toFixed(2)}</p></div><TrendingUp className="w-8 h-8 text-red-600" /></div></CardContent></Card>
        <Card><CardContent className="p-6"><div><p className="text-sm text-slate-500 dark:text-slate-400">Min Amps</p><p className="text-2xl font-bold font-mono-data mt-1 text-green-600">{(Math.min(...chartData.map(r => r.minAmps ?? Infinity)) || 0).toFixed(2)}</p></div></CardContent></Card>
      </div>
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <Card>
          <CardHeader><CardTitle className="text-lg font-heading">MW Metrics</CardTitle></CardHeader>
          <CardContent>
            <ResponsiveContainer width="100%" height={300}>
              <LineChart data={chartData}>
                <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                <XAxis dataKey="date" stroke="hsl(var(--foreground))" fontSize={12} tickFormatter={(v) => v.split('-')[0]} />
                <YAxis stroke="hsl(var(--foreground))" fontSize={12} />
                <Tooltip contentStyle={{ backgroundColor: 'hsl(var(--card))', border: '1px solid hsl(var(--border))', borderRadius: '0.5rem' }} />
                <Legend />
                <Line type="monotone" dataKey="maxMW" stroke="hsl(var(--primary))" strokeWidth={2} name="Max MW" />
                <Line type="monotone" dataKey="minMW" stroke="hsl(var(--secondary))" strokeWidth={2} name="Min MW" />
                <Line type="monotone" dataKey="avgMW" stroke="hsl(var(--muted-foreground))" strokeWidth={2} name="Avg MW" />
              </LineChart>
            </ResponsiveContainer>
          </CardContent>
        </Card>
        <Card>
          <CardHeader><CardTitle className="text-lg font-heading">Amps Metrics</CardTitle></CardHeader>
          <CardContent>
            <ResponsiveContainer width="100%" height={300}>
              <LineChart data={chartData}>
                <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                <XAxis dataKey="date" stroke="hsl(var(--foreground))" fontSize={12} tickFormatter={(v) => v.split('-')[0]} />
                <YAxis stroke="hsl(var(--foreground))" fontSize={12} />
                <Tooltip contentStyle={{ backgroundColor: 'hsl(var(--card))', border: '1px solid hsl(var(--border))', borderRadius: '0.5rem' }} />
                <Legend />
                <Line type="monotone" dataKey="maxAmps" stroke="hsl(var(--primary))" strokeWidth={2} name="Max Amps" />
                <Line type="monotone" dataKey="minAmps" stroke="hsl(var(--secondary))" strokeWidth={2} name="Min Amps" />
                <Line type="monotone" dataKey="avgAmps" stroke="hsl(var(--muted-foreground))" strokeWidth={2} name="Avg Amps" />
              </LineChart>
            </ResponsiveContainer>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
