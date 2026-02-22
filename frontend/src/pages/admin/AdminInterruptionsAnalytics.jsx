import { useEffect, useState } from 'react';
import axios from 'axios';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Checkbox } from '@/components/ui/checkbox';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { BlockLoader } from '@/components/ui/loader';
import { toast } from 'sonner';
import { Activity, CalendarRange } from 'lucide-react';
import { LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import { formatDate } from '@/lib/utils';

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

export default function AdminInterruptionsAnalytics() {
  const [feeders, setFeeders] = useState([]);
  const [selectedFeeders, setSelectedFeeders] = useState([]);
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [entries, setEntries] = useState([]);
  const [rawEntries, setRawEntries] = useState([]);
  const [loading, setLoading] = useState(false);
  const [columns, setColumns] = useState({
    feeder: true,
    count: true,
    totalDuration: true,
    avgDuration: true,
  });
  const [detailColumns, setDetailColumns] = useState({
    date: true,
    feeder: true,
    startTime: true,
    endTime: true,
    duration: true,
  });

  useEffect(() => {
    fetchFeeders();
  }, []);

  const fetchFeeders = async () => {
    try {
      const token = localStorage.getItem('token');
      const resp = await axios.get(`${API}/max-min/feeders`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      const interruptionTypes = (resp.data || []).filter(f =>
        ['feeder_400kv', 'feeder_220kv', 'ict_feeder'].includes(f.type),
      );
      setFeeders(interruptionTypes);
      setSelectedFeeders(interruptionTypes.map(f => f.id));
    } catch (e) {
      console.error(e);
      toast.error('Failed to load interruption feeders');
    }
  };

  const handleFetch = async () => {
    if (!selectedFeeders.length) {
      toast.error('Select at least one feeder');
      return;
    }
    try {
      setLoading(true);
      const token = localStorage.getItem('token');
      const params = {
        feeder_ids: selectedFeeders.join(','),
      };
      if (startDate) params.start_date = startDate;
      if (endDate) params.end_date = endDate;
      const resp = await axios.get(`${API}/admin/analytics/interruptions`, {
        params,
        headers: { Authorization: `Bearer ${token}` },
      });
      const raw = resp.data?.entries || [];
      setRawEntries(raw);
      const byFeeder = {};
      raw.forEach(e => {
        const key = e.feeder_id;
        if (!byFeeder[key]) {
          byFeeder[key] = { count: 0, totalMinutes: 0 };
        }
        byFeeder[key].count += 1;
        const minutes = e.data?.duration_minutes;
        if (typeof minutes === 'number') {
          byFeeder[key].totalMinutes += minutes;
        }
      });
      const rows = Object.keys(byFeeder).map(fid => {
        const stats = byFeeder[fid];
        const feeder = feeders.find(f => f.id === fid) || {};
        const avg = stats.count ? stats.totalMinutes / stats.count : 0;
        return {
          feederId: fid,
          feederName: feeder.name || fid,
          count: stats.count,
          totalMinutes: stats.totalMinutes,
          avgMinutes: avg,
        };
      });
      setEntries(rows);
    } catch (e) {
      console.error(e);
      toast.error('Failed to load interruptions analytics');
    } finally {
      setLoading(false);
    }
  };

  const toggleFeeder = (id) => {
    setSelectedFeeders(prev =>
      prev.includes(id) ? prev.filter(x => x !== id) : [...prev, id],
    );
  };

  const toggleColumn = (key) => {
    setColumns(prev => ({ ...prev, [key]: !prev[key] }));
  };

  const toggleDetailColumn = (key) => {
    setDetailColumns(prev => ({ ...prev, [key]: !prev[key] }));
  };

  const buildDailyTrend = () => {
    const byDate = {};
    rawEntries.forEach(e => {
      if (!e.date) return;
      const key = formatDate(e.date);
      const minutes = e.data && typeof e.data.duration_minutes === 'number' ? e.data.duration_minutes : 0;
      if (!byDate[key]) {
        byDate[key] = { date: key, count: 0, totalMinutes: 0 };
      }
      byDate[key].count += 1;
      byDate[key].totalMinutes += minutes;
    });
    return Object.values(byDate).sort((a, b) => {
      if (a.date < b.date) return -1;
      if (a.date > b.date) return 1;
      return 0;
    });
  };

  const buildFeederMinutes = () => {
    const byFeeder = {};
    rawEntries.forEach(e => {
      const fid = e.feeder_id;
      if (!fid) return;
      const minutes = e.data && typeof e.data.duration_minutes === 'number' ? e.data.duration_minutes : 0;
      if (!byFeeder[fid]) {
        byFeeder[fid] = 0;
      }
      byFeeder[fid] += minutes;
    });
    return Object.keys(byFeeder).map(fid => {
      const feeder = feeders.find(f => f.id === fid) || {};
      return {
        feederId: fid,
        feederName: feeder.name || fid,
        totalMinutes: byFeeder[fid],
      };
    });
  };

  const buildDetailRows = () => {
    return rawEntries
      .map(e => {
        const feeder = feeders.find(f => f.id === e.feeder_id) || {};
        const data = e.data || {};
        return {
          id: `${e.feeder_id || ''}-${e.date || ''}-${data.start_time || ''}`,
          date: e.date ? formatDate(e.date) : '',
          feederName: feeder.name || e.feeder_id || '',
          startTime: data.start_time,
          endTime: data.end_time,
          durationMinutes:
            typeof data.duration_minutes === 'number' ? data.duration_minutes : null,
        };
      })
      .filter(row => row.date && row.feederName && row.startTime);
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row gap-4 md:items-end justify-between">
        <div className="space-y-2">
          <h1 className="text-2xl md:text-3xl font-heading font-bold text-slate-900 dark:text-slate-50">
            Interruption Analytics
          </h1>
          <p className="text-sm text-slate-500 dark:text-slate-400">
            Density and duration profile of interruptions by feeder cluster.
          </p>
        </div>
        <div className="flex flex-wrap gap-3 items-end">
          <div className="space-y-1">
            <label className="text-xs font-medium text-slate-500 flex items-center gap-1">
              <CalendarRange className="w-3 h-3" />
              From
            </label>
            <Input
              type="date"
              value={startDate}
              onChange={e => setStartDate(e.target.value)}
              className="w-40"
            />
          </div>
          <div className="space-y-1">
            <label className="text-xs font-medium text-slate-500">To</label>
            <Input
              type="date"
              value={endDate}
              onChange={e => setEndDate(e.target.value)}
              className="w-40"
            />
          </div>
          <Button onClick={handleFetch} disabled={loading}>
            {loading ? <BlockLoader /> : 'Load Analytics'}
          </Button>
        </div>
      </div>

      <Card>
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <span>Feeders</span>
            <span className="text-xs text-slate-500">
              Single, multiple, or all feeders
            </span>
          </CardTitle>
        </CardHeader>
        <CardContent className="space-y-3">
          <div className="flex flex-wrap gap-2">
            {feeders.map(feeder => (
              <button
                key={feeder.id}
                type="button"
                onClick={() => toggleFeeder(feeder.id)}
                className={`px-3 py-1.5 rounded-full text-xs border ${
                  selectedFeeders.includes(feeder.id)
                    ? 'bg-slate-900 text-white border-slate-900'
                    : 'bg-slate-50 dark:bg-slate-900 text-slate-700 dark:text-slate-300 border-slate-200 dark:border-slate-700'
                }`}
              >
                {feeder.name}
              </button>
            ))}
          </div>
        </CardContent>
      </Card>

      <Card>
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <span className="flex items-center gap-2">
              <Activity className="w-4 h-4 text-sky-600" />
              Interruption Events
            </span>
            <div className="flex flex-wrap gap-3 items-center">
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-date"
                  checked={detailColumns.date}
                  onCheckedChange={() => toggleDetailColumn('date')}
                />
                <label htmlFor="detail-date" className="text-xs text-slate-600 dark:text-slate-300">
                  Date
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-feeder"
                  checked={detailColumns.feeder}
                  onCheckedChange={() => toggleDetailColumn('feeder')}
                />
                <label
                  htmlFor="detail-feeder"
                  className="text-xs text-slate-600 dark:text-slate-300"
                >
                  Feeder
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-start"
                  checked={detailColumns.startTime}
                  onCheckedChange={() => toggleDetailColumn('startTime')}
                />
                <label
                  htmlFor="detail-start"
                  className="text-xs text-slate-600 dark:text-slate-300"
                >
                  Start Time
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-end"
                  checked={detailColumns.endTime}
                  onCheckedChange={() => toggleDetailColumn('endTime')}
                />
                <label htmlFor="detail-end" className="text-xs text-slate-600 dark:text-slate-300">
                  End Time
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-duration"
                  checked={detailColumns.duration}
                  onCheckedChange={() => toggleDetailColumn('duration')}
                />
                <label
                  htmlFor="detail-duration"
                  className="text-xs text-slate-600 dark:text-slate-300"
                >
                  Duration (min)
                </label>
              </div>
            </div>
          </CardTitle>
        </CardHeader>
        <CardContent>
          {loading ? (
            <div className="py-12 flex justify-center">
              <BlockLoader />
            </div>
          ) : rawEntries.length ? (
            <div className="border rounded-md overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    {detailColumns.date && <TableHead>Date</TableHead>}
                    {detailColumns.feeder && <TableHead>Feeder</TableHead>}
                    {detailColumns.startTime && (
                      <TableHead className="text-right">Start Time</TableHead>
                    )}
                    {detailColumns.endTime && (
                      <TableHead className="text-right">End Time</TableHead>
                    )}
                    {detailColumns.duration && (
                      <TableHead className="text-right">Duration (min)</TableHead>
                    )}
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {buildDetailRows().map(row => (
                    <TableRow key={row.id}>
                      {detailColumns.date && <TableCell>{row.date}</TableCell>}
                      {detailColumns.feeder && <TableCell>{row.feederName}</TableCell>}
                      {detailColumns.startTime && (
                        <TableCell className="text-right font-mono-data">
                          {row.startTime}
                        </TableCell>
                      )}
                      {detailColumns.endTime && (
                        <TableCell className="text-right font-mono-data">
                          {row.endTime ?? ''}
                        </TableCell>
                      )}
                      {detailColumns.duration && (
                        <TableCell className="text-right font-mono-data">
                          {row.durationMinutes !== null ? row.durationMinutes.toFixed(2) : ''}
                        </TableCell>
                      )}
                    </TableRow>
                  ))}
                  {!buildDetailRows().length && (
                    <TableRow>
                      <TableCell colSpan={5} className="text-center text-sm text-slate-500 py-6">
                        No interruption events available for the selected period and feeders.
                      </TableCell>
                    </TableRow>
                  )}
                </TableBody>
              </Table>
            </div>
          ) : (
            <p className="text-sm text-slate-500">
              Run analytics to view detailed interruption events.
            </p>
          )}
        </CardContent>
      </Card>

      <Card>
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <span className="flex items-center gap-2">
              <Activity className="w-4 h-4 text-emerald-600" />
              Portfolio Charts
            </span>
          </CardTitle>
        </CardHeader>
        <CardContent>
          {loading ? (
            <div className="py-12 flex justify-center">
              <BlockLoader />
            </div>
          ) : rawEntries.length ? (
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              <div className="h-72">
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart data={buildDailyTrend()}>
                    <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                    <XAxis
                      dataKey="date"
                      stroke="hsl(var(--foreground))"
                      fontSize={12}
                      tickFormatter={v => v.split('-')[0]}
                    />
                    <YAxis stroke="hsl(var(--foreground))" fontSize={12} />
                    <Tooltip
                      contentStyle={{
                        backgroundColor: 'hsl(var(--card))',
                        border: '1px solid hsl(var(--border))',
                        borderRadius: '0.5rem',
                      }}
                    />
                    <Legend />
                    <Line
                      type="monotone"
                      dataKey="count"
                      stroke="hsl(var(--primary))"
                      strokeWidth={2}
                      name="Interruptions"
                    />
                    <Line
                      type="monotone"
                      dataKey="totalMinutes"
                      stroke="hsl(var(--secondary))"
                      strokeWidth={2}
                      name="Total Minutes"
                    />
                  </LineChart>
                </ResponsiveContainer>
              </div>
              <div className="h-72">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={buildFeederMinutes()}>
                    <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                    <XAxis
                      dataKey="feederName"
                      stroke="hsl(var(--foreground))"
                      fontSize={12}
                    />
                    <YAxis stroke="hsl(var(--foreground))" fontSize={12} />
                    <Tooltip
                      contentStyle={{
                        backgroundColor: 'hsl(var(--card))',
                        border: '1px solid hsl(var(--border))',
                        borderRadius: '0.5rem',
                      }}
                    />
                    <Bar
                      dataKey="totalMinutes"
                      fill="hsl(var(--primary))"
                      name="Total Minutes"
                    />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          ) : (
            <p className="text-sm text-slate-500">
              Run analytics to view portfolio charts.
            </p>
          )}
        </CardContent>
      </Card>

      <Card>
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <span className="flex items-center gap-2">
              <Activity className="w-4 h-4 text-amber-600" />
              Interruption Summary
            </span>
            <div className="flex gap-4 items-center">
              <div className="flex items-center gap-2">
                <Checkbox
                  id="col-count"
                  checked={columns.count}
                  onCheckedChange={() => toggleColumn('count')}
                />
                <label htmlFor="col-count" className="text-xs text-slate-600 dark:text-slate-300">
                  Count
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="col-total-duration"
                  checked={columns.totalDuration}
                  onCheckedChange={() => toggleColumn('totalDuration')}
                />
                <label
                  htmlFor="col-total-duration"
                  className="text-xs text-slate-600 dark:text-slate-300"
                >
                  Total Duration
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="col-avg-duration"
                  checked={columns.avgDuration}
                  onCheckedChange={() => toggleColumn('avgDuration')}
                />
                <label
                  htmlFor="col-avg-duration"
                  className="text-xs text-slate-600 dark:text-slate-300"
                >
                  Avg Duration
                </label>
              </div>
            </div>
          </CardTitle>
        </CardHeader>
        <CardContent>
          {loading ? (
            <div className="py-12 flex justify-center">
              <BlockLoader />
            </div>
          ) : (
            <div className="border rounded-md overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    {columns.feeder && <TableHead>Feeder</TableHead>}
                    {columns.count && <TableHead className="text-right">Interruptions</TableHead>}
                    {columns.totalDuration && (
                      <TableHead className="text-right">Total Minutes</TableHead>
                    )}
                    {columns.avgDuration && (
                      <TableHead className="text-right">Avg Minutes</TableHead>
                    )}
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {entries.map(row => (
                    <TableRow key={row.feederId}>
                      {columns.feeder && <TableCell>{row.feederName}</TableCell>}
                      {columns.count && (
                        <TableCell className="text-right font-mono-data">
                          {row.count}
                        </TableCell>
                      )}
                      {columns.totalDuration && (
                        <TableCell className="text-right font-mono-data">
                          {row.totalMinutes.toFixed(2)}
                        </TableCell>
                      )}
                      {columns.avgDuration && (
                        <TableCell className="text-right font-mono-data">
                          {row.avgMinutes.toFixed(2)}
                        </TableCell>
                      )}
                    </TableRow>
                  ))}
                  {!entries.length && (
                    <TableRow>
                      <TableCell colSpan={4} className="text-center text-sm text-slate-500 py-6">
                        No analytics loaded yet. Select feeders and date range, then run Load Analytics.
                      </TableCell>
                    </TableRow>
                  )}
                </TableBody>
              </Table>
            </div>
          )}
        </CardContent>
      </Card>
    </div>
  );
}
