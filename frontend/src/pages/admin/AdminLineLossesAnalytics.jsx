import { useEffect, useState } from 'react';
import axios from 'axios';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Checkbox } from '@/components/ui/checkbox';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { BlockLoader } from '@/components/ui/loader';
import { toast } from 'sonner';
import { TrendingDown, CalendarRange } from 'lucide-react';
import { LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import { formatDate } from '@/lib/utils';

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

export default function AdminLineLossesAnalytics() {
  const [feeders, setFeeders] = useState([]);
  const [selectedFeeders, setSelectedFeeders] = useState([]);
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [entries, setEntries] = useState([]);
  const [rawEntries, setRawEntries] = useState([]);
  const [loading, setLoading] = useState(false);
  const [columns, setColumns] = useState({
    feeder: true,
    avgLoss: true,
    maxLoss: true,
    minLoss: true,
  });
  const [detailColumns, setDetailColumns] = useState({
    date: true,
    feeder: true,
    end1Import: true,
    end2Import: true,
    totalImport: true,
    lossPercent: true,
  });

  useEffect(() => {
    fetchFeeders();
  }, []);

  const fetchFeeders = async () => {
    try {
      const token = localStorage.getItem('token');
      const resp = await axios.get(`${API}/feeders`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      setFeeders(resp.data || []);
      setSelectedFeeders((resp.data || []).map(f => f.id));
    } catch (e) {
      console.error(e);
      toast.error('Failed to load feeders');
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
      const resp = await axios.get(`${API}/admin/analytics/line-losses`, {
        params,
        headers: { Authorization: `Bearer ${token}` },
      });
      const raw = resp.data?.entries || [];
      setRawEntries(raw);
      const byFeeder = {};
      raw.forEach(e => {
        const key = e.feeder_id;
        if (!byFeeder[key]) {
          byFeeder[key] = { losses: [], totalImport: 0 };
        }
        if (typeof e.loss_percent === 'number') {
          byFeeder[key].losses.push(e.loss_percent);
        }
        const totalImport = (e.end1_import_consumption || 0) + (e.end2_import_consumption || 0);
        byFeeder[key].totalImport += totalImport;
      });
      const rows = Object.keys(byFeeder).map(fid => {
        const stats = byFeeder[fid];
        const feeder = feeders.find(f => f.id === fid) || {};
        const losses = stats.losses;
        const avgLoss = losses.length ? losses.reduce((s, v) => s + v, 0) / losses.length : 0;
        const maxLoss = losses.length ? Math.max(...losses) : 0;
        const minLoss = losses.length ? Math.min(...losses) : 0;
        return {
          feederId: fid,
          feederName: feeder.name || fid,
          avgLoss,
          maxLoss,
          minLoss,
          totalImport: stats.totalImport,
        };
      });
      setEntries(rows);
    } catch (e) {
      console.error(e);
      toast.error('Failed to load line losses analytics');
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

  const buildLossTrend = () => {
    const byDate = {};
    rawEntries.forEach(e => {
      if (!e.date || typeof e.loss_percent !== 'number') return;
      const key = formatDate(e.date);
      if (!byDate[key]) {
        byDate[key] = { date: key, avgLoss: 0, count: 0 };
      }
      byDate[key].avgLoss += e.loss_percent;
      byDate[key].count += 1;
    });
    return Object.values(byDate)
      .map(r => ({
        date: r.date,
        avgLoss: r.count ? r.avgLoss / r.count : 0,
      }))
      .sort((a, b) => {
        if (a.date < b.date) return -1;
        if (a.date > b.date) return 1;
        return 0;
      });
  };

  const buildImportByFeeder = () => {
    const byFeeder = {};
    rawEntries.forEach(e => {
      const fid = e.feeder_id;
      if (!fid) return;
      const totalImport = (e.end1_import_consumption || 0) + (e.end2_import_consumption || 0);
      if (!byFeeder[fid]) {
        byFeeder[fid] = 0;
      }
      byFeeder[fid] += totalImport;
    });
    return Object.keys(byFeeder).map(fid => {
      const feeder = feeders.find(f => f.id === fid) || {};
      return {
        feederId: fid,
        feederName: feeder.name || fid,
        totalImport: byFeeder[fid],
      };
    });
  };

  const buildDetailRows = () => {
    return rawEntries
      .map(e => {
        const feeder = feeders.find(f => f.id === e.feeder_id) || {};
        const end1Import = e.end1_import_consumption || 0;
        const end2Import = e.end2_import_consumption || 0;
        const totalImport = end1Import + end2Import;
        return {
          id: `${e.feeder_id || ''}-${e.date || ''}`,
          date: e.date ? formatDate(e.date) : '',
          feederName: feeder.name || e.feeder_id || '',
          end1Import,
          end2Import,
          totalImport,
          lossPercent: typeof e.loss_percent === 'number' ? e.loss_percent : null,
        };
      })
      .filter(row => row.date && row.feederName);
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row gap-4 md:items-end justify-between">
        <div className="space-y-2">
          <h1 className="text-2xl md:text-3xl font-heading font-bold text-slate-900 dark:text-slate-50">
            Line Losses Analytics
          </h1>
          <p className="text-sm text-slate-500 dark:text-slate-400">
            Compare technical losses across feeders over any custom period.
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
              <TrendingDown className="w-4 h-4 text-amber-600" />
              Daily Loss Details
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
                <label htmlFor="detail-feeder" className="text-xs text-slate-600 dark:text-slate-300">
                  Feeder
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-end1-import"
                  checked={detailColumns.end1Import}
                  onCheckedChange={() => toggleDetailColumn('end1Import')}
                />
                <label
                  htmlFor="detail-end1-import"
                  className="text-xs text-slate-600 dark:text-slate-300"
                >
                  End1 Import
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-end2-import"
                  checked={detailColumns.end2Import}
                  onCheckedChange={() => toggleDetailColumn('end2Import')}
                />
                <label
                  htmlFor="detail-end2-import"
                  className="text-xs text-slate-600 dark:text-slate-300"
                >
                  End2 Import
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-total-import"
                  checked={detailColumns.totalImport}
                  onCheckedChange={() => toggleDetailColumn('totalImport')}
                />
                <label
                  htmlFor="detail-total-import"
                  className="text-xs text-slate-600 dark:text-slate-300"
                >
                  Total Import
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-loss-percent"
                  checked={detailColumns.lossPercent}
                  onCheckedChange={() => toggleDetailColumn('lossPercent')}
                />
                <label
                  htmlFor="detail-loss-percent"
                  className="text-xs text-slate-600 dark:text-slate-300"
                >
                  Loss (%)
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
                    {detailColumns.end1Import && (
                      <TableHead className="text-right">End1 Import</TableHead>
                    )}
                    {detailColumns.end2Import && (
                      <TableHead className="text-right">End2 Import</TableHead>
                    )}
                    {detailColumns.totalImport && (
                      <TableHead className="text-right">Total Import</TableHead>
                    )}
                    {detailColumns.lossPercent && (
                      <TableHead className="text-right">Loss (%)</TableHead>
                    )}
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {buildDetailRows().map(row => (
                    <TableRow key={row.id}>
                      {detailColumns.date && <TableCell>{row.date}</TableCell>}
                      {detailColumns.feeder && <TableCell>{row.feederName}</TableCell>}
                      {detailColumns.end1Import && (
                        <TableCell className="text-right font-mono-data">
                          {row.end1Import.toFixed(2)}
                        </TableCell>
                      )}
                      {detailColumns.end2Import && (
                        <TableCell className="text-right font-mono-data">
                          {row.end2Import.toFixed(2)}
                        </TableCell>
                      )}
                      {detailColumns.totalImport && (
                        <TableCell className="text-right font-mono-data">
                          {row.totalImport.toFixed(2)}
                        </TableCell>
                      )}
                      {detailColumns.lossPercent && (
                        <TableCell className="text-right font-mono-data">
                          {row.lossPercent !== null ? row.lossPercent.toFixed(2) : ''}
                        </TableCell>
                      )}
                    </TableRow>
                  ))}
                  {!buildDetailRows().length && (
                    <TableRow>
                      <TableCell colSpan={6} className="text-center text-sm text-slate-500 py-6">
                        No detailed entries available for the selected period and feeders.
                      </TableCell>
                    </TableRow>
                  )}
                </TableBody>
              </Table>
            </div>
          ) : (
            <p className="text-sm text-slate-500">
              Run analytics to view daily line loss details.
            </p>
          )}
        </CardContent>
      </Card>

      <Card>
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <span className="flex items-center gap-2">
              <TrendingDown className="w-4 h-4 text-emerald-600" />
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
                  <LineChart data={buildLossTrend()}>
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
                    <Line
                      type="monotone"
                      dataKey="avgLoss"
                      stroke="hsl(var(--primary))"
                      strokeWidth={2}
                      dot={{ fill: 'hsl(var(--primary))' }}
                      name="Avg Loss %"
                    />
                  </LineChart>
                </ResponsiveContainer>
              </div>
              <div className="h-72">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={buildImportByFeeder()}>
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
                      dataKey="totalImport"
                      fill="hsl(var(--primary))"
                      name="Total Import"
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
              <TrendingDown className="w-4 h-4 text-red-600" />
              Loss Profile
            </span>
            <div className="flex gap-4 items-center">
              <div className="flex items-center gap-2">
                <Checkbox
                  id="col-avg-loss"
                  checked={columns.avgLoss}
                  onCheckedChange={() => toggleColumn('avgLoss')}
                />
                <label htmlFor="col-avg-loss" className="text-xs text-slate-600 dark:text-slate-300">
                  Avg Loss
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="col-max-loss"
                  checked={columns.maxLoss}
                  onCheckedChange={() => toggleColumn('maxLoss')}
                />
                <label htmlFor="col-max-loss" className="text-xs text-slate-600 dark:text-slate-300">
                  Max Loss
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="col-min-loss"
                  checked={columns.minLoss}
                  onCheckedChange={() => toggleColumn('minLoss')}
                />
                <label htmlFor="col-min-loss" className="text-xs text-slate-600 dark:text-slate-300">
                  Min Loss
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
                    {columns.avgLoss && <TableHead className="text-right">Avg Loss (%)</TableHead>}
                    {columns.maxLoss && <TableHead className="text-right">Max Loss (%)</TableHead>}
                    {columns.minLoss && <TableHead className="text-right">Min Loss (%)</TableHead>}
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {entries.map(row => (
                    <TableRow key={row.feederId}>
                      {columns.feeder && <TableCell>{row.feederName}</TableCell>}
                      {columns.avgLoss && (
                        <TableCell className="text-right font-mono-data">
                          {row.avgLoss.toFixed(2)}
                        </TableCell>
                      )}
                      {columns.maxLoss && (
                        <TableCell className="text-right font-mono-data">
                          {row.maxLoss.toFixed(2)}
                        </TableCell>
                      )}
                      {columns.minLoss && (
                        <TableCell className="text-right font-mono-data">
                          {row.minLoss.toFixed(2)}
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
