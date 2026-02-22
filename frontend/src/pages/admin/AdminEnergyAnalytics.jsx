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

export default function AdminEnergyAnalytics() {
  const [sheets, setSheets] = useState([]);
  const [selectedSheets, setSelectedSheets] = useState([]);
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [loading, setLoading] = useState(false);
  const [entries, setEntries] = useState([]);
  const [rawEntries, setRawEntries] = useState([]);
  const [columns, setColumns] = useState({
    date: true,
    sheet: true,
    total: true,
    avg: true,
    days: true,
  });
  const [detailColumns, setDetailColumns] = useState({
    date: true,
    sheet: true,
    total: true,
  });

  useEffect(() => {
    fetchSheets();
  }, []);

  const fetchSheets = async () => {
    try {
      const token = localStorage.getItem('token');
      const resp = await axios.get(`${API}/energy/sheets`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      setSheets(resp.data || []);
      setSelectedSheets((resp.data || []).map(s => s.id));
    } catch (e) {
      console.error(e);
      toast.error('Failed to load energy sheets');
    }
  };

  const handleFetch = async () => {
    if (!selectedSheets.length) {
      toast.error('Select at least one sheet');
      return;
    }
    try {
      setLoading(true);
      const token = localStorage.getItem('token');
      const params = {
        sheet_ids: selectedSheets.join(','),
      };
      if (startDate) params.start_date = startDate;
      if (endDate) params.end_date = endDate;
      const resp = await axios.get(`${API}/admin/analytics/energy`, {
        params,
        headers: { Authorization: `Bearer ${token}` },
      });
      const rawEntries = resp.data?.entries || [];
      setRawEntries(rawEntries);
      const bySheet = {};
      rawEntries.forEach(e => {
        const key = e.sheet_id;
        if (!bySheet[key]) {
          bySheet[key] = { total: 0, days: 0 };
        }
        bySheet[key].total += e.total_consumption || 0;
        bySheet[key].days += 1;
      });
      const merged = Object.keys(bySheet).map(sheetId => {
        const sheet = (resp.data?.sheets || []).find(s => (s.id || s._id) === sheetId) || {};
        const total = bySheet[sheetId].total;
        const days = bySheet[sheetId].days || 1;
        const avg = total / days;
        return {
          sheetId,
          sheetName: sheet.name || sheetId,
          total,
          days,
          avg,
        };
      });
      setEntries(merged);
    } catch (e) {
      console.error(e);
      toast.error('Failed to load energy analytics');
    } finally {
      setLoading(false);
    }
  };

  const toggleSheet = (id) => {
    setSelectedSheets(prev =>
      prev.includes(id) ? prev.filter(x => x !== id) : [...prev, id],
    );
  };

  const toggleColumn = (key) => {
    setColumns(prev => ({ ...prev, [key]: !prev[key] }));
  };

  const toggleDetailColumn = (key) => {
    setDetailColumns(prev => ({ ...prev, [key]: !prev[key] }));
  };

  const buildTrendData = () => {
    const byDate = {};
    rawEntries.forEach(e => {
      if (!e.date || typeof e.total_consumption !== 'number') return;
      const key = formatDate(e.date);
      if (!byDate[key]) {
        byDate[key] = { date: key, total: 0 };
      }
      byDate[key].total += e.total_consumption;
    });
    return Object.values(byDate).sort((a, b) => {
      if (a.date < b.date) return -1;
      if (a.date > b.date) return 1;
      return 0;
    });
  };

  const buildSheetBreakdown = () => {
    const bySheet = {};
    rawEntries.forEach(e => {
      if (!e.date || typeof e.total_consumption !== 'number') return;
      const dateKey = formatDate(e.date);
      const sheetId = e.sheet_id;
      if (!sheetId) return;
      const sheet = (sheets || []).find(s => (s.id || s._id) === sheetId) || {};
      const name = sheet.name || sheetId;
      if (!bySheet[dateKey]) {
        bySheet[dateKey] = { date: dateKey };
      }
      if (!bySheet[dateKey][name]) {
        bySheet[dateKey][name] = 0;
      }
      bySheet[dateKey][name] += e.total_consumption;
    });
    return Object.values(bySheet).sort((a, b) => {
      if (a.date < b.date) return -1;
      if (a.date > b.date) return 1;
      return 0;
    });
  };

  const buildDetailRows = () => {
    return rawEntries
      .map(e => {
        const sheet = sheets.find(s => (s.id || s._id) === e.sheet_id) || {};
        return {
          id: `${e.sheet_id || ''}-${e.date || ''}`,
          date: e.date ? formatDate(e.date) : '',
          sheetName: sheet.name || e.sheet_id || '',
          total: typeof e.total_consumption === 'number' ? e.total_consumption : null,
        };
      })
      .filter(row => row.date && row.sheetName);
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row gap-4 md:items-end justify-between">
        <div className="space-y-2">
          <h1 className="text-2xl md:text-3xl font-heading font-bold text-slate-900 dark:text-slate-50">
            Energy Analytics
          </h1>
          <p className="text-sm text-slate-500 dark:text-slate-400">
            Portfolio view of boundary meter consumption across selected sheets and date range.
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
            <span>Sheets</span>
            <span className="text-xs text-slate-500">
              Select single, multiple, or all sheets
            </span>
          </CardTitle>
        </CardHeader>
        <CardContent className="space-y-3">
          <div className="flex flex-wrap gap-2">
            {sheets.map(sheet => (
              <button
                key={sheet.id}
                type="button"
                onClick={() => toggleSheet(sheet.id)}
                className={`px-3 py-1.5 rounded-full text-xs border ${
                  selectedSheets.includes(sheet.id)
                    ? 'bg-slate-900 text-white border-slate-900'
                    : 'bg-slate-50 dark:bg-slate-900 text-slate-700 dark:text-slate-300 border-slate-200 dark:border-slate-700'
                }`}
              >
                {sheet.name}
              </button>
            ))}
          </div>
        </CardContent>
      </Card>

      <Card>
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <span className="flex items-center gap-2">
              <Activity className="w-4 h-4 text-amber-600" />
              Daily Sheet Details
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
                  id="detail-sheet"
                  checked={detailColumns.sheet}
                  onCheckedChange={() => toggleDetailColumn('sheet')}
                />
                <label
                  htmlFor="detail-sheet"
                  className="text-xs text-slate-600 dark:text-slate-300"
                >
                  Sheet
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-total"
                  checked={detailColumns.total}
                  onCheckedChange={() => toggleDetailColumn('total')}
                />
                <label
                  htmlFor="detail-total"
                  className="text-xs text-slate-600 dark:text-slate-300"
                >
                  Total Consumption
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
                    {detailColumns.sheet && <TableHead>Sheet</TableHead>}
                    {detailColumns.total && (
                      <TableHead className="text-right">Total Consumption</TableHead>
                    )}
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {buildDetailRows().map(row => (
                    <TableRow key={row.id}>
                      {detailColumns.date && <TableCell>{row.date}</TableCell>}
                      {detailColumns.sheet && <TableCell>{row.sheetName}</TableCell>}
                      {detailColumns.total && (
                        <TableCell className="text-right font-mono-data">
                          {row.total !== null ? row.total.toFixed(2) : ''}
                        </TableCell>
                      )}
                    </TableRow>
                  ))}
                  {!buildDetailRows().length && (
                    <TableRow>
                      <TableCell colSpan={3} className="text-center text-sm text-slate-500 py-6">
                        No detailed entries available for the selected sheets and period.
                      </TableCell>
                    </TableRow>
                  )}
                </TableBody>
              </Table>
            </div>
          ) : (
            <p className="text-sm text-slate-500">
              Run analytics to view daily sheet details.
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
                  <LineChart data={buildTrendData()}>
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
                      dataKey="total"
                      stroke="hsl(var(--primary))"
                      strokeWidth={2}
                      dot={{ fill: 'hsl(var(--primary))' }}
                      name="Total Consumption"
                    />
                  </LineChart>
                </ResponsiveContainer>
              </div>
              <div className="h-72">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={buildSheetBreakdown()}>
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
                    {Array.from(
                      new Set(
                        buildSheetBreakdown().flatMap(row =>
                          Object.keys(row).filter(k => k !== 'date'),
                        ),
                      ),
                    ).map(name => (
                      <Bar
                        key={name}
                        dataKey={name}
                        stackId="sheets"
                        fill="hsl(var(--primary))"
                        name={name}
                      />
                    ))}
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
              <Activity className="w-4 h-4 text-blue-600" />
              Aggregated Metrics
            </span>
            <div className="flex gap-4 items-center">
              <div className="flex items-center gap-2">
                <Checkbox
                  id="col-date"
                  checked={columns.date}
                  onCheckedChange={() => toggleColumn('date')}
                />
                <label htmlFor="col-date" className="text-xs text-slate-600 dark:text-slate-300">
                  Date Visible
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="col-total"
                  checked={columns.total}
                  onCheckedChange={() => toggleColumn('total')}
                />
                <label htmlFor="col-total" className="text-xs text-slate-600 dark:text-slate-300">
                  Total
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="col-avg"
                  checked={columns.avg}
                  onCheckedChange={() => toggleColumn('avg')}
                />
                <label htmlFor="col-avg" className="text-xs text-slate-600 dark:text-slate-300">
                  Average
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
                    {columns.sheet && <TableHead>Sheet</TableHead>}
                    {columns.total && <TableHead className="text-right">Total Consumption</TableHead>}
                    {columns.avg && <TableHead className="text-right">Average Daily</TableHead>}
                    {columns.days && <TableHead className="text-right">Days</TableHead>}
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {entries.map(row => (
                    <TableRow key={row.sheetId}>
                      {columns.sheet && <TableCell>{row.sheetName}</TableCell>}
                      {columns.total && (
                        <TableCell className="text-right font-mono-data">
                          {row.total.toFixed(2)}
                        </TableCell>
                      )}
                      {columns.avg && (
                        <TableCell className="text-right font-mono-data">
                          {row.avg.toFixed(2)}
                        </TableCell>
                      )}
                      {columns.days && (
                        <TableCell className="text-right font-mono-data">
                          {row.days}
                        </TableCell>
                      )}
                    </TableRow>
                  ))}
                  {!entries.length && (
                    <TableRow>
                      <TableCell colSpan={4} className="text-center text-sm text-slate-500 py-6">
                        No analytics loaded yet. Select sheets and date range, then run Load Analytics.
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
