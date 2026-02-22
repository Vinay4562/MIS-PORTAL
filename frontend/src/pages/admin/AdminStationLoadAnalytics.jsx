import { Fragment, useEffect, useMemo, useState } from 'react';
import axios from 'axios';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Checkbox } from '@/components/ui/checkbox';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { BlockLoader } from '@/components/ui/loader';
import { toast } from 'sonner';
import { BarChart2, CalendarRange, Download } from 'lucide-react';
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import { formatDate, downloadFile } from '@/lib/utils';

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

export default function AdminStationLoadAnalytics() {
  const [feeders, setFeeders] = useState([]);
  const [selectedFeeders, setSelectedFeeders] = useState([]);
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [mode, setMode] = useState('day');
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(false);
  const [showStation, setShowStation] = useState(true);
  const [metricColumns, setMetricColumns] = useState({
    amps: true,
    mw: true,
    mvar: true,
    mva: true,
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
      const ictOnly = (resp.data || []).filter(f => f.type === 'ict_feeder');
      setFeeders(ictOnly);
      setSelectedFeeders(ictOnly.map(f => f.id));
    } catch (e) {
      console.error(e);
      toast.error('Failed to load ICT feeders for Station Load');
    }
  };

  const handleToggleFeeder = (id) => {
    setSelectedFeeders(prev =>
      prev.includes(id) ? prev.filter(x => x !== id) : [...prev, id],
    );
  };

  const handleFetch = async () => {
    if (!selectedFeeders.length) {
      toast.error('Select at least one ICT');
      return;
    }
    try {
      setLoading(true);
      const token = localStorage.getItem('token');
      const params = {
        feeder_ids: selectedFeeders.join(','),
        mode,
      };
      if (startDate) params.start_date = startDate;
      if (endDate) params.end_date = endDate;
      const resp = await axios.get(`${API}/admin/analytics/station-load`, {
        params,
        headers: { Authorization: `Bearer ${token}` },
      });
      const respFeeders = resp.data?.feeders || [];
      if (respFeeders.length) {
        const ictOnly = respFeeders.filter(f => f.type === 'ict_feeder');
        setFeeders(ictOnly);
      }
      setRows(resp.data?.rows || []);
    } catch (e) {
      console.error(e);
      toast.error('Failed to load station load analytics');
    } finally {
      setLoading(false);
    }
  };

  const buildStationTrend = () => {
    return rows.map(row => {
      const station = row.station || {};
      const label = mode === 'day' ? formatDate(row.period) : row.period;
      return {
        period: label,
        stationMw: typeof station.mw === 'number' ? station.mw : null,
        stationMvar: typeof station.mvar === 'number' ? station.mvar : null,
        stationMva: typeof station.mva === 'number' ? station.mva : null,
      };
    });
  };

  const sortedFeeders = [...feeders].sort((a, b) => (a.name || '').localeCompare(b.name || ''));
  const activeMetricKeys = ['amps', 'mw', 'mvar', 'mva'].filter(key => metricColumns[key]);
  const stationMetricKeys = showStation
    ? ['mva', 'mw', 'mvar'].filter(key => metricColumns[key])
    : [];

  const toggleMetricColumn = (key) => {
    setMetricColumns(prev => ({
      ...prev,
      [key]: !prev[key],
    }));
  };

  const handleToggleAllFeeders = () => {
    setSelectedFeeders(prev =>
      prev.length === sortedFeeders.length ? [] : sortedFeeders.map(f => f.id),
    );
  };

  const handleToggleStation = () => {
    setShowStation(prev => !prev);
  };

  const handleExport = async () => {
    if (!rows.length) {
      toast.error('No data to export. Run analytics first.');
      return;
    }
    if (!selectedFeeders.length) {
      toast.error('Select at least one ICT to export.');
      return;
    }
    const metricKeys = Object.keys(metricColumns).filter(key => metricColumns[key]);
    if (!metricKeys.length) {
      toast.error('Select at least one metric to export.');
      return;
    }
    try {
      const token = localStorage.getItem('token');
      const params = {
        feeder_ids: selectedFeeders.join(','),
        mode,
        metrics: metricKeys.join(','),
        include_station: showStation,
      };
      if (startDate) params.start_date = startDate;
      if (endDate) params.end_date = endDate;
      const resp = await axios.get(`${API}/admin/analytics/station-load/export`, {
        params,
        headers: { Authorization: `Bearer ${token}` },
        responseType: 'blob',
      });
      const labelMode = mode === 'day' ? 'Daily' : 'Monthly';
      const parts = ['StationLoad', labelMode];
      if (startDate) parts.push(startDate);
      if (endDate) parts.push(endDate);
      const filename = `${parts.join('_')}.xlsx`;
      await downloadFile(resp.data, filename);
    } catch (error) {
      console.error('Station Load export failed', error);
      toast.error('Failed to export Station Load data');
    }
  };

  const formatInt = (value) => {
    if (typeof value !== 'number') return '';
    return Math.round(value).toString();
  };

  const formatDecimal = (value) => {
    if (typeof value !== 'number') return '';
    return value.toFixed(2);
  };

  const formatTimeDisplay = (value) => {
    if (!value) return '';
    const s = String(value).trim();
    const parts = s.split(':');
    if (parts.length >= 2) {
      const hh = parts[0].padStart(2, '0');
      const mm = parts[1].slice(0, 2).padStart(2, '0');
      return `${hh}:${mm}`;
    }
    return s;
  };

  const getStationDateTime = (row) => {
    const perIct = row.per_ict || {};
    let best = null;
    Object.values(perIct).forEach(v => {
      if (typeof v.mw === 'number') {
        if (!best || v.mw > best.mw) {
          best = v;
        }
      }
    });
    const dateRaw = best?.date || row.period;
    const timeRaw = best?.time || '';
    const date = dateRaw ? formatDate(dateRaw) : '';
    const time = timeRaw ? formatTimeDisplay(timeRaw) : '';
    return { date, time };
  };

  const maxRow = useMemo(() => {
    if (!rows.length) return null;
    let best = null;
    rows.forEach(r => {
      const station = r.station || {};
      if (typeof station.mw === 'number') {
        if (!best || station.mw > ((best.station || {}).mw ?? -Infinity)) {
          best = r;
        }
      }
    });
    return best;
  }, [rows]);

  let maxRowDisplay = null;
  if (maxRow) {
    const { date, time } = getStationDateTime(maxRow);
    const station = maxRow.station || {};
    maxRowDisplay = {
      date,
      time,
      mw: typeof station.mw === 'number' ? station.mw : null,
      mva: typeof station.mva === 'number' ? station.mva : null,
      period: maxRow.period,
    };
  }

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row gap-4 md:items-end justify-between">
        <div className="space-y-2">
          <h1 className="text-2xl md:text-3xl font-heading font-bold text-slate-900 dark:text-slate-50">
            Station Load
          </h1>
          <p className="text-sm text-slate-500 dark:text-slate-400">
            Maximum ICT loading and aggregated station load across the selected period.
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
          <div className="space-y-1">
            <p className="text-xs font-medium text-slate-500">Display Mode</p>
            <div className="inline-flex rounded-full border border-slate-200 dark:border-slate-700 p-0.5 bg-slate-50/60 dark:bg-slate-900/60">
              <button
                type="button"
                onClick={() => setMode('day')}
                className={`px-3 py-1.5 text-xs rounded-full ${
                  mode === 'day'
                    ? 'bg-slate-900 text-white'
                    : 'text-slate-600 dark:text-slate-300'
                }`}
              >
                Day-wise
              </button>
              <button
                type="button"
                onClick={() => setMode('month')}
                className={`px-3 py-1.5 text-xs rounded-full ${
                  mode === 'month'
                    ? 'bg-slate-900 text-white'
                    : 'text-slate-600 dark:text-slate-300'
                }`}
              >
                Month-wise
              </button>
            </div>
          </div>
          <Button onClick={handleFetch} disabled={loading}>
            {loading ? <BlockLoader /> : 'Load Analytics'}
          </Button>
        </div>
      </div>

      <Card>
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <span>ICT Feeders</span>
            <span className="text-xs text-slate-500">
              Single, multiple, or all ICTs (ICT-1 to ICT-4)
            </span>
          </CardTitle>
        </CardHeader>
        <CardContent className="space-y-3">
          <div className="flex flex-wrap gap-2">
            {sortedFeeders.length > 0 && (
              <button
                type="button"
                onClick={handleToggleAllFeeders}
                className={`px-3 py-1.5 rounded-full text-xs border ${
                  selectedFeeders.length === sortedFeeders.length && sortedFeeders.length > 0
                    ? 'bg-slate-900 text-white border-slate-900'
                    : 'bg-slate-50 dark:bg-slate-900 text-slate-700 dark:text-slate-300 border-slate-200 dark:border-slate-700'
                }`}
              >
                All ICTs
              </button>
            )}
            <button
              type="button"
              onClick={handleToggleStation}
              className={`px-3 py-1.5 rounded-full text-xs border ${
                showStation
                  ? 'bg-slate-900 text-white border-slate-900'
                  : 'bg-slate-50 dark:bg-slate-900 text-slate-700 dark:text-slate-300 border-slate-200 dark:border-slate-700'
              }`}
            >
              Station Load
            </button>
            {sortedFeeders.map(feeder => (
              <button
                key={feeder.id}
                type="button"
                onClick={() => handleToggleFeeder(feeder.id)}
                className={`px-3 py-1.5 rounded-full text-xs border ${
                  selectedFeeders.includes(feeder.id)
                    ? 'bg-slate-900 text-white border-slate-900'
                    : 'bg-slate-50 dark:bg-slate-900 text-slate-700 dark:text-slate-300 border-slate-200 dark:border-slate-700'
                }`}
              >
                {feeder.name}
              </button>
            ))}
            {!sortedFeeders.length && (
              <span className="text-xs text-slate-500">
                No ICT feeders available. Initialise Max–Min ICT feeders first.
              </span>
            )}
          </div>
        </CardContent>
      </Card>

      <Card>
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <span className="flex items-center gap-2">
              <BarChart2 className="w-4 h-4 text-emerald-600" />
              Station MW Trend
            </span>
          </CardTitle>
        </CardHeader>
        <CardContent>
          {loading ? (
            <div className="py-12 flex justify-center">
              <BlockLoader />
            </div>
          ) : rows.length ? (
            <div className="h-72">
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={buildStationTrend()}>
                  <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                  <XAxis
                    dataKey="period"
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
                  <Legend />
                  <Line
                    type="monotone"
                    dataKey="stationMw"
                    stroke="hsl(var(--primary))"
                    strokeWidth={2}
                    name="Station MW"
                  />
                </LineChart>
              </ResponsiveContainer>
            </div>
          ) : (
            <p className="text-sm text-slate-500">
              Run analytics to view station MW trend.
            </p>
          )}
        </CardContent>
      </Card>

      <Card>
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <span className="flex items-center gap-2">
              <BarChart2 className="w-4 h-4 text-sky-600" />
              Station MVAR Trend
            </span>
          </CardTitle>
        </CardHeader>
        <CardContent>
          {loading ? (
            <div className="py-12 flex justify-center">
              <BlockLoader />
            </div>
          ) : rows.length ? (
            <div className="h-72">
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={buildStationTrend()}>
                  <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                  <XAxis
                    dataKey="period"
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
                  <Legend />
                  <Line
                    type="monotone"
                    dataKey="stationMvar"
                    stroke="hsl(var(--secondary))"
                    strokeWidth={2}
                    name="Station MVAR"
                  />
                </LineChart>
              </ResponsiveContainer>
            </div>
          ) : (
            <p className="text-sm text-slate-500">
              Run analytics to view station MVAR trend.
            </p>
          )}
        </CardContent>
      </Card>

      <Card>
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <span className="flex items-center gap-2">
              <BarChart2 className="w-4 h-4 text-amber-600" />
              Station Load (MW and MVA)
            </span>
          </CardTitle>
        </CardHeader>
        <CardContent>
          {loading ? (
            <div className="py-12 flex justify-center">
              <BlockLoader />
            </div>
          ) : rows.length ? (
            <div className="h-72">
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={buildStationTrend()}>
                  <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                  <XAxis
                    dataKey="period"
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
                  <Legend />
                  <Line
                    type="monotone"
                    dataKey="stationMw"
                    stroke="hsl(var(--primary))"
                    strokeWidth={2}
                    name="Station MW"
                  />
                  <Line
                    type="monotone"
                    dataKey="stationMva"
                    stroke="hsl(var(--accent))"
                    strokeWidth={2}
                    name="Station MVA"
                  />
                </LineChart>
              </ResponsiveContainer>
            </div>
          ) : (
            <p className="text-sm text-slate-500">
              Run analytics to view combined station load trend.
            </p>
          )}
        </CardContent>
      </Card>

      <Card>
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <span className="flex items-center gap-2">
              <BarChart2 className="w-4 h-4 text-blue-600" />
              Station Load Table
            </span>
            <div className="flex flex-col items-end gap-1">
              <span className="text-xs text-slate-500">
                Maximum amps, MW, MVAR and MVA per ICT and aggregated station load
              </span>
              <div className="flex flex-wrap gap-3 items-center justify-end">
                <div className="flex items-center gap-1">
                  <Checkbox
                    id="metric-amps"
                    checked={metricColumns.amps}
                    onCheckedChange={() => toggleMetricColumn('amps')}
                  />
                  <label
                    htmlFor="metric-amps"
                    className="text-xs text-slate-600 dark:text-slate-300"
                  >
                    Amps
                  </label>
                </div>
                <div className="flex items-center gap-1">
                  <Checkbox
                    id="metric-mw"
                    checked={metricColumns.mw}
                    onCheckedChange={() => toggleMetricColumn('mw')}
                  />
                  <label htmlFor="metric-mw" className="text-xs text-slate-600 dark:text-slate-300">
                    MW
                  </label>
                </div>
                <div className="flex items-center gap-1">
                  <Checkbox
                    id="metric-mvar"
                    checked={metricColumns.mvar}
                    onCheckedChange={() => toggleMetricColumn('mvar')}
                  />
                  <label
                    htmlFor="metric-mvar"
                    className="text-xs text-slate-600 dark:text-slate-300"
                  >
                    MVAR
                  </label>
                </div>
                <div className="flex items-center gap-1">
                  <Checkbox
                    id="metric-mva"
                    checked={metricColumns.mva}
                    onCheckedChange={() => toggleMetricColumn('mva')}
                  />
                  <label
                    htmlFor="metric-mva"
                    className="text-xs text-slate-600 dark:text-slate-300"
                  >
                    MVA
                  </label>
                </div>
                <Button
                  variant="outline"
                  size="sm"
                  onClick={handleExport}
                  disabled={!rows.length || loading}
                  className="flex items-center gap-1"
                >
                  <Download className="w-4 h-4" />
                  Export
                </Button>
              </div>
            </div>
          </CardTitle>
        </CardHeader>
        <CardContent>
          {loading ? (
            <div className="py-12 flex justify-center">
              <BlockLoader />
            </div>
          ) : rows.length ? (
            <div className="border rounded-md overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead
                      rowSpan={2}
                      className="whitespace-nowrap border-l border-slate-300"
                    >
                      Date
                    </TableHead>
                    <TableHead
                      rowSpan={2}
                      className="whitespace-nowrap border-l border-slate-300"
                    >
                      Time
                    </TableHead>
                    {sortedFeeders.map((f, idx) => (
                      <TableHead
                        key={`${f.id}-group`}
                        colSpan={activeMetricKeys.length || 1}
                        className="text-center border-l border-slate-300"
                      >
                        {f.name}
                      </TableHead>
                    ))}
                    {stationMetricKeys.length > 0 && (
                      <TableHead
                        colSpan={stationMetricKeys.length}
                        className="text-center border-l border-slate-300"
                      >
                        Station Load
                      </TableHead>
                    )}
                  </TableRow>
                  <TableRow>
                    {sortedFeeders.map(f => (
                      <Fragment key={f.id}>
                        {activeMetricKeys.map((metricKey, idxMetric) => (
                          <TableHead
                            key={`${f.id}-${metricKey}-head`}
                            className={`text-right whitespace-nowrap ${
                              idxMetric === 0 ? 'border-l border-slate-300' : ''
                            }`}
                          >
                            {metricKey === 'amps'
                              ? 'Amps'
                              : metricKey === 'mw'
                              ? 'MW'
                              : metricKey === 'mvar'
                              ? 'MVAR'
                              : 'MVA'}
                          </TableHead>
                        ))}
                      </Fragment>
                    ))}
                    {stationMetricKeys.map((metricKey, idxMetric) => (
                      <TableHead
                        key={`station-${metricKey}-head`}
                        className={`text-right whitespace-nowrap ${
                          idxMetric === 0 ? 'border-l border-slate-300' : ''
                        }`}
                      >
                        {metricKey === 'mva'
                          ? 'MVA'
                          : metricKey === 'mw'
                          ? 'MW'
                          : 'MVAR'}
                      </TableHead>
                    ))}
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {rows.map(row => {
                    const station = row.station || {};
                    const perIct = row.per_ict || {};
                    const { date, time } = getStationDateTime(row);
                    const isMax = maxRowDisplay && row.period === maxRowDisplay.period;
                    return (
                      <TableRow
                        key={row.period}
                        className={`transition-colors ${
                          isMax
                            ? 'bg-yellow-100/80 hover:bg-yellow-100/80 dark:bg-yellow-900/40 dark:hover:bg-yellow-900/40'
                            : 'hover:bg-amber-50/80 dark:hover:bg-slate-800'
                        }`}
                      >
                        <TableCell className="font-mono-data whitespace-nowrap border-l border-slate-300">
                          {date}
                        </TableCell>
                        <TableCell className="font-mono-data whitespace-nowrap border-l border-slate-300">
                          {time}
                        </TableCell>
                        {sortedFeeders.map(f => {
                          const v = perIct[f.id] || {};
                          return (
                            <Fragment key={`${row.period}-${f.id}`}>
                              {activeMetricKeys.map((metricKey, idxMetric) => (
                                <TableCell
                                  key={`${row.period}-${f.id}-${metricKey}`}
                                  className={`text-right font-mono-data ${
                                    idxMetric === 0 ? 'border-l border-slate-300' : ''
                                  }`}
                                >
                                  {metricKey === 'amps'
                                    ? formatInt(v.amps)
                                    : metricKey === 'mw'
                                    ? formatInt(v.mw)
                                    : metricKey === 'mvar'
                                    ? formatDecimal(v.mvar)
                                    : formatDecimal(v.mva)}
                                </TableCell>
                              ))}
                            </Fragment>
                          );
                        })}
                        {stationMetricKeys.map((metricKey, idxMetric) => (
                          <TableCell
                            key={`${row.period}-station-${metricKey}`}
                            className={`text-right font-mono-data ${
                              idxMetric === 0 ? 'border-l border-slate-300' : ''
                            }`}
                          >
                            {metricKey === 'mva'
                              ? formatDecimal(station.mva)
                              : metricKey === 'mw'
                              ? formatInt(station.mw)
                              : formatDecimal(station.mvar)}
                          </TableCell>
                        ))}
                      </TableRow>
                    );
                  })}
                </TableBody>
              </Table>
            </div>
          ) : (
            <p className="text-sm text-slate-500">
              Run analytics to see per-ICT and station load values.
            </p>
          )}
        </CardContent>
      </Card>
      {maxRowDisplay && (
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center justify-between">
              <span className="flex items-center gap-2">
                <BarChart2 className="w-4 h-4 text-red-600" />
                Maximum Station Load Summary
              </span>
            </CardTitle>
          </CardHeader>
          <CardContent>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4 text-sm">
              <div>
                <div className="text-xs text-slate-500">Date</div>
                <div className="font-mono-data">{maxRowDisplay.date}</div>
              </div>
              <div>
                <div className="text-xs text-slate-500">Time</div>
                <div className="font-mono-data">{maxRowDisplay.time}</div>
              </div>
              <div>
                <div className="text-xs text-slate-500">Station MW</div>
                <div className="font-mono-data">
                  {maxRowDisplay.mw != null ? formatInt(maxRowDisplay.mw) : ''}
                </div>
              </div>
              <div>
                <div className="text-xs text-slate-500">Station MVA</div>
                <div className="font-mono-data">
                  {maxRowDisplay.mva != null ? formatDecimal(maxRowDisplay.mva) : ''}
                </div>
              </div>
            </div>
          </CardContent>
        </Card>
      )}
    </div>
  );
}
