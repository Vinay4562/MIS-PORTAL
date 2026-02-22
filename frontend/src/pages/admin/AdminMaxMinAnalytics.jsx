import { useEffect, useState } from 'react';
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

export default function AdminMaxMinAnalytics() {
  const [feeders, setFeeders] = useState([]);
  const [selectedFeeders, setSelectedFeeders] = useState([]);
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [entries, setEntries] = useState([]);
  const [rawEntries, setRawEntries] = useState([]);
  const [loading, setLoading] = useState(false);
  const [viewMode, setViewMode] = useState('day');
  const [yearSummaries, setYearSummaries] = useState({});
  const [columns, setColumns] = useState({
    feeder: true,
    maxMw: true,
    minMw: true,
    avgMw: true,
  });
  const [detailColumns, setDetailColumns] = useState({
    date: true,
    feeder: true,
    maxAmps: true,
    maxMw: true,
    maxMvar: true,
    maxTime: true,
    minAmps: true,
    minMw: true,
    minMvar: true,
    minDate: true,
    minTime: true,
    avgAmps: true,
    avgMw: true,
  });

  useEffect(() => {
    fetchFeeders();
  }, []);

  const syncYearSummaries = async (raw, token) => {
    if (!raw || !raw.length) {
      setYearSummaries({});
      return;
    }
    const combos = {};
    raw.forEach(e => {
      if (!e.date || !e.feeder_id) return;
      const parts = String(e.date).split('-');
      if (parts.length < 2) return;
      const y = parts[0];
      const m = parts[1];
      const key = `${e.feeder_id}__${y}-${m}`;
      if (!combos[key]) {
        combos[key] = { feederId: e.feeder_id, year: y, month: m };
      }
    });
    const results = {};
    const comboList = Object.values(combos);
    await Promise.all(
      comboList.map(async c => {
        try {
          const resp = await axios.get(
            `${API}/max-min/summary/${c.feederId}/${c.year}/${parseInt(c.month, 10)}`,
            token
              ? {
                  headers: { Authorization: `Bearer ${token}` },
                }
              : undefined,
          );
          const data = resp.data || [];
          const full = Array.isArray(data)
            ? data.find(x => x.name === 'Full Month')
            : null;
          if (full) {
            const maxAmps =
              typeof full.max_amps === 'number'
                ? full.max_amps
                : parseFloat(full.max_amps);
            const minAmps =
              typeof full.min_amps === 'number'
                ? full.min_amps
                : parseFloat(full.min_amps);
            const avgAmps =
              typeof full.avg_amps === 'number'
                ? full.avg_amps
                : parseFloat(full.avg_amps);
            const maxMw =
              typeof full.max_mw === 'number'
                ? full.max_mw
                : parseFloat(full.max_mw);
            const minMw =
              typeof full.min_mw === 'number'
                ? full.min_mw
                : parseFloat(full.min_mw);
            const avgMw =
              typeof full.avg_mw === 'number'
                ? full.avg_mw
                : parseFloat(full.avg_mw);
            results[`${c.feederId}__${c.year}-${c.month}`] = {
              feederId: c.feederId,
              year: c.year,
              month: c.month,
              maxAmps: Number.isNaN(maxAmps) ? null : maxAmps,
              minAmps: Number.isNaN(minAmps) ? null : minAmps,
              avgAmps: Number.isNaN(avgAmps) ? null : avgAmps,
              maxMw: Number.isNaN(maxMw) ? null : maxMw,
              minMw: Number.isNaN(minMw) ? null : minMw,
              avgMw: Number.isNaN(avgMw) ? null : avgMw,
              maxTime: full.max_mw_time || full.max_amps_time || null,
              minTime: full.min_mw_time || full.min_amps_time || null,
            };
          }
        } catch (err) {
          console.error('Failed to load max–min monthly summary', err);
        }
      }),
    );
    setYearSummaries(results);
  };

  const fetchFeeders = async () => {
    try {
      const token = localStorage.getItem('token');
      const resp = await axios.get(`${API}/max-min/feeders`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      const onlyFeederTypes = (resp.data || []).filter(f =>
        ['feeder_400kv', 'feeder_220kv', 'ict_feeder'].includes(f.type),
      );
      setFeeders(onlyFeederTypes);
      setSelectedFeeders(onlyFeederTypes.map(f => f.id));
    } catch (e) {
      console.error(e);
      toast.error('Failed to load max–min feeders');
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
      const resp = await axios.get(`${API}/admin/analytics/max-min`, {
        params,
        headers: { Authorization: `Bearer ${token}` },
      });
      const raw = resp.data?.entries || [];
      setRawEntries(raw);
      await syncYearSummaries(raw, token);
      const byFeeder = {};
      raw.forEach(e => {
        const key = e.feeder_id;
        if (!byFeeder[key]) {
          byFeeder[key] = { maxValues: [], minValues: [], avgValues: [] };
        }
        const data = e.data || {};
        if (data.max && typeof data.max.mw === 'number') {
          byFeeder[key].maxValues.push(data.max.mw);
        }
        if (data.min && typeof data.min.mw === 'number') {
          byFeeder[key].minValues.push(data.min.mw);
        }
        if (data.avg && typeof data.avg.mw === 'number') {
          byFeeder[key].avgValues.push(data.avg.mw);
        }
      });
      const rows = Object.keys(byFeeder).map(fid => {
        const stats = byFeeder[fid];
        const feeder = feeders.find(f => f.id === fid) || {};
        const avg = arr => (arr.length ? arr.reduce((s, v) => s + v, 0) / arr.length : 0);
        return {
          feederId: fid,
          feederName: feeder.name || fid,
          maxMw: avg(stats.maxValues),
          minMw: avg(stats.minValues),
          avgMw: avg(stats.avgValues),
        };
      });
      rows.sort((a, b) => {
        const ai = getFeederOrderIndex(a.feederId);
        const bi = getFeederOrderIndex(b.feederId);
        if (ai !== bi) return ai - bi;
        if (a.feederName < b.feederName) return -1;
        if (a.feederName > b.feederName) return 1;
        return 0;
      });
      setEntries(rows);
    } catch (e) {
      console.error(e);
      toast.error('Failed to load max–min analytics');
    } finally {
      setLoading(false);
    }
  };

  const handleExport = async () => {
    try {
      const rows = buildDetailRows();
      if (!rows.length) {
        toast.error('No data to export');
        return;
      }

      const token = localStorage.getItem('token');
      const columnDefs = [
        { key: 'feeder', label: 'Feeder', field: 'feederName' },
        { key: 'maxAmps', label: 'Max Amps', field: 'maxAmps' },
        { key: 'maxMw', label: 'Max MW', field: 'maxMw' },
        { key: 'date', label: 'Date', field: 'date' },
        { key: 'maxTime', label: 'Max Time', field: 'maxTime' },
        { key: 'maxMvar', label: 'Max MVAR', field: 'maxMvar' },
        { key: 'minAmps', label: 'Min Amps', field: 'minAmps' },
        { key: 'minMw', label: 'Min MW', field: 'minMw' },
        { key: 'minMvar', label: 'Min MVAR', field: 'minMvar' },
        { key: 'minDate', label: 'Min Date', field: 'minDate' },
        { key: 'minTime', label: 'Min Time', field: 'minTime' },
        { key: 'avgAmps', label: 'Avg Amps', field: 'avgAmps' },
        { key: 'avgMw', label: 'Avg MW', field: 'avgMw' },
      ];

      const activeColumns = columnDefs.filter(col => {
        if (col.key === 'feeder') return detailColumns.feeder;
        if (col.key === 'date') return detailColumns.date;
        if (col.key === 'maxAmps') return detailColumns.maxAmps;
        if (col.key === 'maxMw') return detailColumns.maxMw;
        if (col.key === 'maxMvar') return detailColumns.maxMvar;
        if (col.key === 'maxTime') return detailColumns.maxTime;
        if (col.key === 'minAmps') return detailColumns.minAmps;
        if (col.key === 'minMw') return detailColumns.minMw;
        if (col.key === 'minMvar') return detailColumns.minMvar;
        if (col.key === 'minDate') return detailColumns.minDate;
        if (col.key === 'minTime') return detailColumns.minTime;
        if (col.key === 'avgAmps') return detailColumns.avgAmps;
        if (col.key === 'avgMw') return detailColumns.avgMw;
        return false;
      });

      if (!activeColumns.length) {
        toast.error('Select at least one column to export');
        return;
      }

      const payload = {
        columns: activeColumns,
        rows,
        meta: {
          startDate,
          endDate,
          viewMode,
        },
      };

      const resp = await axios.post(`${API}/admin/analytics/max-min/export-view`, payload, {
        responseType: 'blob',
        headers: token
          ? {
              Authorization: `Bearer ${token}`,
            }
          : undefined,
      });

      const filenameParts = ['Admin_MaxMin_Analytics'];
      if (startDate) filenameParts.push(startDate);
      if (endDate) filenameParts.push(endDate);
      const filename = `${filenameParts.join('_')}.xlsx`;

      await downloadFile(resp.data, filename);
      toast.success('Export completed successfully');
    } catch (e) {
      console.error(e);
      toast.error('Failed to export data');
    }
  };

  const handleToggleAllFeeders = () => {
    setSelectedFeeders(prev =>
      prev.length === feeders.length ? [] : feeders.map(f => f.id),
    );
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

  const getFeederOrderIndex = (id) => {
    const idx = feeders.findIndex(f => f.id === id);
    return idx === -1 ? Number.MAX_SAFE_INTEGER : idx;
  };

  const formatDateDdMmYyyy = (value) => {
    if (!value) return '';
    const dt = new Date(value);
    if (Number.isNaN(dt.getTime())) return value;
    const dd = String(dt.getDate()).padStart(2, '0');
    const mm = String(dt.getMonth() + 1).padStart(2, '0');
    const yyyy = dt.getFullYear();
    return `${dd}-${mm}-${yyyy}`;
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

  const formatDecimal = (value) => {
    if (typeof value !== 'number') return value ?? '';
    return value.toFixed(2);
  };

  const buildMwTrend = () => {
    if (!rawEntries.length) return [];

    if (viewMode === 'year') {
      const groups = {};
      const hasSummaries =
        yearSummaries && Object.keys(yearSummaries).length > 0;

      if (hasSummaries) {
        Object.values(yearSummaries).forEach(s => {
          const year = parseInt(s.year, 10);
          const month = parseInt(s.month, 10);
          if (!year || !month) return;
          const key = `${year}-${String(month).padStart(2, '0')}`;
          const dt = new Date(year, month - 1, 1);
          const label = dt.toLocaleDateString(undefined, {
            month: 'short',
            year: 'numeric',
          });

          if (!groups[key]) {
            groups[key] = {
              label,
              maxMwMax: null,
              minMwMin: null,
              avgMwSum: 0,
              avgCount: 0,
            };
          }

          const max =
            typeof s.maxMw === 'number' && !Number.isNaN(s.maxMw)
              ? s.maxMw
              : null;
          const min =
            typeof s.minMw === 'number' && !Number.isNaN(s.minMw)
              ? s.minMw
              : null;
          const avg =
            typeof s.avgMw === 'number' && !Number.isNaN(s.avgMw)
              ? s.avgMw
              : null;

          if (max !== null) {
            if (groups[key].maxMwMax === null || max > groups[key].maxMwMax) {
              groups[key].maxMwMax = max;
            }
          }
          if (min !== null) {
            if (groups[key].minMwMin === null || min < groups[key].minMwMin) {
              groups[key].minMwMin = min;
            }
          }
          if (avg !== null) {
            groups[key].avgMwSum += avg;
            groups[key].avgCount += 1;
          }
        });
      } else {
        rawEntries.forEach(e => {
          if (!e.date) return;
          const d = e.data || {};
          const max =
            d.max && typeof d.max.mw === 'number' ? d.max.mw : null;
          const min =
            d.min && typeof d.min.mw === 'number' ? d.min.mw : null;
          const avg =
            d.avg && typeof d.avg.mw === 'number' ? d.avg.mw : null;
          if (max === null && min === null && avg === null) return;

          const dt = new Date(e.date);
          if (Number.isNaN(dt.getTime())) return;
          const year = dt.getFullYear();
          const month = dt.getMonth() + 1;
          const key = `${year}-${String(month).padStart(2, '0')}`;
          const label = dt.toLocaleDateString(undefined, {
            month: 'short',
            year: 'numeric',
          });

          if (!groups[key]) {
            groups[key] = {
              label,
              maxMwMax: null,
              minMwMin: null,
              avgMwSum: 0,
              avgCount: 0,
            };
          }
          if (max !== null) {
            if (groups[key].maxMwMax === null || max > groups[key].maxMwMax) {
              groups[key].maxMwMax = max;
            }
          }
          if (min !== null) {
            if (groups[key].minMwMin === null || min < groups[key].minMwMin) {
              groups[key].minMwMin = min;
            }
          }
          if (avg !== null) {
            groups[key].avgMwSum += avg;
            groups[key].avgCount += 1;
          }
        });
      }

      return Object.entries(groups)
        .map(([key, r]) => ({
          key,
          date: r.label,
          maxMw: r.maxMwMax ?? 0,
          minMw: r.minMwMin ?? 0,
          avgMw: r.avgCount ? r.avgMwSum / r.avgCount : 0,
        }))
        .sort((a, b) => {
          if (a.key < b.key) return -1;
          if (a.key > b.key) return 1;
          return 0;
        })
        .map(r => ({
          date: r.date,
          maxMw: r.maxMw,
          minMw: r.minMw,
          avgMw: r.avgMw,
        }));
    }

    // Day-wise and Month-wise: show daily data within the selected range
    const byDate = {};
    rawEntries.forEach(e => {
      if (!e.date) return;
      const d = e.data || {};
      const max = d.max && typeof d.max.mw === 'number' ? d.max.mw : null;
      const min = d.min && typeof d.min.mw === 'number' ? d.min.mw : null;
      const avg = d.avg && typeof d.avg.mw === 'number' ? d.avg.mw : null;
      if (max === null && min === null && avg === null) return;
      const label = formatDateDdMmYyyy(e.date);
      if (!byDate[label]) {
        byDate[label] = { maxMwSum: 0, minMwSum: 0, avgMwSum: 0, count: 0 };
      }
      if (max !== null) byDate[label].maxMwSum += max;
      if (min !== null) byDate[label].minMwSum += min;
      if (avg !== null) byDate[label].avgMwSum += avg;
      byDate[label].count += 1;
    });

    return Object.entries(byDate)
      .map(([label, r]) => ({
        date: label,
        maxMw: r.count ? r.maxMwSum / r.count : 0,
        minMw: r.count ? r.minMwSum / r.count : 0,
        avgMw: r.count ? r.avgMwSum / r.count : 0,
      }))
      .sort((a, b) => {
        if (a.date < b.date) return -1;
        if (a.date > b.date) return 1;
        return 0;
      });
  };

  const buildDetailRows = () => {
    if (viewMode !== 'year') {
      return rawEntries
        .map(e => {
          const feeder = feeders.find(f => f.id === e.feeder_id) || {};
          const data = e.data || {};
          const max = data.max || {};
          const min = data.min || {};
          const avg = data.avg || {};
          return {
            id: `${e.feeder_id || ''}-${e.date || ''}`,
            date: e.date ? formatDateDdMmYyyy(e.date) : '',
            minDate: e.date ? formatDateDdMmYyyy(e.date) : '',
            sortDate: e.date || '',
            feederId: e.feeder_id || '',
            feederName: feeder.name || e.feeder_id || '',
            maxAmps: max.amps,
            maxMw: max.mw,
            maxMvar: max.mvar,
            maxTime: max.time,
            minAmps: min.amps,
            minMw: min.mw,
            minMvar: min.mvar,
            minTime: min.time,
            avgAmps: avg.amps,
            avgMw: avg.mw,
          };
        })
        .filter(row => row.date && row.feederName)
        .sort((a, b) => {
          const allSelected =
            feeders.length > 0 && selectedFeeders.length === feeders.length;
          const da = a.sortDate ? new Date(a.sortDate) : null;
          const db = b.sortDate ? new Date(b.sortDate) : null;
          if (allSelected) {
            const ai = getFeederOrderIndex(a.feederId);
            const bi = getFeederOrderIndex(b.feederId);
            if (ai !== bi) return ai - bi;
            if (da && db) {
              if (da < db) return -1;
              if (da > db) return 1;
            } else if (da && !db) {
              return -1;
            } else if (!da && db) {
              return 1;
            }
          } else {
            if (da && db) {
              if (da < db) return -1;
              if (da > db) return 1;
            } else if (da && !db) {
              return -1;
            } else if (!da && db) {
              return 1;
            }
            const ai = getFeederOrderIndex(a.feederId);
            const bi = getFeederOrderIndex(b.feederId);
            if (ai !== bi) return ai - bi;
          }
          if (a.feederName < b.feederName) return -1;
          if (a.feederName > b.feederName) return 1;
          return 0;
        })
        .map(({ sortDate, feederId, ...rest }) => rest);
    }

    const groups = {};
    rawEntries.forEach(e => {
      if (!e.date || !e.feeder_id) return;
      const feeder = feeders.find(f => f.id === e.feeder_id) || {};
      const data = e.data || {};
      const max = data.max || {};
      const min = data.min || {};
      const avg = data.avg || {};

      const dt = new Date(e.date);
      if (Number.isNaN(dt.getTime())) return;
      const year = dt.getFullYear();
      const month = dt.getMonth() + 1;
      const monthKey = `${year}-${String(month).padStart(2, '0')}`;
      const key = `${e.feeder_id}__${monthKey}`;
      const monthDateIso = `${year}-${String(month).padStart(2, '0')}-01`;
      const dateLabel = formatDateDdMmYyyy(monthDateIso);

      if (!groups[key]) {
        groups[key] = {
          id: key,
          feederId: e.feeder_id,
          dateLabel,
          feederName: feeder.name || e.feeder_id || '',
          maxAmps: null,
          maxMw: null,
          maxMvar: null,
          maxDate: null,
          minDate: null,
          sortDateIso: monthDateIso,
          minAmps: null,
          minMw: null,
          minMvar: null,
          maxTime: null,
          minTime: null,
          avgAmpsSum: 0,
          avgAmpsCount: 0,
          avgMwSum: 0,
          avgMwCount: 0,
        };
      }

      if (typeof max.amps === 'number') {
        if (groups[key].maxAmps === null || max.amps > groups[key].maxAmps) {
          groups[key].maxAmps = max.amps;
        }
      }
      if (typeof max.mw === 'number') {
        if (groups[key].maxMw === null || max.mw > groups[key].maxMw) {
          groups[key].maxMw = max.mw;
          groups[key].maxDate = e.date;
          groups[key].sortDateIso = e.date;
        }
      }
      if (typeof max.mvar === 'number') {
        if (groups[key].maxMvar === null || max.mvar > groups[key].maxMvar) {
          groups[key].maxMvar = max.mvar;
        }
      }
      if (max.time !== undefined && max.time !== null) {
        if (groups[key].maxTime === null || (typeof max.mw === 'number' && max.mw === groups[key].maxMw)) {
          groups[key].maxTime = max.time;
        }
      }

      if (typeof min.amps === 'number') {
        if (groups[key].minAmps === null || min.amps < groups[key].minAmps) {
          groups[key].minAmps = min.amps;
        }
      }
      if (typeof min.mw === 'number') {
        if (groups[key].minMw === null || min.mw < groups[key].minMw) {
          groups[key].minMw = min.mw;
          groups[key].minDate = e.date;
        }
      }
      if (typeof min.mvar === 'number') {
        if (groups[key].minMvar === null || min.mvar < groups[key].minMvar) {
          groups[key].minMvar = min.mvar;
        }
      }
      if (min.time !== undefined && min.time !== null) {
        if (groups[key].minTime === null || (typeof min.mw === 'number' && min.mw === groups[key].minMw)) {
          groups[key].minTime = min.time;
        }
      }

      if (typeof avg.amps === 'number') {
        groups[key].avgAmpsSum += avg.amps;
        groups[key].avgAmpsCount += 1;
      }
      if (typeof avg.mw === 'number') {
        groups[key].avgMwSum += avg.mw;
        groups[key].avgMwCount += 1;
      }
    });

    return Object.values(groups)
      .map(g => {
        const summary = yearSummaries[g.id];
        const dateValue = g.maxDate ? formatDateDdMmYyyy(g.maxDate) : g.dateLabel;
        const minDateValue = g.minDate ? formatDateDdMmYyyy(g.minDate) : '';
        const sortDate = g.maxDate || g.sortDateIso || null;
        const avgAmps =
          summary && typeof summary.avgAmps === 'number'
            ? summary.avgAmps
            : g.avgAmpsCount
            ? g.avgAmpsSum / g.avgAmpsCount
            : null;
        const avgMw =
          summary && typeof summary.avgMw === 'number'
            ? summary.avgMw
            : g.avgMwCount
            ? g.avgMwSum / g.avgMwCount
            : null;

        return {
          id: g.id,
          feederId: g.feederId,
          date: dateValue,
          minDate: minDateValue,
          sortDate,
          feederName: g.feederName,
          maxAmps:
            summary &&
            summary.maxAmps !== null &&
            summary.maxAmps !== undefined
              ? summary.maxAmps
              : g.maxAmps,
          maxMw:
            summary &&
            summary.maxMw !== null &&
            summary.maxMw !== undefined
              ? summary.maxMw
              : g.maxMw,
          maxMvar: g.maxMvar,
          maxTime:
            summary && summary.maxTime
              ? summary.maxTime
              : g.maxTime,
          minAmps:
            summary &&
            summary.minAmps !== null &&
            summary.minAmps !== undefined
              ? summary.minAmps
              : g.minAmps,
          minMw:
            summary &&
            summary.minMw !== null &&
            summary.minMw !== undefined
              ? summary.minMw
              : g.minMw,
          minMvar: g.minMvar,
          minTime:
            summary && summary.minTime
              ? summary.minTime
              : g.minTime,
          avgAmps,
          avgMw,
        };
      })
      .sort((a, b) => {
        const allSelected =
          feeders.length > 0 && selectedFeeders.length === feeders.length;
        const da = a.sortDate ? new Date(a.sortDate) : null;
        const db = b.sortDate ? new Date(b.sortDate) : null;
        if (allSelected) {
          const ai = getFeederOrderIndex(a.feederId);
          const bi = getFeederOrderIndex(b.feederId);
          if (ai !== bi) return ai - bi;
          if (da && db) {
            if (da < db) return -1;
            if (da > db) return 1;
          } else if (da && !db) {
            return -1;
          } else if (!da && db) {
            return 1;
          }
        } else {
          if (da && db) {
            if (da < db) return -1;
            if (da > db) return 1;
          } else if (da && !db) {
            return -1;
          } else if (!da && db) {
            return 1;
          }
          const ai = getFeederOrderIndex(a.feederId);
          const bi = getFeederOrderIndex(b.feederId);
          if (ai !== bi) return ai - bi;
        }
        if (a.feederName < b.feederName) return -1;
        if (a.feederName > b.feederName) return 1;
        return 0;
      })
      .map(({ sortDate, feederId, ...rest }) => rest);
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row gap-4 md:items-end justify-between">
        <div className="space-y-2">
          <h1 className="text-2xl md:text-3xl font-heading font-bold text-slate-900 dark:text-slate-50">
            Max–Min Analytics
          </h1>
          <p className="text-sm text-slate-500 dark:text-slate-400">
            Envelope view of maximum and minimum loading across critical feeders.
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
            <p className="text-xs font-medium text-slate-500">View</p>
            <div className="inline-flex rounded-full border border-slate-200 dark:border-slate-700 p-0.5 bg-slate-50/60 dark:bg-slate-900/60">
              <button
                type="button"
                onClick={() => setViewMode('day')}
                className={`px-3 py-1.5 text-xs rounded-full ${
                  viewMode === 'day'
                    ? 'bg-slate-900 text-white'
                    : 'text-slate-600 dark:text-slate-300'
                }`}
              >
                Day-wise
              </button>
              <button
                type="button"
                onClick={() => setViewMode('month')}
                className={`px-3 py-1.5 text-xs rounded-full ${
                  viewMode === 'month'
                    ? 'bg-slate-900 text-white'
                    : 'text-slate-600 dark:text-slate-300'
                }`}
              >
                Month-wise
              </button>
              <button
                type="button"
                onClick={() => setViewMode('year')}
                className={`px-3 py-1.5 text-xs rounded-full ${
                  viewMode === 'year'
                    ? 'bg-slate-900 text-white'
                    : 'text-slate-600 dark:text-slate-300'
                }`}
              >
                Year-wise
              </button>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <Button onClick={handleFetch} disabled={loading}>
              {loading ? <BlockLoader /> : 'Load Analytics'}
            </Button>
            <Button
              variant="secondary"
              onClick={handleExport}
              disabled={loading || !rawEntries.length}
            >
              <Download className="w-4 h-4 mr-2" />
              Export
            </Button>
          </div>
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
            {feeders.length > 0 && (
              <button
                type="button"
                onClick={handleToggleAllFeeders}
                className={`px-3 py-1.5 rounded-full text-xs border ${
                  selectedFeeders.length === feeders.length && feeders.length
                    ? 'bg-slate-900 text-white border-slate-900'
                    : 'bg-slate-50 dark:bg-slate-900 text-slate-700 dark:text-slate-300 border-slate-200 dark:border-slate-700'
                }`}
              >
                {selectedFeeders.length === feeders.length && feeders.length
                  ? 'Deselect All'
                  : 'Select All'}
              </button>
            )}
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
              <BarChart2 className="w-4 h-4 text-emerald-600" />
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
            <div className="h-72">
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={buildMwTrend()}>
                  <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                  <XAxis
                    dataKey="date"
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
                    dataKey="maxMw"
                    stroke="hsl(var(--primary))"
                    strokeWidth={2}
                    name="Avg Max MW"
                  />
                  <Line
                    type="monotone"
                    dataKey="minMw"
                    stroke="hsl(var(--secondary))"
                    strokeWidth={2}
                    name="Avg Min MW"
                  />
                  <Line
                    type="monotone"
                    dataKey="avgMw"
                    stroke="hsl(var(--muted-foreground))"
                    strokeWidth={2}
                    name="Avg MW"
                  />
                </LineChart>
              </ResponsiveContainer>
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
              <BarChart2 className="w-4 h-4 text-blue-600" />
              Loading Summary
            </span>
            <div className="flex gap-4 items-center">
              <div className="flex items-center gap-2">
                <Checkbox
                  id="col-max-mw"
                  checked={columns.maxMw}
                  onCheckedChange={() => toggleColumn('maxMw')}
                />
                <label htmlFor="col-max-mw" className="text-xs text-slate-600 dark:text-slate-300">
                  Max MW
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="col-min-mw"
                  checked={columns.minMw}
                  onCheckedChange={() => toggleColumn('minMw')}
                />
                <label htmlFor="col-min-mw" className="text-xs text-slate-600 dark:text-slate-300">
                  Min MW
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="col-avg-mw"
                  checked={columns.avgMw}
                  onCheckedChange={() => toggleColumn('avgMw')}
                />
                <label htmlFor="col-avg-mw" className="text-xs text-slate-600 dark:text-slate-300">
                  Avg MW
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
                    {columns.maxMw && <TableHead className="text-right">Avg Max MW</TableHead>}
                    {columns.minMw && <TableHead className="text-right">Avg Min MW</TableHead>}
                    {columns.avgMw && <TableHead className="text-right">Avg MW</TableHead>}
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {entries.map(row => (
                    <TableRow key={row.feederId}>
                      {columns.feeder && <TableCell>{row.feederName}</TableCell>}
                      {columns.maxMw && (
                        <TableCell className="text-right font-mono-data">
                          {row.maxMw.toFixed(2)}
                        </TableCell>
                      )}
                      {columns.minMw && (
                        <TableCell className="text-right font-mono-data">
                          {row.minMw.toFixed(2)}
                        </TableCell>
                      )}
                      {columns.avgMw && (
                        <TableCell className="text-right font-mono-data">
                          {row.avgMw.toFixed(2)}
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

      <Card>
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <span className="flex items-center gap-2">
              <BarChart2 className="w-4 h-4 text-amber-600" />
              Detailed Max–Min Entries
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
                  id="detail-max-amps"
                  checked={detailColumns.maxAmps}
                  onCheckedChange={() => toggleDetailColumn('maxAmps')}
                />
                <label htmlFor="detail-max-amps" className="text-xs text-slate-600 dark:text-slate-300">
                  Max Amps
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-max-mw"
                  checked={detailColumns.maxMw}
                  onCheckedChange={() => toggleDetailColumn('maxMw')}
                />
                <label htmlFor="detail-max-mw" className="text-xs text-slate-600 dark:text-slate-300">
                  Max MW
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-max-mvar"
                  checked={detailColumns.maxMvar}
                  onCheckedChange={() => toggleDetailColumn('maxMvar')}
                />
                <label htmlFor="detail-max-mvar" className="text-xs text-slate-600 dark:text-slate-300">
                  Max MVAR
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-min-amps"
                  checked={detailColumns.minAmps}
                  onCheckedChange={() => toggleDetailColumn('minAmps')}
                />
                <label htmlFor="detail-min-amps" className="text-xs text-slate-600 dark:text-slate-300">
                  Min Amps
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-min-mw"
                  checked={detailColumns.minMw}
                  onCheckedChange={() => toggleDetailColumn('minMw')}
                />
                <label htmlFor="detail-min-mw" className="text-xs text-slate-600 dark:text-slate-300">
                  Min MW
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-min-mvar"
                  checked={detailColumns.minMvar}
                  onCheckedChange={() => toggleDetailColumn('minMvar')}
                />
                <label htmlFor="detail-min-mvar" className="text-xs text-slate-600 dark:text-slate-300">
                  Min MVAR
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-min-date"
                  checked={detailColumns.minDate}
                  onCheckedChange={() => toggleDetailColumn('minDate')}
                />
                <label htmlFor="detail-min-date" className="text-xs text-slate-600 dark:text-slate-300">
                  Min Date
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-max-time"
                  checked={detailColumns.maxTime}
                  onCheckedChange={() => toggleDetailColumn('maxTime')}
                />
                <label htmlFor="detail-max-time" className="text-xs text-slate-600 dark:text-slate-300">
                  Max Time
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-min-time"
                  checked={detailColumns.minTime}
                  onCheckedChange={() => toggleDetailColumn('minTime')}
                />
                <label htmlFor="detail-min-time" className="text-xs text-slate-600 dark:text-slate-300">
                  Min Time
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-avg-amps"
                  checked={detailColumns.avgAmps}
                  onCheckedChange={() => toggleDetailColumn('avgAmps')}
                />
                <label htmlFor="detail-avg-amps" className="text-xs text-slate-600 dark:text-slate-300">
                  Avg Amps
                </label>
              </div>
              <div className="flex items-center gap-2">
                <Checkbox
                  id="detail-avg-mw"
                  checked={detailColumns.avgMw}
                  onCheckedChange={() => toggleDetailColumn('avgMw')}
                />
                <label htmlFor="detail-avg-mw" className="text-xs text-slate-600 dark:text-slate-300">
                  Avg MW
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
                    {detailColumns.feeder && <TableHead>Feeder</TableHead>}
                    {detailColumns.maxAmps && <TableHead className="text-right">Max Amps</TableHead>}
                    {detailColumns.maxMw && <TableHead className="text-right">Max MW</TableHead>}
                    {detailColumns.date && <TableHead>Date</TableHead>}
                    {detailColumns.maxTime && <TableHead className="text-right">Max Time</TableHead>}
                    {detailColumns.maxMvar && <TableHead className="text-right">Max MVAR</TableHead>}
                    {detailColumns.minAmps && <TableHead className="text-right">Min Amps</TableHead>}
                    {detailColumns.minMw && <TableHead className="text-right">Min MW</TableHead>}
                    {detailColumns.minMvar && <TableHead className="text-right">Min MVAR</TableHead>}
                    {detailColumns.minDate && <TableHead className="text-right">Min Date</TableHead>}
                    {detailColumns.minTime && <TableHead className="text-right">Min Time</TableHead>}
                    {detailColumns.avgAmps && <TableHead className="text-right">Avg Amps</TableHead>}
                    {detailColumns.avgMw && <TableHead className="text-right">Avg MW</TableHead>}
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {buildDetailRows().map(row => (
                    <TableRow key={row.id}>
                      {detailColumns.feeder && <TableCell>{row.feederName}</TableCell>}
                      {detailColumns.maxAmps && (
                        <TableCell className="text-right font-mono-data">
                          {row.maxAmps ?? ''}
                        </TableCell>
                      )}
                      {detailColumns.maxMw && (
                        <TableCell className="text-right font-mono-data">
                          {row.maxMw ?? ''}
                        </TableCell>
                      )}
                      {detailColumns.date && <TableCell>{row.date}</TableCell>}
                      {detailColumns.maxTime && (
                        <TableCell className="text-right font-mono-data">
                          {formatTimeDisplay(row.maxTime)}
                        </TableCell>
                      )}
                      {detailColumns.maxMvar && (
                        <TableCell className="text-right font-mono-data">
                          {row.maxMvar ?? ''}
                        </TableCell>
                      )}
                      {detailColumns.minAmps && (
                        <TableCell className="text-right font-mono-data">
                          {row.minAmps ?? ''}
                        </TableCell>
                      )}
                      {detailColumns.minMw && (
                        <TableCell className="text-right font-mono-data">
                          {row.minMw ?? ''}
                        </TableCell>
                      )}
                      {detailColumns.minMvar && (
                        <TableCell className="text-right font-mono-data">
                          {row.minMvar ?? ''}
                        </TableCell>
                      )}
                      {detailColumns.minDate && (
                        <TableCell className="text-right font-mono-data">
                          {row.minDate ?? ''}
                        </TableCell>
                      )}
                      {detailColumns.minTime && (
                        <TableCell className="text-right font-mono-data">
                          {formatTimeDisplay(row.minTime)}
                        </TableCell>
                      )}
                      {detailColumns.avgAmps && (
                        <TableCell className="text-right font-mono-data">
                          {formatDecimal(row.avgAmps)}
                        </TableCell>
                      )}
                      {detailColumns.avgMw && (
                        <TableCell className="text-right font-mono-data">
                          {formatDecimal(row.avgMw)}
                        </TableCell>
                      )}
                    </TableRow>
                  ))}
                  {!buildDetailRows().length && (
                    <TableRow>
                      <TableCell colSpan={12} className="text-center text-sm text-slate-500 py-6">
                        No detailed entries available for the selected period and feeders.
                      </TableCell>
                    </TableRow>
                  )}
                </TableBody>
              </Table>
            </div>
          ) : (
            <p className="text-sm text-slate-500">
              Run analytics to view detailed max–min entries.
            </p>
          )}
        </CardContent>
      </Card>
    </div>
  );
}
