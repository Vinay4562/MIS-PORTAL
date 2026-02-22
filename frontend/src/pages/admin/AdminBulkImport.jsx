import { useEffect, useState } from 'react';
import axios from 'axios';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Checkbox } from '@/components/ui/checkbox';
import { Textarea } from '@/components/ui/textarea';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { ScrollArea } from '@/components/ui/scroll-area';
import { BlockLoader } from '@/components/ui/loader';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter } from '@/components/ui/dialog';
import { toast } from 'sonner';
import { Database, Activity, TrendingDown, BarChart2, Zap } from 'lucide-react';
import { downloadFile, formatDate } from '@/lib/utils';

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

const MODULES = [
  {
    id: 'energy',
    label: 'Energy',
    description: 'Boundary meter energy entries from sheets',
    icon: Activity,
  },
  {
    id: 'line-losses',
    label: 'Line Losses',
    description: 'Feeder-wise daily energy and loss entries',
    icon: TrendingDown,
  },
  {
    id: 'max-min',
    label: 'Max–Min',
    description: 'Daily max–min voltage and load snapshots',
    icon: BarChart2,
  },
  {
    id: 'station-load',
    label: 'Station Load (ICT)',
    description: 'ICT max–min inputs used for Station Load analytics',
    icon: BarChart2,
  },
  {
    id: 'interruptions',
    label: 'Interruptions',
    description: 'Outage events across all eligible feeders',
    icon: Zap,
  },
];

function formatDurationMinutes(value) {
  if (value == null || value === '') return '';
  const minutes = parseFloat(value);
  if (Number.isNaN(minutes)) return String(value);
  const total = Math.round(minutes);
  const dayMinutes = 24 * 60;
  const days = Math.floor(total / dayMinutes);
  const remainder = total % dayMinutes;
  const h = Math.floor(remainder / 60);
  const m = remainder % 60;
  const timePart = `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
  if (days <= 0) return timePart;
  if (days === 1) return `1 Day ${timePart}`;
  return `${days} Days ${timePart}`;
}

const ENDPOINT_MAP = {
  energy: '/admin/bulk-import/energy',
  'line-losses': '/admin/bulk-import/line-losses',
  'max-min': '/admin/bulk-import/max-min',
  'station-load': '/admin/bulk-import/max-min',
  interruptions: '/admin/bulk-import/interruptions',
};

export default function AdminBulkImport() {
  const now = new Date();
  const [moduleId, setModuleId] = useState('line-losses');
  const [periodType, setPeriodType] = useState('monthly');
  const [year, setYear] = useState(String(now.getFullYear()));
  const [month, setMonth] = useState(String(now.getMonth() + 1).padStart(2, '0'));
  const [overwrite, setOverwrite] = useState(false);
  const [source, setSource] = useState('json');
  const [rawPayload, setRawPayload] = useState('');
  const [file, setFile] = useState(null);
  const [feederId, setFeederId] = useState('');
  const [sheetId, setSheetId] = useState('');
  const [feeders, setFeeders] = useState([]);
  const [sheets, setSheets] = useState([]);
  const [maxMinFeeders, setMaxMinFeeders] = useState([]);
  const [maxMinFeederId, setMaxMinFeederId] = useState('');
  const [interruptFeederId, setInterruptFeederId] = useState('');
  const ictFeeders = maxMinFeeders.filter(f => f.type === 'ict_feeder');
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [confirmReplaceOpen, setConfirmReplaceOpen] = useState(false);
  const [confirmContext, setConfirmContext] = useState('');
  const [pendingImportConfig, setPendingImportConfig] = useState(null);
  const [previewOpen, setPreviewOpen] = useState(false);
  const [previewRows, setPreviewRows] = useState([]);
  const [previewModule, setPreviewModule] = useState(null);
  const [previewMeta, setPreviewMeta] = useState(null);
  const [previewLoading, setPreviewLoading] = useState(false);

  useEffect(() => {
    const bootstrap = async () => {
      const token = localStorage.getItem('token');
      if (!token) return;
      try {
        const [feederResp, sheetResp, maxMinResp] = await Promise.all([
          axios.get(`${API}/feeders`, {
            headers: { Authorization: `Bearer ${token}` },
          }),
          axios.get(`${API}/energy/sheets`, {
            headers: { Authorization: `Bearer ${token}` },
          }),
          axios.get(`${API}/max-min/feeders`, {
            headers: { Authorization: `Bearer ${token}` },
          }),
        ]);
        setFeeders(feederResp.data || []);
        setSheets(sheetResp.data || []);
        setMaxMinFeeders(maxMinResp.data || []);
      } catch (e) {
        console.error(e);
      }
    };
    bootstrap();
  }, []);

  const performImport = async (overwriteFlag) => {
    const yearInt = parseInt(year, 10);
    const monthInt = parseInt(month, 10);

    try {
      setLoading(true);
      setResult(null);
      const token = localStorage.getItem('token');
      let resp;

      if (source === 'excel') {
        if (!file) {
          toast.error('Select an Excel file to upload');
          return;
        }
        let baseUrl = '';
        if (moduleId === 'line-losses') {
          if (!feederId.trim()) {
            toast.error('Select a feeder for Line Losses import');
            return;
          }
          baseUrl = `${API}/admin/bulk-import/line-losses/excel/${encodeURIComponent(feederId.trim())}`;
        } else if (moduleId === 'energy') {
          if (!sheetId.trim()) {
            toast.error('Select a sheet for Energy import');
            return;
          }
          baseUrl = `${API}/admin/bulk-import/energy/excel/${encodeURIComponent(sheetId.trim())}`;
        } else if (moduleId === 'max-min' || moduleId === 'station-load') {
          if (!maxMinFeederId.trim()) {
            toast.error(moduleId === 'station-load' ? 'Select an ICT feeder for Station Load import' : 'Select a feeder for Max–Min import');
            return;
          }
          baseUrl = `${API}/admin/bulk-import/max-min/excel/${encodeURIComponent(maxMinFeederId.trim())}`;
        } else if (moduleId === 'interruptions') {
          if (!interruptFeederId.trim()) {
            toast.error('Select a feeder for Interruptions import');
            return;
          }
          baseUrl = `${API}/admin/bulk-import/interruptions/excel/${encodeURIComponent(interruptFeederId.trim())}`;
        } else {
          toast.error('Excel import is not supported for this module');
          return;
        }
        const formData = new FormData();
        formData.append('file', file);
        const params = new URLSearchParams();
        if (!Number.isNaN(yearInt)) {
          params.append('year', String(yearInt));
        }
        if (periodType === 'monthly' && !Number.isNaN(monthInt)) {
          params.append('month', String(monthInt));
        }
        params.append('overwrite', overwriteFlag ? 'true' : 'false');
        const url = params.toString() ? `${baseUrl}?${params.toString()}` : baseUrl;
        resp = await axios.post(url, formData, {
          headers: {
            Authorization: `Bearer ${token}`,
          },
        });
      } else {
        if (!rawPayload.trim()) {
          toast.error('Paste a JSON payload for the selected module');
          return;
        }
        let basePayload;
        try {
          basePayload = JSON.parse(rawPayload);
        } catch (e) {
          toast.error('Invalid JSON payload. Please check the syntax.');
          return;
        }

        const endpoint = ENDPOINT_MAP[moduleId];
        if (!endpoint) {
          toast.error('Unsupported module selected for bulk import');
          return;
        }

        const payload = {
          ...basePayload,
          overwrite: overwriteFlag,
        };

        if (!Number.isNaN(yearInt)) {
          payload.year = yearInt;
        }
        if (periodType === 'monthly' && !Number.isNaN(monthInt)) {
          payload.month = monthInt;
        } else if (periodType === 'yearly') {
          payload.month = null;
        }

        resp = await axios.post(`${API}${endpoint}`, payload, {
          headers: {
            Authorization: `Bearer ${token}`,
          },
        });
      }

      setResult(resp.data || null);
      toast.success('Bulk import completed');
    } catch (e) {
      console.error(e);
      const detail = e?.response?.data?.detail || 'Bulk import failed';
      toast.error(typeof detail === 'string' ? detail : 'Bulk import failed');
    } finally {
      setLoading(false);
      setPendingImportConfig(null);
    }
  };

  const handleRunImport = async () => {
    const yearInt = parseInt(year, 10);
    const monthInt = parseInt(month, 10);

    try {
      const yearIntValid = !Number.isNaN(yearInt) ? yearInt : null;
      const monthIntValid =
        periodType === 'monthly' && !Number.isNaN(monthInt) ? monthInt : null;

      let checkUrl = '';
      let checkPayload = {};

      if (source === 'excel') {
        if (moduleId === 'line-losses') {
          if (!feederId.trim()) {
            toast.error('Select a feeder for Line Losses import');
            return;
          }
          checkUrl = `${API}/admin/bulk-import/check/line-losses`;
          checkPayload = {
            feeder_id: feederId.trim(),
            year: yearIntValid,
            month: monthIntValid,
          };
        } else if (moduleId === 'energy') {
          if (!sheetId.trim()) {
            toast.error('Select a sheet for Energy import');
            return;
          }
          checkUrl = `${API}/admin/bulk-import/check/energy`;
          checkPayload = {
            sheet_id: sheetId.trim(),
            year: yearIntValid,
            month: monthIntValid,
          };
        } else if (moduleId === 'max-min' || moduleId === 'station-load') {
          if (!maxMinFeederId.trim()) {
            toast.error(
              moduleId === 'station-load'
                ? 'Select an ICT feeder for Station Load import'
                : 'Select a feeder for Max–Min import',
            );
            return;
          }
          checkUrl = `${API}/admin/bulk-import/check/max-min`;
          checkPayload = {
            feeder_id: maxMinFeederId.trim(),
            year: yearIntValid,
            month: monthIntValid,
          };
        } else if (moduleId === 'interruptions') {
          if (!interruptFeederId.trim()) {
            toast.error('Select a feeder for Interruptions import');
            return;
          }
          checkUrl = `${API}/admin/bulk-import/check/interruptions`;
          checkPayload = {
            feeder_ids: [interruptFeederId.trim()],
            year: yearIntValid,
            month: monthIntValid,
          };
        }
      } else {
        if (!rawPayload.trim()) {
          toast.error('Paste a JSON payload for the selected module');
          return;
        }
        let basePayload;
        try {
          basePayload = JSON.parse(rawPayload);
        } catch (e) {
          toast.error('Invalid JSON payload. Please check the syntax.');
          return;
        }

        if (moduleId === 'line-losses') {
          if (!basePayload.feeder_id) {
            toast.error('Payload must include feeder_id for Line Losses');
            return;
          }
          checkUrl = `${API}/admin/bulk-import/check/line-losses`;
          checkPayload = {
            feeder_id: basePayload.feeder_id,
            year: yearIntValid,
            month: monthIntValid,
          };
        } else if (moduleId === 'energy') {
          if (!basePayload.sheet_id) {
            toast.error('Payload must include sheet_id for Energy');
            return;
          }
          checkUrl = `${API}/admin/bulk-import/check/energy`;
          checkPayload = {
            sheet_id: basePayload.sheet_id,
            year: yearIntValid,
            month: monthIntValid,
          };
        } else if (moduleId === 'max-min' || moduleId === 'station-load') {
          if (!basePayload.feeder_id) {
            toast.error('Payload must include feeder_id for Max–Min/Station Load');
            return;
          }
          checkUrl = `${API}/admin/bulk-import/check/max-min`;
          checkPayload = {
            feeder_id: basePayload.feeder_id,
            year: yearIntValid,
            month: monthIntValid,
          };
        } else if (moduleId === 'interruptions') {
          const entries = Array.isArray(basePayload.entries) ? basePayload.entries : [];
          const feederIds = Array.from(
            new Set(entries.map(e => e.feeder_id).filter(Boolean)),
          );
          if (!feederIds.length) {
            toast.error('Payload must include entries with feeder_id for Interruptions');
            return;
          }
          checkUrl = `${API}/admin/bulk-import/check/interruptions`;
          checkPayload = {
            feeder_ids: feederIds,
            year: yearIntValid,
            month: monthIntValid,
          };
        }
      }

      if (!checkUrl) {
        await performImport(overwrite);
        return;
      }

      const token = localStorage.getItem('token');
      const respCheck = await axios.post(checkUrl, checkPayload, {
        headers: {
          Authorization: `Bearer ${token}`,
        },
      });
      const data = respCheck.data || {};
      if (data.has_existing) {
        const labelParts = [];
        if (yearIntValid) {
          labelParts.push(String(yearIntValid));
        }
        if (monthIntValid) {
          labelParts.push(monthIntValid.toString().padStart(2, '0'));
        }
        const periodLabel = labelParts.length ? labelParts.join('-') : 'selected period';
        const context =
          moduleId === 'energy'
            ? `Existing Energy entries detected for sheet and ${periodLabel}.`
            : moduleId === 'line-losses'
            ? `Existing Line Losses entries detected for feeder and ${periodLabel}.`
            : moduleId === 'max-min' || moduleId === 'station-load'
            ? `Existing Max–Min entries detected for feeder and ${periodLabel}.`
            : `Existing Interruption entries detected for selected feeders and ${periodLabel}.`;
        setConfirmContext(context);
        setPendingImportConfig({ overwriteFlag: true });
        setConfirmReplaceOpen(true);
      } else {
        await performImport(overwrite);
      }
    } catch (e) {
      console.error(e);
      const detail = e?.response?.data?.detail || 'Bulk import failed';
      toast.error(typeof detail === 'string' ? detail : 'Bulk import failed');
    }
  };

  const handleDownloadTemplate = async () => {
    const yearInt = parseInt(year, 10);
    const monthInt = parseInt(month, 10);
    if (Number.isNaN(yearInt) || Number.isNaN(monthInt)) {
      toast.error('Enter a valid year and month to download a template');
      return;
    }
    const token = localStorage.getItem('token');
    try {
      let url = '';
      let filename = '';
      if (moduleId === 'line-losses') {
        if (!feederId.trim()) {
          toast.error('Select a feeder for Line Losses');
          return;
        }
        const trimmedId = feederId.trim();
        url = `${API}/export/${encodeURIComponent(trimmedId)}/${yearInt}/${monthInt}`;
        const feeder = feeders.find(f => f.id === trimmedId);
        filename = `${feeder?.name || 'LineLosses'}_${yearInt}_${monthInt.toString().padStart(2, '0')}.xlsx`;
      } else if (moduleId === 'energy') {
        if (!sheetId.trim()) {
          toast.error('Select a sheet for Energy');
          return;
        }
        const trimmedId = sheetId.trim();
        url = `${API}/energy/export/${encodeURIComponent(trimmedId)}/${yearInt}/${monthInt}`;
        const sheet = sheets.find(s => (s.id || s._id) === trimmedId);
        filename = `${sheet?.name || 'Energy'}_${yearInt}_${monthInt.toString().padStart(2, '0')}.xlsx`;
      } else if (moduleId === 'max-min') {
        if (!maxMinFeederId.trim()) {
          toast.error('Select a feeder for Max–Min');
          return;
        }
        const trimmedId = maxMinFeederId.trim();
        url = `${API}/max-min/export/${encodeURIComponent(trimmedId)}/${yearInt}/${monthInt}`;
        const feeder = maxMinFeeders.find(f => f.id === trimmedId);
        filename = `${feeder?.name || 'MaxMin'}_${yearInt}_${monthInt.toString().padStart(2, '0')}.xlsx`;
      } else if (moduleId === 'station-load') {
        if (!maxMinFeederId.trim()) {
          toast.error('Select an ICT feeder for Station Load');
          return;
        }
        const trimmedId = maxMinFeederId.trim();
        url = `${API}/max-min/export/${encodeURIComponent(trimmedId)}/${yearInt}/${monthInt}`;
        const feeder = ictFeeders.find(f => f.id === trimmedId);
        filename = `StationLoad_${feeder?.name || 'ICT'}_${yearInt}_${monthInt.toString().padStart(2, '0')}.xlsx`;
      } else if (moduleId === 'interruptions') {
        if (!interruptFeederId.trim()) {
          toast.error('Select a feeder for Interruptions');
          return;
        }
        const trimmedId = interruptFeederId.trim();
        url = `${API}/interruptions/export/${encodeURIComponent(trimmedId)}/${yearInt}/${monthInt}`;
        const feeder = maxMinFeeders.find(f => f.id === trimmedId);
        filename = `Interruptions_${feeder?.name || 'Feeder'}_${yearInt}_${monthInt.toString().padStart(2, '0')}.xlsx`;
      } else {
        toast.error('Template download not available for this module');
        return;
      }
      const response = await axios.get(url, {
        responseType: 'blob',
        headers: token
          ? {
              Authorization: `Bearer ${token}`,
            }
          : undefined,
      });
      await downloadFile(response.data, filename);
      toast.success('Template downloaded');
    } catch (e) {
      console.error(e);
      toast.error('Failed to download template');
    }
  };

  const currentModule = MODULES.find(m => m.id === moduleId);

  const handleApplyMaxMinTemplate = () => {
    if (!maxMinFeederId) {
      toast.error('Select a feeder for Max–Min');
      return;
    }
    const template = {
      feeder_id: maxMinFeederId,
      entries: [
        {
          date: '2026-01-01',
          data: {},
        },
      ],
    };
    setRawPayload(JSON.stringify(template, null, 2));
  };

  const handleApplyInterruptionsTemplate = () => {
    if (!interruptFeederId) {
      toast.error('Select a feeder for Interruptions');
      return;
    }
    const template = {
      entries: [
        {
          feeder_id: interruptFeederId,
          date: '2026-01-01',
          start_time: '00:00',
          end_time: '00:30',
          duration_minutes: 30,
          description: '',
        },
      ],
    };
    setRawPayload(JSON.stringify(template, null, 2));
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row gap-4 md:items-end justify-between">
        <div className="space-y-2">
          <h1 className="text-2xl md:text-3xl font-heading font-bold text-slate-900 dark:text-slate-50">
            Bulk Import Console
          </h1>
          <p className="text-sm text-slate-500 dark:text-slate-400">
            Orchestrate month-wise and year-wise imports across energy, line losses, max–min and interruptions.
          </p>
        </div>
        <div className="flex items-center gap-2 text-xs text-slate-500 dark:text-slate-400">
          <Database className="w-4 h-4" />
          <span>Admin-only orchestration · Uses secure /api/admin/bulk-import/** endpoints</span>
        </div>
      </div>

      <Card>
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <span>Module & Period</span>
            <span className="text-xs text-slate-500">
              Select target module and time horizon for this run
            </span>
          </CardTitle>
        </CardHeader>
        <CardContent className="space-y-6">
          <div className="space-y-3">
            <p className="text-xs font-medium text-slate-500">Module</p>
            <div className="flex flex-wrap gap-3">
              {MODULES.map(m => {
                const Icon = m.icon;
                const active = moduleId === m.id;
                return (
                  <button
                    key={m.id}
                    type="button"
                    onClick={() => setModuleId(m.id)}
                    className={`px-4 py-3 rounded-xl border flex items-start gap-3 text-left transition-colors ${
                      active
                        ? 'bg-slate-900 text-white border-slate-900'
                        : 'bg-white/70 dark:bg-slate-900/70 text-slate-800 dark:text-slate-100 border-slate-200 dark:border-slate-700 hover:bg-slate-50 dark:hover:bg-slate-800'
                    }`}
                  >
                    <div className={`mt-1 rounded-full p-1.5 ${active ? 'bg-slate-800' : 'bg-slate-100 dark:bg-slate-800'}`}>
                      <Icon className="w-4 h-4" />
                    </div>
                    <div>
                      <div className="text-sm font-semibold">{m.label}</div>
                      <div className="text-xs text-slate-500 dark:text-slate-400">
                        {m.description}
                      </div>
                    </div>
                  </button>
                );
              })}
            </div>
          </div>

          <div className="grid gap-4 md:grid-cols-3">
            <div className="space-y-2">
              <p className="text-xs font-medium text-slate-500">Period Type</p>
              <div className="inline-flex rounded-full border border-slate-200 dark:border-slate-700 p-0.5 bg-slate-50/60 dark:bg-slate-900/60">
                <button
                  type="button"
                  onClick={() => setPeriodType('monthly')}
                  className={`px-3 py-1.5 text-xs rounded-full ${
                    periodType === 'monthly'
                      ? 'bg-slate-900 text-white'
                      : 'text-slate-600 dark:text-slate-300'
                  }`}
                >
                  Monthly
                </button>
                <button
                  type="button"
                  onClick={() => setPeriodType('yearly')}
                  className={`px-3 py-1.5 text-xs rounded-full ${
                    periodType === 'yearly'
                      ? 'bg-slate-900 text-white'
                      : 'text-slate-600 dark:text-slate-300'
                  }`}
                >
                  Yearly
                </button>
              </div>
            </div>
            <div className="space-y-2">
              <label className="text-xs font-medium text-slate-500">Year</label>
              <Input
                type="number"
                min="2000"
                max="2100"
                value={year}
                onChange={e => setYear(e.target.value)}
                className="max-w-[160px]"
              />
            </div>
            {periodType === 'monthly' && (
              <div className="space-y-2">
                <label className="text-xs font-medium text-slate-500">Month</label>
                <Input
                  type="number"
                  min="1"
                  max="12"
                  value={month}
                  onChange={e => setMonth(e.target.value)}
                  className="max-w-[160px]"
                />
              </div>
            )}
          </div>

          <div className="flex items-center gap-2">
            <Checkbox
              id="overwrite"
              checked={overwrite}
              onCheckedChange={val => setOverwrite(Boolean(val))}
            />
            <label htmlFor="overwrite" className="text-xs text-slate-600 dark:text-slate-300">
              Overwrite existing entries that match the same date (and start time for interruptions)
            </label>
          </div>
        </CardContent>
      </Card>

      {confirmReplaceOpen && (
        <div className="fixed inset-0 z-40 flex items-center justify-center bg-black/40">
          <div className="bg-white dark:bg-slate-900 rounded-lg shadow-xl max-w-md w-full p-6 space-y-4">
            <h2 className="text-lg font-semibold text-slate-900 dark:text-slate-50">
              Existing data detected
            </h2>
            <p className="text-sm text-slate-600 dark:text-slate-300">
              {confirmContext || 'Existing data was found for the selected period. How would you like to proceed?'}
            </p>
            <p className="text-xs text-slate-500 dark:text-slate-400">
              Cancel will stop this import run. Replace will overwrite matching existing records
              with the new data while keeping all other entries unchanged.
            </p>
            <div className="flex justify-end gap-2 pt-2">
              <Button
                variant="outline"
                size="sm"
                onClick={() => {
                  setConfirmReplaceOpen(false);
                  setPendingImportConfig(null);
                }}
              >
                Cancel
              </Button>
              <Button
                size="sm"
                onClick={async () => {
                  setConfirmReplaceOpen(false);
                  setOverwrite(true);
                  const cfg = pendingImportConfig || { overwriteFlag: true };
                  await performImport(cfg.overwriteFlag);
                }}
              >
                Replace
              </Button>
            </div>
          </div>
        </div>
      )}

      <Card>
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <span>Payload</span>
            <span className="text-xs text-slate-500">
              Choose JSON or Excel source for {currentModule?.label} imports
            </span>
          </CardTitle>
        </CardHeader>
        <CardContent className="space-y-4">
          <p className="text-xs text-slate-500 dark:text-slate-400">
            Use JSON to call the underlying bulk-import APIs directly or switch to Excel to upload
            module files that are parsed and mapped using existing import rules.
          </p>
          <div className="flex items-center gap-2 text-xs">
            <button
              type="button"
              onClick={() => setSource('json')}
              className={`px-3 py-1.5 rounded-full border ${
                source === 'json'
                  ? 'bg-slate-900 text-white border-slate-900'
                  : 'bg-slate-50 dark:bg-slate-900 text-slate-700 dark:text-slate-200 border-slate-200 dark:border-slate-700'
              }`}
            >
              JSON payload
            </button>
            <button
              type="button"
              onClick={() => setSource('excel')}
              className={`px-3 py-1.5 rounded-full border ${
                source === 'excel'
                  ? 'bg-slate-900 text-white border-slate-900'
                  : 'bg-slate-50 dark:bg-slate-900 text-slate-700 dark:text-slate-200 border-slate-200 dark:border-slate-700'
              }`}
            >
              Excel file
            </button>
          </div>

          {source === 'json' ? (
            <>
              {moduleId === 'max-min' && (
                <div className="flex items-center gap-2 text-xs">
                  <select
                    className="flex-1 rounded-md border border-slate-200 bg-white px-2 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-900"
                    value={maxMinFeederId}
                    onChange={e => setMaxMinFeederId(e.target.value)}
                  >
                    <option value="">Select Max–Min feeder</option>
                      {maxMinFeeders.map(f => (
                      <option key={f.id} value={f.id}>
                        {f.name}
                      </option>
                    ))}
                  </select>
                  <Button
                    type="button"
                    variant="outline"
                    size="sm"
                    onClick={handleApplyMaxMinTemplate}
                  >
                    Use template
                  </Button>
                </div>
              )}
              {moduleId === 'station-load' && (
                <div className="flex items-center gap-2 text-xs">
                  <select
                    className="flex-1 rounded-md border border-slate-200 bg-white px-2 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-900"
                    value={maxMinFeederId}
                    onChange={e => setMaxMinFeederId(e.target.value)}
                  >
                    <option value="">Select ICT feeder</option>
                    {ictFeeders.map(f => (
                      <option key={f.id} value={f.id}>
                        {f.name}
                      </option>
                    ))}
                  </select>
                  <Button
                    type="button"
                    variant="outline"
                    size="sm"
                    onClick={handleApplyMaxMinTemplate}
                  >
                    Use template
                  </Button>
                </div>
              )}
              {moduleId === 'interruptions' && (
                <div className="flex items-center gap-2 text-xs">
                  <select
                    className="flex-1 rounded-md border border-slate-200 bg-white px-2 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-900"
                    value={interruptFeederId}
                    onChange={e => setInterruptFeederId(e.target.value)}
                  >
                    <option value="">Select feeder for interruption entry</option>
                    {maxMinFeeders
                      .filter(f =>
                        ['feeder_400kv', 'feeder_220kv', 'ict_feeder', 'reactor_feeder', 'bay_feeder'].includes(
                          f.type,
                        ),
                      )
                      .map(f => (
                        <option key={f.id} value={f.id}>
                          {f.name}
                        </option>
                      ))}
                  </select>
                  <Button
                    type="button"
                    variant="outline"
                    size="sm"
                    onClick={handleApplyInterruptionsTemplate}
                  >
                    Use template
                  </Button>
                </div>
              )}
              <Textarea
                rows={10}
                value={rawPayload}
                onChange={e => setRawPayload(e.target.value)}
                placeholder={
                  moduleId === 'line-losses'
                    ? '{"feeder_id": \"...\", \"entries\": [{\"date\": \"2026-01-01\", \"end1_import_final\": 0, ...}]}'
                    : moduleId === 'energy'
                      ? '{"sheet_id": \"...\", \"entries\": [{\"date\": \"2026-01-01\", \"readings\": [...]}]}'
                      : moduleId === 'max-min'
                        ? '{"feeder_id": \"...\", \"entries\": [{\"date\": \"2026-01-01\", \"data\": {...}}]}'
                        : '{"entries\": [{\"feeder_id\": \"...\", \"date\": \"2026-01-01\", \"start_time\": \"00:00\", ...}]}'
                }
                className="font-mono text-xs"
              />
            </>
          ) : moduleId === 'line-losses' || moduleId === 'energy' || moduleId === 'max-min' || moduleId === 'station-load' || moduleId === 'interruptions' ? (
            <div className="space-y-4">
              <div className="grid gap-4 md:grid-cols-2">
                {moduleId === 'line-losses' && (
                  <div className="space-y-2">
                    <label className="text-xs font-medium text-slate-500">Feeder</label>
                    <div className="flex gap-2">
                      <select
                        className="flex-1 rounded-md border border-slate-200 bg-white px-2 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-900"
                        value={feederId}
                        onChange={e => setFeederId(e.target.value)}
                      >
                        <option value="">Select feeder</option>
                        {feeders.map(f => (
                          <option key={f.id} value={f.id}>
                            {f.name}
                          </option>
                        ))}
                      </select>
                      <Input
                        value={feederId}
                        onChange={e => setFeederId(e.target.value)}
                        placeholder="Or paste feeder_id"
                        className="w-40 text-xs"
                      />
                    </div>
                  </div>
                )}
                {moduleId === 'energy' && (
                  <div className="space-y-2">
                    <label className="text-xs font-medium text-slate-500">Sheet</label>
                    <div className="flex gap-2">
                      <select
                        className="flex-1 rounded-md border border-slate-200 bg-white px-2 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-900"
                        value={sheetId}
                        onChange={e => setSheetId(e.target.value)}
                      >
                        <option value="">Select sheet</option>
                        {sheets.map(s => (
                          <option key={s.id || s._id} value={s.id || s._id}>
                            {s.name}
                          </option>
                        ))}
                      </select>
                      <Input
                        value={sheetId}
                        onChange={e => setSheetId(e.target.value)}
                        placeholder="Or paste sheet_id"
                        className="w-40 text-xs"
                      />
                    </div>
                  </div>
                )}
                {moduleId === 'max-min' && (
                  <div className="space-y-2">
                    <label className="text-xs font-medium text-slate-500">Feeder</label>
                    <div className="flex gap-2">
                      <select
                        className="flex-1 rounded-md border border-slate-200 bg-white px-2 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-900"
                        value={maxMinFeederId}
                        onChange={e => setMaxMinFeederId(e.target.value)}
                      >
                        <option value="">Select feeder</option>
                        {maxMinFeeders.map(f => (
                          <option key={f.id} value={f.id}>
                            {f.name}
                          </option>
                        ))}
                      </select>
                      <Input
                        value={maxMinFeederId}
                        onChange={e => setMaxMinFeederId(e.target.value)}
                        placeholder="Or paste feeder_id"
                        className="w-40 text-xs"
                      />
                    </div>
                  </div>
                )}
                {moduleId === 'station-load' && (
                  <div className="space-y-2">
                    <label className="text-xs font-medium text-slate-500">ICT feeder</label>
                    <div className="flex gap-2">
                      <select
                        className="flex-1 rounded-md border border-slate-200 bg-white px-2 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-900"
                        value={maxMinFeederId}
                        onChange={e => setMaxMinFeederId(e.target.value)}
                      >
                        <option value="">Select ICT feeder</option>
                        {ictFeeders.map(f => (
                          <option key={f.id} value={f.id}>
                            {f.name}
                          </option>
                        ))}
                      </select>
                      <Input
                        value={maxMinFeederId}
                        onChange={e => setMaxMinFeederId(e.target.value)}
                        placeholder="Or paste feeder_id"
                        className="w-40 text-xs"
                      />
                    </div>
                  </div>
                )}
                {moduleId === 'interruptions' && (
                  <div className="space-y-2">
                    <label className="text-xs font-medium text-slate-500">Feeder</label>
                    <div className="flex gap-2">
                      <select
                        className="flex-1 rounded-md border border-slate-200 bg-white px-2 py-1.5 text-xs dark:border-slate-700 dark:bg-slate-900"
                        value={interruptFeederId}
                        onChange={e => setInterruptFeederId(e.target.value)}
                      >
                        <option value="">Select feeder</option>
                        {maxMinFeeders
                          .filter(f =>
                            ['feeder_400kv', 'feeder_220kv', 'ict_feeder', 'reactor_feeder', 'bay_feeder'].includes(
                              f.type,
                            ),
                          )
                          .map(f => (
                            <option key={f.id} value={f.id}>
                              {f.name}
                            </option>
                          ))}
                      </select>
                      <Input
                        value={interruptFeederId}
                        onChange={e => setInterruptFeederId(e.target.value)}
                        placeholder="Or paste feeder_id"
                        className="w-40 text-xs"
                      />
                    </div>
                  </div>
                )}
                <div className="space-y-2">
                  <div className="flex items-center justify-between">
                    <label className="text-xs font-medium text-slate-500">Excel File</label>
                    <button
                      type="button"
                      onClick={handleDownloadTemplate}
                      className="text-[11px] text-indigo-600 hover:underline"
                    >
                      Download template
                    </button>
                  </div>
                  <Input
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={e => setFile(e.target.files?.[0] || null)}
                  />
                </div>
              </div>
              <p className="text-[11px] text-slate-500 dark:text-slate-400">
                Files are parsed using the same column-matching rules as the regular module imports,
                then imported via the admin bulk-import engine with per-month statistics and errors.
                Use Download template to get the correct column layout for this module.
              </p>
            </div>
          ) : (
            <p className="text-xs text-slate-500 dark:text-slate-400">
              Excel-based import is available for all modules shown in this console.
            </p>
          )}
          <div className="flex justify-end gap-2">
            {source === 'excel' && (
              <Button
                type="button"
                variant="outline"
                onClick={async () => {
                  const yearInt = parseInt(year, 10);
                  const monthInt = parseInt(month, 10);
                  const yearVal = !Number.isNaN(yearInt) ? yearInt : null;
                  const monthVal =
                    periodType === 'monthly' && !Number.isNaN(monthInt) ? monthInt : null;
                  if (!file) {
                    toast.error('Select an Excel file to preview');
                    return;
                  }
                  const token = localStorage.getItem('token');
                  if (!token) {
                    toast.error('Authentication token is missing');
                    return;
                  }

                  let baseUrl = '';
                  if (moduleId === 'line-losses') {
                    if (!feederId.trim()) {
                      toast.error('Select a feeder for Line Losses');
                      return;
                    }
                    baseUrl = `${API}/admin/bulk-import/line-losses/excel-preview/${encodeURIComponent(
                      feederId.trim(),
                    )}`;
                  } else if (moduleId === 'energy') {
                    if (!sheetId.trim()) {
                      toast.error('Select a sheet for Energy');
                      return;
                    }
                    baseUrl = `${API}/admin/bulk-import/energy/excel-preview/${encodeURIComponent(
                      sheetId.trim(),
                    )}`;
                  } else if (moduleId === 'max-min' || moduleId === 'station-load') {
                    if (!maxMinFeederId.trim()) {
                      toast.error(
                        moduleId === 'station-load'
                          ? 'Select an ICT feeder for Station Load'
                          : 'Select a feeder for Max–Min',
                      );
                      return;
                    }
                    baseUrl = `${API}/admin/bulk-import/max-min/excel-preview/${encodeURIComponent(
                      maxMinFeederId.trim(),
                    )}`;
                  } else if (moduleId === 'interruptions') {
                    if (!interruptFeederId.trim()) {
                      toast.error('Select a feeder for Interruptions');
                      return;
                    }
                    baseUrl = `${API}/admin/bulk-import/interruptions/excel-preview/${encodeURIComponent(
                      interruptFeederId.trim(),
                    )}`;
                  } else {
                    toast.error('Excel preview is not supported for this module');
                    return;
                  }

                  const formData = new FormData();
                  formData.append('file', file);
                  const params = new URLSearchParams();
                  if (!Number.isNaN(yearInt)) {
                    params.append('year', String(yearInt));
                  }
                  if (periodType === 'monthly' && !Number.isNaN(monthInt)) {
                    params.append('month', String(monthInt));
                  }
                  const url = params.toString() ? `${baseUrl}?${params.toString()}` : baseUrl;

                  try {
                    setPreviewLoading(true);
                    const resp = await axios.post(url, formData, {
                      headers: {
                        Authorization: `Bearer ${token}`,
                      },
                    });
                    const rows = Array.isArray(resp.data) ? resp.data : [];
                    setPreviewRows(rows);
                    setPreviewModule(moduleId);
                    if (moduleId === 'max-min' || moduleId === 'station-load') {
                      const feeder = maxMinFeeders.find(f => f.id === maxMinFeederId.trim());
                      setPreviewMeta({
                        year: yearVal,
                        month: monthVal,
                        feederType: feeder?.type || null,
                        feederName: feeder?.name || '',
                      });
                    } else if (moduleId === 'line-losses') {
                      const feeder = feeders.find(f => f.id === feederId.trim());
                      setPreviewMeta({
                        year: yearVal,
                        month: monthVal,
                        feederName: feeder?.name || '',
                      });
                    } else if (moduleId === 'energy') {
                      const sheet = sheets.find(s => (s.id || s._id) === sheetId.trim());
                      setPreviewMeta({
                        year: yearVal,
                        month: monthVal,
                        sheetName: sheet?.name || '',
                      });
                    } else if (moduleId === 'interruptions') {
                      const feeder = maxMinFeeders.find(f => f.id === interruptFeederId.trim());
                      setPreviewMeta({
                        year: yearVal,
                        month: monthVal,
                        feederName: feeder?.name || '',
                      });
                    } else {
                      setPreviewMeta(null);
                    }
                    setPreviewOpen(true);
                  } catch (e) {
                    console.error(e);
                    const detail = e?.response?.data?.detail || 'Failed to preview Excel import';
                    toast.error(typeof detail === 'string' ? detail : 'Failed to preview Excel import');
                  } finally {
                    setPreviewLoading(false);
                  }
                }}
                disabled={loading || previewLoading}
              >
                {previewLoading ? 'Previewing...' : 'Preview Excel'}
              </Button>
            )}
            <Button onClick={handleRunImport} disabled={loading || previewLoading}>
              {loading ? (
                <span className="inline-flex items-center gap-2">
                  <BlockLoader />
                  Running bulk import...
                </span>
              ) : (
                'Run Bulk Import'
              )}
            </Button>
          </div>
        </CardContent>
      </Card>

      <Card>
        <CardHeader>
          <CardTitle>Run Summary</CardTitle>
        </CardHeader>
        <CardContent>
          {!result ? (
            <p className="text-sm text-slate-500">
              No bulk import run yet. Configure a module, paste payload and execute a run to see
              per-month statistics and detailed error reporting.
            </p>
          ) : (
            <div className="space-y-6">
              <div className="grid gap-4 md:grid-cols-4">
                <div className="space-y-1">
                  <p className="text-xs text-slate-500">Module</p>
                  <p className="text-sm font-medium text-slate-900 dark:text-slate-50">
                    {result.module || currentModule?.label}
                  </p>
                </div>
                <div className="space-y-1">
                  <p className="text-xs text-slate-500">Year</p>
                  <p className="text-sm font-medium text-slate-900 dark:text-slate-50">
                    {result.year || 'Auto-detected'}
                  </p>
                </div>
                <div className="space-y-1">
                  <p className="text-xs text-slate-500">Scope</p>
                  <p className="text-sm font-medium text-slate-900 dark:text-slate-50">
                    {periodType === 'monthly' ? 'Monthly' : 'Yearly'}
                  </p>
                </div>
                <div className="space-y-1">
                  <p className="text-xs text-slate-500">Overwrite Mode</p>
                  <p className="text-sm font-medium text-slate-900 dark:text-slate-50">
                    {result.overwrite ? 'Overwrite existing entries' : 'Skip existing entries'}
                  </p>
                </div>
              </div>

              <div className="grid gap-4 md:grid-cols-3">
                <StatCard label="Total Entries" value={result.total_entries} />
                <StatCard label="Inserted" value={result.inserted} tone="success" />
                <StatCard label="Skipped / Validation Errors" value={(result.skipped_existing || 0) + (result.validation_errors || 0)} tone="warning" />
              </div>

              <div className="space-y-3">
                <p className="text-xs font-medium text-slate-500">Per-Month Breakdown</p>
                <div className="border rounded-md overflow-hidden bg-white dark:bg-slate-950">
                  <Table>
                    <TableHeader>
                      <TableRow>
                        <TableHead>Month</TableHead>
                        <TableHead className="text-right">Inserted</TableHead>
                        <TableHead className="text-right">Skipped Existing</TableHead>
                        <TableHead className="text-right">Validation Errors</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {(result.per_month || []).map(row => (
                        <TableRow key={row.month}>
                          <TableCell>{row.month}</TableCell>
                          <TableCell className="text-right font-mono-data">{row.inserted}</TableCell>
                          <TableCell className="text-right font-mono-data">{row.skipped_existing}</TableCell>
                          <TableCell className="text-right font-mono-data">{row.validation_errors}</TableCell>
                        </TableRow>
                      ))}
                      {(!result.per_month || !result.per_month.length) && (
                        <TableRow>
                          <TableCell colSpan={4} className="text-center text-sm text-slate-500 py-4">
                            No monthly breakdown available.
                          </TableCell>
                        </TableRow>
                      )}
                    </TableBody>
                  </Table>
                </div>
              </div>

              <div className="space-y-3">
                <p className="text-xs font-medium text-slate-500">Detailed Errors</p>
                <ScrollArea className="h-56 border rounded-md bg-white dark:bg-slate-950">
                  <Table>
                    <TableHeader>
                      <TableRow>
                        <TableHead className="w-16">Index</TableHead>
                        <TableHead>Reason</TableHead>
                        <TableHead>Meta</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {(result.errors || []).map((err, idx) => (
                        <TableRow key={idx}>
                          <TableCell className="text-xs font-mono-data">{err.index ?? '-'}</TableCell>
                          <TableCell className="text-xs">{err.reason}</TableCell>
                          <TableCell className="text-[11px] text-slate-500 max-w-[540px] truncate">
                            {JSON.stringify(
                              {
                                date: err.date,
                                feeder_id: err.feeder_id,
                                sheet_id: err.sheet_id,
                              },
                            )}
                          </TableCell>
                        </TableRow>
                      ))}
                      {(!result.errors || !result.errors.length) && (
                        <TableRow>
                          <TableCell colSpan={3} className="text-center text-sm text-slate-500 py-4">
                            No validation errors recorded for this run.
                          </TableCell>
                        </TableRow>
                      )}
                    </TableBody>
                  </Table>
                </ScrollArea>
              </div>
            </div>
          )}
        </CardContent>
      </Card>

      <Dialog
        open={previewOpen}
        onOpenChange={open => {
          setPreviewOpen(open);
        }}
      >
        <DialogContent className="max-w-5xl max-h-[90vh] flex flex-col">
          <DialogHeader>
            <DialogTitle>
              {previewModule === 'line-losses'
                ? 'Line Losses Excel Preview'
                : previewModule === 'energy'
                  ? 'Energy Excel Preview'
                  : previewModule === 'max-min'
                    ? 'Max–Min Excel Preview'
                    : previewModule === 'station-load'
                      ? 'Station Load Excel Preview'
                      : 'Interruptions Excel Preview'}
            </DialogTitle>
            {previewMeta && (
              <p className="mt-1 text-xs text-slate-500">
                {previewModule === 'energy' && previewMeta.sheetName
                  ? `Sheet: ${previewMeta.sheetName}`
                  : previewMeta.feederName
                    ? `Feeder: ${previewMeta.feederName}`
                    : null}
                {typeof previewMeta.year === 'number' &&
                  ` • Year: ${previewMeta.year}`}
                {typeof previewMeta.month === 'number' &&
                  ` • Month: ${String(previewMeta.month).padStart(2, '0')}`}
              </p>
            )}
          </DialogHeader>
          <div className="flex-1 overflow-hidden flex flex-col min-h-0">
            <div className="text-sm text-slate-500 mb-2 flex justify-between shrink-0">
              <span>
                Found {previewRows.length} records.{' '}
                {previewRows.filter(r => r.exists).length > 0
                  ? `${previewRows.filter(r => r.exists).length} existing records have existing entries for this period.`
                  : ''}
              </span>
            </div>
            <ScrollArea className="h-[60vh] border rounded-md">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>Date</TableHead>
                    {previewModule === 'interruptions' && (
                      <>
                        <TableHead>Start Time</TableHead>
                        <TableHead>End Time</TableHead>
                        <TableHead>Duration</TableHead>
                        <TableHead>Cause / Description</TableHead>
                      </>
                    )}
                    {previewModule === 'line-losses' && (
                      <>
                        <TableHead>End1 Import</TableHead>
                        <TableHead>End1 Export</TableHead>
                        <TableHead>End2 Import</TableHead>
                        <TableHead>End2 Export</TableHead>
                      </>
                    )}
                    {previewModule === 'energy' && (
                      <>
                        <TableHead>Readings Count</TableHead>
                      </>
                    )}
                    {(previewModule === 'max-min' || previewModule === 'station-load') && (
                      <>
                        <TableHead>Max</TableHead>
                        <TableHead>Min</TableHead>
                        <TableHead>Avg</TableHead>
                      </>
                    )}
                    <TableHead>Status</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {previewRows.map((row, idx) => {
                    if (previewModule === 'interruptions') {
                      const endDateStr = row.end_date || row.date;
                      const endTimeStr = row.end_time || '';
                      const toDisplay =
                        !endTimeStr
                          ? ''
                          : endDateStr === row.date
                            ? endTimeStr
                            : `${formatDate(endDateStr)} ${endTimeStr}`;
                      const causeBase = row.cause_of_interruption || row.description || '';
                      return (
                        <TableRow
                          key={idx}
                          className={row.exists ? 'bg-yellow-50 dark:bg-yellow-900/20 opacity-70' : ''}
                        >
                          <TableCell>{formatDate(row.date)}</TableCell>
                          <TableCell>{row.start_time}</TableCell>
                          <TableCell>{toDisplay}</TableCell>
                          <TableCell>{formatDurationMinutes(row.duration_minutes)}</TableCell>
                          <TableCell>{causeBase}</TableCell>
                          <TableCell>
                            {row.exists ? (
                              <span className="text-yellow-600 font-medium text-xs">Existing</span>
                            ) : (
                              <span className="text-green-600 font-medium text-xs">New</span>
                            )}
                          </TableCell>
                        </TableRow>
                      );
                    }

                    if (previewModule === 'line-losses') {
                      return (
                        <TableRow
                          key={idx}
                          className={row.exists ? 'bg-yellow-50 dark:bg-yellow-900/20 opacity-70' : ''}
                        >
                          <TableCell>{formatDate(row.date)}</TableCell>
                          <TableCell className="font-mono-data text-xs">
                            {row.end1_import_final ?? ''}
                          </TableCell>
                          <TableCell className="font-mono-data text-xs">
                            {row.end1_export_final ?? ''}
                          </TableCell>
                          <TableCell className="font-mono-data text-xs">
                            {row.end2_import_final ?? ''}
                          </TableCell>
                          <TableCell className="font-mono-data text-xs">
                            {row.end2_export_final ?? ''}
                          </TableCell>
                          <TableCell>
                            {row.exists ? (
                              <span className="text-yellow-600 font-medium text-xs">Existing</span>
                            ) : (
                              <span className="text-green-600 font-medium text-xs">New</span>
                            )}
                          </TableCell>
                        </TableRow>
                      );
                    }

                    if (previewModule === 'energy') {
                      const readings = Array.isArray(row.readings) ? row.readings : [];
                      return (
                        <TableRow
                          key={idx}
                          className={row.exists ? 'bg-yellow-50 dark:bg-yellow-900/20 opacity-70' : ''}
                        >
                          <TableCell>{formatDate(row.date)}</TableCell>
                          <TableCell className="font-mono-data text-xs">
                            {readings.length}
                          </TableCell>
                          <TableCell>
                            {row.exists ? (
                              <span className="text-yellow-600 font-medium text-xs">Existing</span>
                            ) : (
                              <span className="text-green-600 font-medium text-xs">New</span>
                            )}
                          </TableCell>
                        </TableRow>
                      );
                    }

                    if (previewModule === 'max-min' || previewModule === 'station-load') {
                      const data = row.data || {};
                      const feederType = previewMeta?.feederType;
                      let maxLabel = '';
                      let minLabel = '';
                      let avgLabel = '';
                      if (feederType === 'bus_station') {
                        const max400 = data.max_bus_voltage_400kv?.value;
                        const max220 = data.max_bus_voltage_220kv?.value;
                        const min400 = data.min_bus_voltage_400kv?.value;
                        const min220 = data.min_bus_voltage_220kv?.value;
                        maxLabel = [max400, max220].filter(v => v != null).join(' / ');
                        minLabel = [min400, min220].filter(v => v != null).join(' / ');
                        const maxMw = data.station_load?.max_mw;
                        const mvar = data.station_load?.mvar;
                        avgLabel = [maxMw, mvar].filter(v => v != null).join(' / ');
                      } else {
                        const maxData = data.max || {};
                        const minData = data.min || {};
                        const avgData = data.avg || {};
                        maxLabel = [maxData.amps, maxData.mw].filter(v => v != null).join(' / ');
                        minLabel = [minData.amps, minData.mw].filter(v => v != null).join(' / ');
                        avgLabel = [avgData.amps, avgData.mw].filter(v => v != null).join(' / ');
                      }
                      return (
                        <TableRow
                          key={idx}
                          className={row.exists ? 'bg-yellow-50 dark:bg-yellow-900/20 opacity-70' : ''}
                        >
                          <TableCell>{formatDate(row.date)}</TableCell>
                          <TableCell className="font-mono-data text-xs">{maxLabel}</TableCell>
                          <TableCell className="font-mono-data text-xs">{minLabel}</TableCell>
                          <TableCell className="font-mono-data text-xs">{avgLabel}</TableCell>
                          <TableCell>
                            {row.exists ? (
                              <span className="text-yellow-600 font-medium text-xs">Existing</span>
                            ) : (
                              <span className="text-green-600 font-medium text-xs">New</span>
                            )}
                          </TableCell>
                        </TableRow>
                      );
                    }

                    return (
                      <TableRow key={idx}>
                        <TableCell>{formatDate(row.date)}</TableCell>
                        <TableCell>
                          {row.exists ? (
                            <span className="text-yellow-600 font-medium text-xs">Existing</span>
                          ) : (
                            <span className="text-green-600 font-medium text-xs">New</span>
                          )}
                        </TableCell>
                      </TableRow>
                    );
                  })}
                  {previewRows.length === 0 && (
                    <TableRow>
                      <TableCell
                        colSpan={
                          previewModule === 'interruptions'
                            ? 6
                            : previewModule === 'line-losses'
                              ? 6
                              : previewModule === 'energy'
                                ? 3
                                : previewModule === 'max-min' || previewModule === 'station-load'
                                  ? 5
                                  : 2
                        }
                        className="text-center h-24 text-slate-500"
                      >
                        No rows detected in this file for the selected period.
                      </TableCell>
                    </TableRow>
                  )}
                </TableBody>
              </Table>
            </ScrollArea>
          </div>
          <DialogFooter>
            <Button
              variant="outline"
              onClick={() => {
                setPreviewOpen(false);
              }}
              disabled={loading}
            >
              Close
            </Button>
            <Button
              onClick={async () => {
                setPreviewOpen(false);
                await handleRunImport();
              }}
              disabled={loading || previewRows.length === 0}
            >
              Import
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}

function StatCard({ label, value, tone }) {
  const toneClasses =
    tone === 'success'
      ? 'text-emerald-600 bg-emerald-50 dark:text-emerald-300 dark:bg-emerald-900/20'
      : tone === 'warning'
        ? 'text-amber-600 bg-amber-50 dark:text-amber-300 dark:bg-amber-900/20'
        : 'text-slate-700 bg-slate-50 dark:text-slate-200 dark:bg-slate-900/40';

  return (
    <div className={`rounded-xl px-4 py-3 ${toneClasses}`}>
      <p className="text-xs">{label}</p>
      <p className="text-xl font-semibold mt-1 font-mono-data">{value ?? 0}</p>
    </div>
  );
}
