import { useState, useEffect, useRef } from 'react';
import axios from 'axios';
import { toast } from 'sonner';
import { Button } from '@/components/ui/button';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { FullPageLoader, BlockLoader } from '@/components/ui/loader';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter } from '@/components/ui/dialog';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Download, Calendar, RefreshCcw, Upload, ChevronLeft, ChevronRight, MoreHorizontal, Edit, Trash2, FileText } from 'lucide-react';
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { formatDate } from '@/lib/utils';
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

const FEEDER_ORDER = [
  "Bus Voltages & Station Load",
  "400KV MAHESHWARAM-2",
  "400KV MAHESHWARAM-1",
  "400KV NARSAPUR-1",
  "400KV NARSAPUR-2",
  "400KV KETHIREDDYPALLY-1",
  "400KV KETHIREDDYPALLY-2",
  "400KV NIZAMABAD-1",
  "400KV NIZAMABAD-2",
  "ICT-1 (315MVA)",
  "ICT-2 (315MVA)",
  "ICT-3 (315MVA)",
  "ICT-4 (500MVA)",
  "220KV PARIGI-1",
  "220KV PARIGI-2",
  "220KV THANDUR",
  "220KV GACHIBOWLI-1",
  "220KV GACHIBOWLI-2",
  "220KV KETHIREDDYPALLY",
  "220KV YEDDUMAILARAM-1",
  "220KV YEDDUMAILARAM-2",
  "220KV SADASIVAPET-1",
  "220KV SADASIVAPET-2"
];

const MONTH_LABELS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

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

function splitCauseAndRelay(text) {
  const base = text || '';
  const lower = base.toLowerCase();
  const stripTime = (s) =>
    s.replace(/\bat\s+\d{1,2}[:\.]\d{2}\s*(?:hrs?|hours?)?\b/gi, '');
  const trimEdges = (s) =>
    s.replace(/^[\s\.\-,:]+/, '').replace(/[\s\.\-,:]+$/, '');

  if (lower.includes('a/r success') && lower.includes('with following indications')) {
    const idxFollow = lower.indexOf('with following indications');
    let relay = base.slice(idxFollow + 'with following indications'.length);
    relay = trimEdges(relay);
    let causeSegment = base.slice(0, idxFollow);
    const idxAr = lower.indexOf('a/r success');
    if (idxAr !== -1) {
      causeSegment = base.slice(idxAr, idxFollow);
    }
    let cause = stripTime(causeSegment);
    cause = trimEdges(cause);
    return { cause, relay };
  }

  if (lower.includes('lc issued')) {
    const idxLc = lower.indexOf('lc issued');
    const segment = base.slice(idxLc);
    const lowerSeg = segment.toLowerCase();
    const idxFor = lowerSeg.indexOf(' for ');
    let causeSegment;
    let relaySegment;
    if (idxFor !== -1) {
      causeSegment = segment.slice(0, idxFor);
      relaySegment = segment.slice(idxFor);
    } else {
      causeSegment = segment;
      relaySegment = '';
    }
    let cause = stripTime(causeSegment);
    cause = trimEdges(cause);
    const relay = trimEdges(relaySegment || '');
    return { cause, relay };
  }

  const idxForPlain = lower.indexOf(' for ');
  if (idxForPlain !== -1) {
    const causeSegment = base.slice(0, idxForPlain);
    let cause = stripTime(causeSegment);
    cause = trimEdges(cause);
    const relay = trimEdges(base.slice(idxForPlain));
    return { cause, relay };
  }

  return { cause: base, relay: '' };
}

export default function Interruptions() {
  const [feeders, setFeeders] = useState([]);
  const [selectedFeeder, setSelectedFeeder] = useState(null);
  const [entries, setEntries] = useState([]);
  const [year, setYear] = useState(new Date().getFullYear());
  const [month, setMonth] = useState(new Date().getMonth() + 1);
  const [showDateSelector, setShowDateSelector] = useState(true);
  const [loading, setLoading] = useState(false);
  const [initialized, setInitialized] = useState(false);
  const [importPreviewOpen, setImportPreviewOpen] = useState(false);
  const [importData, setImportData] = useState([]);
  const [selectedImportRows, setSelectedImportRows] = useState({});
  const fileInputRef = useRef(null);
  const [showStickyFeeder, setShowStickyFeeder] = useState(false);
  const feederSelectorRef = useRef(null);
  const [editingEntry, setEditingEntry] = useState(null);
  const [editDate, setEditDate] = useState('');
  const [editStartTime, setEditStartTime] = useState('');
  const [editEndTime, setEditEndTime] = useState('');
  const [editDurationMinutes, setEditDurationMinutes] = useState('');
  const [editCause, setEditCause] = useState('');
  const [editRelay, setEditRelay] = useState('');
  const [editBreakdownDeclared, setEditBreakdownDeclared] = useState('');
  const [editFaultIdentified, setEditFaultIdentified] = useState('');
  const [editFaultLocation, setEditFaultLocation] = useState('');
  const [editRemarks, setEditRemarks] = useState('');
  const [editActionTaken, setEditActionTaken] = useState('');
  const [editEndDate, setEditEndDate] = useState('');
  const [savingEdit, setSavingEdit] = useState(false);
  const [deleteConfirmEntry, setDeleteConfirmEntry] = useState(null);
  const [deletingEntry, setDeletingEntry] = useState(false);
  const [entryModalOpen, setEntryModalOpen] = useState(false);
  const [reportsOpen, setReportsOpen] = useState(false);
  const [reportsLoading, setReportsLoading] = useState(false);
  const [reportsEntriesByFeeder, setReportsEntriesByFeeder] = useState({});
  const [includeBayFeeders, setIncludeBayFeeders] = useState(false);

  useEffect(() => {
    initializeFeeders();
  }, []);

  useEffect(() => {
    const main = document.querySelector('main');
    if (!main) return;
    const handleScroll = () => {
      setShowStickyFeeder(main.scrollTop > 0);
    };
    handleScroll();
    main.addEventListener('scroll', handleScroll);
    return () => main.removeEventListener('scroll', handleScroll);
  }, []);

  const initializeFeeders = async () => {
    try {
      await axios.post(`${API}/max-min/init`);
      const response = await axios.get(`${API}/max-min/feeders`);
      const usable = response.data.filter(
        f =>
          f.type === 'feeder_400kv' ||
          f.type === 'feeder_220kv' ||
          f.type === 'ict_feeder' ||
          f.type === 'reactor_feeder' ||
          f.type === 'bay_feeder'
      );
      setFeeders(usable);
      setInitialized(true);
    } catch (error) {
      toast.error('Failed to load feeders');
    }
  };

  const getSortedFeeders = () => {
    if (!feeders) return [];
    const normalFeeders = feeders.filter(
      f =>
        f.type === 'feeder_400kv' ||
        f.type === 'feeder_220kv' ||
        f.type === 'ict_feeder' ||
        f.type === 'reactor_feeder'
    );
    const bayFeeders = includeBayFeeders
      ? feeders.filter(f => f.type === 'bay_feeder')
      : [];
    const all = [...normalFeeders, ...bayFeeders];
    return all.sort((a, b) => {
      const indexA = FEEDER_ORDER.indexOf(a.name);
      const indexB = FEEDER_ORDER.indexOf(b.name);
      if (indexA !== -1 && indexB !== -1) return indexA - indexB;
      if (indexA !== -1) return -1;
      if (indexB !== -1) return 1;
      return a.name.localeCompare(b.name);
    });
  };

  const handleSubmitDateSelection = () => {
    if (feeders.length > 0) {
      const sorted = getSortedFeeders();
      const feederToUse = selectedFeeder || sorted[0];
      setSelectedFeeder(feederToUse);
      setShowDateSelector(false);
      if (typeof window !== 'undefined') {
        window.dispatchEvent(new Event('collapse-sidebar'));
      }
      fetchEntries(feederToUse.id, year, month);
    } else {
      toast.error("No feeders available");
    }
  };

  const fetchEntries = async (feederId, selectedYear, selectedMonth) => {
    setLoading(true);
    try {
      const response = await axios.get(`${API}/interruptions/entries/${feederId}`, {
        params: { year: selectedYear, month: selectedMonth }
      });
      const sortedEntries = response.data.sort((a, b) => {
        if (a.date === b.date) {
          const ta = (a.data?.start_time || "");
          const tb = (b.data?.start_time || "");
          return ta.localeCompare(tb);
        }
        return a.date.localeCompare(b.date);
      });
      setEntries(sortedEntries);
    } catch (error) {
      toast.error('Failed to load interruption entries');
    } finally {
      setLoading(false);
    }
  };

  const handleFeederChange = (feederId) => {
    const feeder = feeders.find(f => f.id === feederId);
    if (feeder) {
      setSelectedFeeder(feeder);
      fetchEntries(feeder.id, year, month);
    }
  };

  const goToPrevFeeder = () => {
    const sorted = getSortedFeeders();
    const idx = sorted.findIndex(f => f.id === selectedFeeder?.id);
    if (idx > 0) {
      const target = sorted[idx - 1];
      setSelectedFeeder(target);
      fetchEntries(target.id, year, month);
    }
  };

  const goToNextFeeder = () => {
    const sorted = getSortedFeeders();
    const idx = sorted.findIndex(f => f.id === selectedFeeder?.id);
    if (idx !== -1 && idx < sorted.length - 1) {
      const target = sorted[idx + 1];
      setSelectedFeeder(target);
      fetchEntries(target.id, year, month);
    }
  };

  const openReports = async () => {
    if (!feeders || feeders.length === 0) {
      toast.error('No feeders available for reports');
      return;
    }
    setReportsOpen(true);
    setReportsLoading(true);
    try {
      const relevantFeeders = feeders.filter(f =>
        f.type === 'feeder_400kv' ||
        f.type === 'feeder_220kv' ||
        f.type === 'ict_feeder' ||
        f.type === 'reactor_feeder' ||
        f.type === 'bay_feeder'
      );
      const requests = relevantFeeders.map(f =>
        axios
          .get(`${API}/interruptions/entries/${f.id}`, {
            params: { year, month }
          })
          .then(res => [f.id, res.data])
          .catch(() => [f.id, []])
      );
      const results = await Promise.all(requests);
      const map = {};
      results.forEach(([id, data]) => {
        const sortedEntries = (data || []).slice().sort((a, b) => {
          if (a.date === b.date) {
            const ta = (a.data?.start_time || "");
            const tb = (b.data?.start_time || "");
            return ta.localeCompare(tb);
          }
          return a.date.localeCompare(b.date);
        });
        map[id] = sortedEntries;
      });
      setReportsEntriesByFeeder(map);
    } catch (error) {
      toast.error('Failed to load interruption reports');
    } finally {
      setReportsLoading(false);
    }
  };

  const closeReports = () => {
    setReportsOpen(false);
  };

  const buildReportGroups = (types) => {
    const typeList = Array.isArray(types) ? types : [types];
    if (!feeders) return [];
    const is400 = typeList.length === 1 && typeList[0] === 'feeder_400kv';
    const is220 = typeList.length === 1 && typeList[0] === 'feeder_220kv';

    let relevantFeeders = feeders.filter(f => typeList.includes(f.type));

    if (is400) {
      const bayFeeders = feeders.filter(
        f => f.type === 'bay_feeder' && f.name.startsWith('4-')
      );
      relevantFeeders = [...relevantFeeders, ...bayFeeders];
    } else if (is220) {
      const bayFeeders = feeders.filter(
        f => f.type === 'bay_feeder' && f.name.startsWith('2-')
      );
      relevantFeeders = [...relevantFeeders, ...bayFeeders];
    }

    relevantFeeders = relevantFeeders.sort((a, b) => a.name.localeCompare(b.name));

    return relevantFeeders
      .map((feeder) => ({
        name: feeder.name,
        id: feeder.id,
        entries: reportsEntriesByFeeder[feeder.id] || [],
      }))
      .filter((group) => group.entries && group.entries.length > 0);
  };

  const handleImportClick = () => {
    fileInputRef.current?.click();
  };

  const handleFileSelect = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    event.target.value = '';
    const formData = new FormData();
    formData.append('file', file);
    setLoading(true);
    try {
      const response = await axios.post(
        `${API}/interruptions/preview-import-all`,
        formData,
        {
          params: { year, month },
          headers: { 'Content-Type': 'multipart/form-data' },
        }
      );
      setImportData(response.data);
      const initialSelected = {};
      response.data.forEach((row, index) => {
        initialSelected[index] = !row.exists;
      });
      setSelectedImportRows(initialSelected);
      setImportPreviewOpen(true);
    } catch (error) {
      toast.error(error.response?.data?.detail || 'Failed to preview WhatsApp import');
    } finally {
      setLoading(false);
    }
  };

  const handleImportConfirm = async () => {
    setLoading(true);
    try {
      const selectedEntries = importData.filter((_, index) => selectedImportRows[index]);
      if (selectedEntries.length === 0) {
        toast.error('Please select at least one interruption to import');
        setLoading(false);
        return;
      }
      await axios.post(`${API}/interruptions/import-entries-all`, {
        entries: selectedEntries
      });
      toast.success('Interruption data imported successfully');
      setImportPreviewOpen(false);
      if (selectedFeeder) {
        fetchEntries(selectedFeeder.id, year, month);
      }
    } catch (error) {
      toast.error('Failed to import interruption data');
    } finally {
      setLoading(false);
    }
  };

  const handleExport = async () => {
    if (!selectedFeeder) return;
    try {
      const response = await axios.get(
        `${API}/interruptions/export/${selectedFeeder.id}/${year}/${month}`,
        { responseType: 'blob' }
      );
      const blob = new Blob([response.data], { type: response.headers['content-type'] });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Interruptions_${selectedFeeder.name}_${year}_${month.toString().padStart(2, '0')}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (error) {
      toast.error('Failed to export interruption data');
    }
  };

  const handleExportAll = async () => {
    try {
      const response = await axios.get(
        `${API}/interruptions/export-all/${year}/${month}`,
        { responseType: 'blob' }
      );
      const blob = new Blob([response.data], { type: response.headers['content-type'] });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Interruptions_All_${year}_${month.toString().padStart(2, '0')}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (error) {
      toast.error('Failed to export all interruption data');
    }
  };

  const handleRefresh = () => {
    if (!selectedFeeder) return;
    fetchEntries(selectedFeeder.id, year, month);
    toast.success('Data refreshed');
  };

  const recomputeDuration = (fromDate, fromTime, toDate, toTime) => {
    if (!fromDate || !fromTime || !toDate || !toTime) {
      setEditDurationMinutes('');
      return;
    }
    const start = new Date(`${fromDate}T${fromTime}`);
    const end = new Date(`${toDate}T${toTime}`);
    if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) || end <= start) {
      setEditDurationMinutes('');
      return;
    }
    const diffMs = end.getTime() - start.getTime();
    const minutes = Math.round(diffMs / 60000);
    setEditDurationMinutes(String(minutes));
  };

  const openEditModal = (entry) => {
    const data = entry.data || {};
    setEditingEntry(entry);
    setEditDate(entry.date || '');
    const endDateStr = data.end_date || entry.date || '';
    setEditEndDate(endDateStr);
    setEditStartTime(data.start_time || '');
    setEditEndTime(data.end_time || '');
    setEditDurationMinutes(
      data.duration_minutes == null || data.duration_minutes === ''
        ? ''
        : String(data.duration_minutes)
    );
    const desc = data.description || '';
    let cause = data.cause_of_interruption || '';
    let relay = data.relay_indications_lc_work || '';
    if ((!cause || cause === desc) && desc) {
      const parsed = splitCauseAndRelay(desc);
      cause = parsed.cause || desc;
      if (!relay) {
        relay = parsed.relay || '';
      }
    }
    setEditCause(cause);
    setEditRelay(relay);
    setEditBreakdownDeclared(data.breakdown_declared || '');
    setEditFaultIdentified(data.fault_identified_during_patrolling || '');
    setEditFaultLocation(data.fault_location || '');
    setEditRemarks(data.remarks || '');
    setEditActionTaken(data.action_taken || '');
    setEntryModalOpen(true);
  };

  const openNewEntryModal = () => {
    if (!selectedFeeder) {
      toast.error('Please select a feeder first');
      return;
    }
    setEditingEntry(null);
    const now = new Date();
    const y = year;
    const m = String(month).padStart(2, '0');
    const d = String(now.getDate()).padStart(2, '0');
    const defaultDate = `${y}-${m}-${d}`;
    setEditDate(defaultDate);
    setEditEndDate(defaultDate);
    setEditStartTime('');
    setEditEndTime('');
    setEditDurationMinutes('');
    setEditCause('');
    setEditRelay('');
    setEditBreakdownDeclared('');
    setEditFaultIdentified('');
    setEditFaultLocation('');
    setEditRemarks('');
    setEditActionTaken('');
    setEntryModalOpen(true);
  };

  const closeEditModal = () => {
    if (savingEdit) return;
    setEntryModalOpen(false);
    setEditingEntry(null);
    setEditDate('');
    setEditEndDate('');
    setEditStartTime('');
    setEditEndTime('');
    setEditDurationMinutes('');
    setEditCause('');
    setEditRelay('');
    setEditBreakdownDeclared('');
    setEditFaultIdentified('');
    setEditFaultLocation('');
    setEditRemarks('');
    setEditActionTaken('');
  };

  const handleSaveEntry = async () => {
    if (!selectedFeeder) {
      toast.error('Please select a feeder first');
      return;
    }
    if (!editDate) {
      toast.error('Please select a date');
      return;
    }
    try {
      setSavingEdit(true);
      const updatedData = {
        start_time: editStartTime || null,
        end_time: editEndTime || null,
        end_date: editEndDate || editDate,
        duration_minutes: editDurationMinutes === '' ? null : Number(editDurationMinutes),
        cause_of_interruption: editCause || null,
        relay_indications_lc_work: editRelay || null,
        breakdown_declared: editBreakdownDeclared || null,
        fault_identified_during_patrolling: editFaultIdentified || null,
        fault_location: editFaultLocation || null,
        remarks: editRemarks || null,
        action_taken: editActionTaken || null,
      };
      const payload = {
        date: editDate,
        data: updatedData,
      };
      let response;
      if (editingEntry) {
        response = await axios.put(
          `${API}/interruptions/entries/${editingEntry.id}`,
          payload
        );
      } else {
        response = await axios.post(
          `${API}/interruptions/entries/${selectedFeeder.id}`,
          payload
        );
      }
      const savedEntry = response.data;
      setEntries((prev) =>
        [...prev.filter((e) => e.id !== savedEntry.id), savedEntry].sort((a, b) => {
          if (a.date === b.date) {
            const ta = a.data?.start_time || '';
            const tb = b.data?.start_time || '';
            return ta.localeCompare(tb);
          }
          return a.date.localeCompare(b.date);
        })
      );
      toast.success(editingEntry ? 'Interruption entry updated successfully' : 'Interruption entry created successfully');
      closeEditModal();
    } catch (error) {
      const message = error.response?.data?.detail || (editingEntry ? 'Failed to update interruption entry' : 'Failed to create interruption entry');
      toast.error(message);
    } finally {
      setSavingEdit(false);
    }
  };

  const openDeleteConfirm = (entry) => {
    setDeleteConfirmEntry(entry);
  };

  const closeDeleteConfirm = () => {
    if (deletingEntry) return;
    setDeleteConfirmEntry(null);
  };

  const handleConfirmDelete = async () => {
    if (!deleteConfirmEntry) return;
    try {
      setDeletingEntry(true);
      await axios.delete(`${API}/interruptions/entries/${deleteConfirmEntry.id}`);
      setEntries((prev) => prev.filter((e) => e.id !== deleteConfirmEntry.id));
      toast.success('Interruption entry deleted successfully');
      setDeleteConfirmEntry(null);
    } catch (error) {
      toast.error('Failed to delete interruption entry');
    } finally {
      setDeletingEntry(false);
    }
  };

  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];

  if (!initialized) {
    return <FullPageLoader text="Loading Feeders..." />;
  }

  if (showDateSelector) {
    return (
      <div className="flex items-center justify-center" style={{ minHeight: 'calc(100vh - 200px)' }}>
        <Card className="w-full max-w-md shadow-lg">
          <CardHeader>
            <CardTitle className="text-2xl font-heading flex items-center gap-2">
              <Calendar className="w-6 h-6" />
              Select Period
            </CardTitle>
          </CardHeader>
          <CardContent className="space-y-6">
            <div className="space-y-2">
              <label className="text-sm font-medium">Year</label>
              <Select value={year.toString()} onValueChange={(v) => setYear(parseInt(v))}>
                <SelectTrigger>
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  {Array.from({ length: 10 }, (_, i) => new Date().getFullYear() - 5 + i).map(y => (
                    <SelectItem key={y} value={y.toString()}>{y}</SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            <div className="space-y-2">
              <label className="text-sm font-medium">Month</label>
              <Select value={month.toString()} onValueChange={(v) => setMonth(parseInt(v))}>
                <SelectTrigger>
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  {monthNames.map((name, idx) => (
                    <SelectItem key={idx + 1} value={(idx + 1).toString()}>{name}</SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            <Button onClick={handleSubmitDateSelection} className="w-full" size="lg">
              Load Data
            </Button>
          </CardContent>
        </Card>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-6 bg-white dark:bg-slate-950 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-800">
        <div className="flex w-full justify-between items-start lg:w-auto lg:block">
          <div>
            <div className="flex items-center gap-3">
              <h1 className="text-3xl md:text-4xl font-heading font-extrabold text-transparent bg-clip-text bg-gradient-to-r from-indigo-600 via-purple-600 to-pink-600">
                Interruptions
              </h1>
              <span className="text-sm md:text-base text-slate-500 font-medium mt-2 flex items-center gap-2">
                <span className="inline-block w-2.5 h-2.5 rounded-full bg-emerald-500 ring-4 ring-emerald-50"></span>
                {monthNames[month - 1]} {year}
              </span>
            </div>
          </div>

          <div className="md:hidden">
            <DropdownMenu>
              <DropdownMenuTrigger asChild>
                <Button variant="outline" size="icon">
                  <MoreHorizontal className="h-4 w-4" />
                </Button>
              </DropdownMenuTrigger>
              <DropdownMenuContent align="end">
                <DropdownMenuItem onClick={handleRefresh} disabled={!selectedFeeder}>
                  <RefreshCcw className="w-4 h-4 mr-2" />
                  Refresh
                </DropdownMenuItem>
                <DropdownMenuItem onClick={() => setShowDateSelector(true)}>
                  <Calendar className="w-4 h-4 mr-2" />
                  Period
                </DropdownMenuItem>
                <DropdownMenuItem onClick={handleImportClick} disabled={!selectedFeeder}>
                  <Upload className="w-4 h-4 mr-2" />
                  Import WhatsApp
                </DropdownMenuItem>
                <DropdownMenuItem onClick={handleExport} disabled={!selectedFeeder || entries.length === 0}>
                  <Download className="w-4 h-4 mr-2" />
                  Export
                </DropdownMenuItem>
                <DropdownMenuItem onClick={handleExportAll} disabled={!feeders.length}>
                  <Download className="w-4 h-4 mr-2" />
                  Export All
                </DropdownMenuItem>
              </DropdownMenuContent>
            </DropdownMenu>
          </div>
        </div>

        <div className="flex items-center gap-3 w-full lg:w-auto justify-end">
          <div className="hidden md:flex flex-wrap items-center gap-3 w-full lg:w-auto">
            <Button
              variant="outline"
              onClick={handleRefresh}
              disabled={!selectedFeeder}
              className="flex-1 lg:flex-none border-slate-200 hover:bg-slate-50 text-slate-700 font-medium"
            >
              <RefreshCcw className="w-4 h-4 mr-2 text-indigo-500" />
              Refresh
            </Button>
            <Button
              variant="outline"
              onClick={() => setShowDateSelector(true)}
              className="flex-1 lg:flex-none border-slate-200 hover:bg-slate-50 text-slate-700 font-medium"
            >
              <Calendar className="w-4 h-4 mr-2 text-indigo-500" />
              Period
            </Button>
            <Button
              onClick={handleImportClick}
              variant="outline"
              disabled={!selectedFeeder}
              className="flex-1 lg:flex-none border-slate-200 hover:bg-slate-50 text-slate-700 font-medium"
            >
              <Upload className="w-4 h-4 mr-2 text-indigo-500" />
              Import WhatsApp
            </Button>
            <Button
              onClick={openNewEntryModal}
              variant="default"
              disabled={!selectedFeeder}
              className="flex-1 lg:flex-none bg-indigo-600 text-white hover:bg-indigo-700 font-medium"
            >
              + Entry
            </Button>
            <Button
              variant="secondary"
              onClick={handleExport}
              disabled={!selectedFeeder || entries.length === 0}
              className="flex-1 lg:flex-none bg-emerald-50 text-emerald-700 hover:bg-emerald-100 border border-emerald-200 font-medium"
            >
              <Download className="w-4 h-4 mr-2" />
              Export
            </Button>
            <Button
              variant="secondary"
              onClick={handleExportAll}
              disabled={!feeders.length}
              className="flex-1 lg:flex-none bg-blue-50 text-blue-700 hover:bg-blue-100 border border-blue-200 font-medium"
            >
              <Download className="w-4 h-4 mr-2" />
              Export All
            </Button>
          </div>

          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileSelect}
            className="hidden"
            accept=".txt"
          />
        </div>
      </div>

      <div ref={feederSelectorRef} className="flex items-end justify-between mb-2">
        <Card className="w-full max-w-xl border-0 shadow-md ring-1 ring-slate-100 overflow-hidden">
          <div className="h-1 w-full bg-gradient-to-r from-indigo-500 to-purple-500"></div>
          <CardContent className="pt-6 pb-6">
            <div className="flex items-center justify-between mb-3">
              <label className="text-xs font-bold uppercase tracking-wider text-slate-500 flex items-center gap-2">
                <div className="w-1.5 h-1.5 rounded-full bg-indigo-500"></div>
                Select Feeder
              </label>
              <Button
                type="button"
                variant="ghost"
                size="sm"
                className="text-[11px] h-7 px-2 text-indigo-600"
                onClick={() => setIncludeBayFeeders((v) => !v)}
              >
                {includeBayFeeders ? 'Hide Bay Feeders' : 'Show Bay Feeders'}
              </Button>
            </div>
            <div className="flex items-center justify-between gap-4">
              <div className="flex items-center gap-2 flex-1">
                <Button
                  variant="outline"
                  size="icon"
                  onClick={goToPrevFeeder}
                  disabled={!selectedFeeder}
                  title="Previous Feeder"
                  className="h-12 w-12 shrink-0"
                >
                  <ChevronLeft className="w-5 h-5" />
                </Button>
                <Select
                  value={selectedFeeder?.id || ''}
                  onValueChange={handleFeederChange}
                >
                  <SelectTrigger className="w-full h-12 text-lg border-slate-200 focus:ring-2 focus:ring-indigo-500/20 bg-slate-50/50 transition-all hover:bg-white">
                    <SelectValue placeholder="Select a feeder" />
                  </SelectTrigger>
                  <SelectContent className="max-h-[300px]">
                    {getSortedFeeders().map(feeder => (
                      <SelectItem key={feeder.id} value={feeder.id} className="focus:bg-indigo-50 focus:text-indigo-700 py-3 cursor-pointer">
                        {feeder.name}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
                <Button
                  variant="outline"
                  size="icon"
                  onClick={goToNextFeeder}
                  disabled={!selectedFeeder}
                  title="Next Feeder"
                  className="h-12 w-12 shrink-0"
                >
                  <ChevronRight className="w-5 h-5" />
                </Button>
              </div>
            </div>
          </CardContent>
        </Card>
        <Button
          type="button"
          variant="outline"
          className="ml-4 h-12 shrink-0"
          onClick={openReports}
        >
          <FileText className="w-4 h-4 mr-2" />
          Reports
        </Button>
      </div>

      {selectedFeeder && showStickyFeeder && (
        <div className="sticky top-0 z-20 bg-white/95 border-y border-slate-200 py-2 mb-4">
          <div className="text-sm font-semibold text-slate-700">
            Viewing Feeder: <span className="font-bold">{selectedFeeder.name}</span>
          </div>
        </div>
      )}

      {selectedFeeder && (
        <Card className="shadow-lg border-0 ring-1 ring-slate-100 overflow-hidden bg-white relative">
          {loading && <BlockLoader text="Loading entries..." className="z-20" />}
          <CardHeader className="bg-slate-50/50 border-b border-slate-100 py-5">
            <div className="flex items-center justify-between">
              <CardTitle className="text-lg font-bold text-slate-700 flex items-center gap-2">
                <div className="p-2 bg-indigo-100 text-indigo-600 rounded-lg">
                  <Calendar className="w-5 h-5" />
                </div>
                Interruption Events
              </CardTitle>
              <span className="text-sm font-semibold text-slate-500 bg-white px-4 py-1.5 rounded-full border shadow-sm">
                {entries.length} Events
              </span>
            </div>
          </CardHeader>
          <CardContent className="p-0 overflow-x-auto">
            <Table>
              <TableHeader>
                <TableRow className="bg-slate-100 border-b border-slate-200">
                  <TableHead className="w-[60px] font-bold text-slate-700 text-center">Sl. No</TableHead>
                  <TableHead className="w-[120px] font-bold text-slate-700 text-center">Date</TableHead>
                  <TableHead className="font-bold text-slate-700 text-center">Time From</TableHead>
                  <TableHead className="font-bold text-slate-700 text-center">Time To</TableHead>
                  <TableHead className="font-bold text-slate-700 text-center">Duration</TableHead>
                  <TableHead className="font-bold text-slate-700">Cause of Interruption</TableHead>
                  <TableHead className="font-bold text-slate-700">Relay Indications / LC Work carried out</TableHead>
                  <TableHead className="font-bold text-slate-700 text-center">Breakdown Declared (Yes/No)</TableHead>
                  <TableHead className="font-bold text-slate-700 text-center">Fault Identified During Patrolling</TableHead>
                  <TableHead className="font-bold text-slate-700 text-center">Fault Location</TableHead>
                  <TableHead className="font-bold text-slate-700">Remarks</TableHead>
                  <TableHead className="font-bold text-slate-700">Action Taken</TableHead>
                  <TableHead className="w-[80px] font-bold text-slate-700 text-center">Actions</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {entries.length === 0 ? (
                  <TableRow>
                    <TableCell colSpan={13} className="text-center py-8 text-muted-foreground">
                      No interruptions found for this month. Use Import WhatsApp to add data.
                    </TableCell>
                  </TableRow>
                ) : (
                  entries.map((entry, index) => {
                    const data = entry.data || {};
                    const desc = data.description || '';
                    let cause = data.cause_of_interruption || '';
                    let relay = data.relay_indications_lc_work || '';
                    if ((!cause || cause === desc) && desc) {
                      const parsed = splitCauseAndRelay(desc);
                      cause = parsed.cause || desc;
                      if (!relay) {
                        relay = parsed.relay || '';
                      }
                    }
                    const endDateStr = data.end_date || entry.date;
                    const endTimeStr = data.end_time || '';
                    const toDisplay =
                      !endTimeStr
                        ? ''
                        : endDateStr === entry.date
                        ? endTimeStr
                        : `${formatDate(endDateStr)} ${endTimeStr}`;
                    return (
                      <TableRow key={entry.id}>
                        <TableCell className="text-center">
                          {index + 1}
                        </TableCell>
                        <TableCell className="font-medium whitespace-nowrap text-center">
                          {formatDate(entry.date)}
                        </TableCell>
                        <TableCell className="text-center">
                          {data.start_time || ''}
                        </TableCell>
                        <TableCell className="text-center">
                          {toDisplay}
                        </TableCell>
                        <TableCell className="text-center">
                          {formatDurationMinutes(data.duration_minutes)}
                        </TableCell>
                        <TableCell className="max-w-xs">
                          <span className="block whitespace-pre-wrap">
                            {cause}
                          </span>
                        </TableCell>
                        <TableCell className="max-w-xl">
                          <span className="block whitespace-pre-wrap">
                            {relay}
                          </span>
                        </TableCell>
                        <TableCell className="text-center">
                          {data.breakdown_declared || ''}
                        </TableCell>
                        <TableCell className="text-center">
                          {data.fault_identified_during_patrolling || ''}
                        </TableCell>
                        <TableCell className="text-center">
                          {data.fault_location || ''}
                        </TableCell>
                        <TableCell>
                          <span className="block whitespace-pre-wrap">
                            {data.remarks || ''}
                          </span>
                        </TableCell>
                        <TableCell>
                          <span className="block whitespace-pre-wrap">
                            {data.action_taken || ''}
                          </span>
                        </TableCell>
                        <TableCell className="text-center">
                          <div className="flex items-center justify-center gap-1">
                            <Button
                              size="icon"
                              variant="ghost"
                              className="h-8 w-8 text-slate-600 hover:text-slate-800 hover:bg-slate-100"
                              onClick={() => openEditModal(entry)}
                            >
                              <Edit className="w-4 h-4" />
                            </Button>
                            <Button
                              size="icon"
                              variant="ghost"
                              className="h-8 w-8 text-red-600 hover:text-red-700 hover:bg-red-50"
                              onClick={() => openDeleteConfirm(entry)}
                            >
                              <Trash2 className="w-4 h-4" />
                            </Button>
                          </div>
                        </TableCell>
                      </TableRow>
                    );
                  })
                )}
              </TableBody>
            </Table>
          </CardContent>
        </Card>
      )}

      <Dialog open={entryModalOpen} onOpenChange={(open) => { if (!open) closeEditModal(); }}>
        <DialogContent className="max-w-3xl max-h-[90vh] overflow-y-auto">
          <DialogHeader>
            <DialogTitle>Edit Interruption Entry</DialogTitle>
          </DialogHeader>
          <div className="space-y-4">
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
              <div className="space-y-1">
                <Label>From Date</Label>
                <Input
                  type="date"
                  value={editDate}
                  onChange={(e) => {
                    const value = e.target.value;
                    setEditDate(value);
                    recomputeDuration(value, editStartTime, editEndDate, editEndTime);
                  }}
                />
              </div>
              <div className="space-y-1">
                <Label>From Time</Label>
                <Input
                  value={editStartTime}
                  onChange={(e) => {
                    const value = e.target.value;
                    setEditStartTime(value);
                    recomputeDuration(editDate, value, editEndDate, editEndTime);
                  }}
                  placeholder="HH:MM"
                />
              </div>
              <div className="space-y-1">
                <Label>To Date</Label>
                <Input
                  type="date"
                  value={editEndDate}
                  onChange={(e) => {
                    const value = e.target.value;
                    setEditEndDate(value);
                    recomputeDuration(editDate, editStartTime, value, editEndTime);
                  }}
                />
              </div>
              <div className="space-y-1">
                <Label>To Time</Label>
                <Input
                  value={editEndTime}
                  onChange={(e) => {
                    const value = e.target.value;
                    setEditEndTime(value);
                    recomputeDuration(editDate, editStartTime, editEndDate, value);
                  }}
                  placeholder="HH:MM"
                />
              </div>
            </div>
            <div className="space-y-1">
              <Label>Duration</Label>
              <Input
                value={formatDurationMinutes(editDurationMinutes)}
                readOnly
              />
            </div>
            <div className="space-y-1">
              <Label>Cause of Interruption</Label>
              <textarea
                className="w-full border rounded-md px-3 py-2 text-sm"
                rows={3}
                value={editCause}
                onChange={(e) => setEditCause(e.target.value)}
              />
            </div>
            <div className="space-y-1">
              <Label>Relay Indications / LC Work carried out</Label>
              <textarea
                className="w-full border rounded-md px-3 py-2 text-sm"
                rows={3}
                value={editRelay}
                onChange={(e) => setEditRelay(e.target.value)}
              />
            </div>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div className="space-y-1">
                <Label>Breakdown Declared (Yes/No)</Label>
                <Input
                  value={editBreakdownDeclared}
                  onChange={(e) => setEditBreakdownDeclared(e.target.value)}
                />
              </div>
              <div className="space-y-1">
                <Label>Fault Identified During Patrolling</Label>
                <Input
                  value={editFaultIdentified}
                  onChange={(e) => setEditFaultIdentified(e.target.value)}
                />
              </div>
              <div className="space-y-1">
                <Label>Fault Location</Label>
                <Input
                  value={editFaultLocation}
                  onChange={(e) => setEditFaultLocation(e.target.value)}
                />
              </div>
            </div>
            <div className="space-y-1">
              <Label>Remarks</Label>
              <textarea
                className="w-full border rounded-md px-3 py-2 text-sm"
                rows={2}
                value={editRemarks}
                onChange={(e) => setEditRemarks(e.target.value)}
              />
            </div>
            <div className="space-y-1">
              <Label>Action Taken</Label>
              <textarea
                className="w-full border rounded-md px-3 py-2 text-sm"
                rows={2}
                value={editActionTaken}
                onChange={(e) => setEditActionTaken(e.target.value)}
              />
            </div>
          </div>
          <DialogFooter>
            <Button variant="outline" onClick={closeEditModal} disabled={savingEdit}>
              Cancel
            </Button>
            <Button onClick={handleSaveEntry} disabled={savingEdit}>
              {savingEdit ? 'Saving...' : 'Save Changes'}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <Dialog open={reportsOpen} onOpenChange={(open) => { if (!open) closeReports(); }}>
        <DialogContent className="max-w-5xl max-h-[90vh] overflow-y-auto">
          <DialogHeader>
            <DialogTitle>
              Interruption Reports – {year}-{String(month).padStart(2, '0')}
            </DialogTitle>
          </DialogHeader>
          {reportsLoading ? (
            <BlockLoader text="Loading reports..." />
          ) : (
            <Tabs defaultValue="400kv" className="w-full mt-2">
              <TabsList className="grid w-full grid-cols-3 mb-4">
                <TabsTrigger value="400kv">400KV Feeders</TabsTrigger>
                <TabsTrigger value="220kv">220KV Feeders</TabsTrigger>
                <TabsTrigger value="ict_reactor">ICTs & 125MVAR Reactor</TabsTrigger>
              </TabsList>
              <TabsContent value="400kv">
                <div className="border rounded-md overflow-x-auto">
                  <Table>
                    <TableHeader>
                      <TableRow>
                        <TableHead colSpan={11} className="text-center border bg-muted h-auto py-2">
                          {`Interruptions of 400KV SS Shankarpally 400KV Feeders for the Month of ${MONTH_LABELS[month - 1] || month}-${year}`}
                        </TableHead>
                      </TableRow>
                      <TableRow>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 w-[60px]">Sl. No</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[120px]">Date</TableHead>
                        <TableHead colSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[140px]">Time</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[80px]">Duration</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[260px]">Cause of Interruption</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[260px]">Relay Indications / LC Work carried out</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[130px]">Break down declared or Not</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[150px]">Fault identified in patrolling</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[120px]">Fault Location</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[200px]">Remarks and action taken</TableHead>
                      </TableRow>
                      <TableRow>
                        <TableHead className="text-center border bg-muted h-auto py-2 min-w-[70px]">From</TableHead>
                        <TableHead className="text-center border bg-muted h-auto py-2 min-w-[70px]">To</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {(() => {
                        const groups = buildReportGroups('feeder_400kv');
                        if (!groups.length) {
                          return (
                            <TableRow>
                              <TableCell colSpan={11} className="text-center py-4">
                                No interruptions found for 400KV feeders in this period.
                              </TableCell>
                            </TableRow>
                          );
                        }
                        return groups.map((group) => (
                          <>
                            <TableRow key={`${group.name}-header`}>
                              <TableCell colSpan={11} className="border font-semibold text-center">
                                {group.name}
                              </TableCell>
                            </TableRow>
                            {group.entries.map((entry, index) => {
                                const data = entry.data || {};
                                const desc = data.description || '';
                                let cause = data.cause_of_interruption || '';
                                let relay = data.relay_indications_lc_work || '';
                                if ((!cause || cause === desc) && desc) {
                                  const parsed = splitCauseAndRelay(desc);
                                  cause = parsed.cause || desc;
                                  if (!relay) {
                                    relay = parsed.relay || '';
                                  }
                                }
                                const endDateStr = data.end_date || entry.date;
                                const endTimeStr = data.end_time || '';
                                const toDisplay =
                                  !endTimeStr
                                    ? ''
                                    : endDateStr === entry.date
                                    ? endTimeStr
                                    : `${formatDate(endDateStr)} ${endTimeStr}`;
                                const remarksCombined = [data.remarks, data.action_taken].filter(Boolean).join('\n');
                                return (
                                  <TableRow key={entry.id || `${group.name}-${index}`}>
                                    <TableCell className="text-center border p-2">{index + 1}</TableCell>
                                    <TableCell className="text-center border p-2 whitespace-nowrap">
                                      {formatDate(entry.date)}
                                    </TableCell>
                                    <TableCell className="text-center border p-2">
                                      {data.start_time || ''}
                                    </TableCell>
                                    <TableCell className="text-center border p-2">
                                      {toDisplay}
                                    </TableCell>
                                    <TableCell className="text-center border p-2">
                                      {formatDurationMinutes(data.duration_minutes)}
                                    </TableCell>
                                    <TableCell className="border p-2 whitespace-pre-wrap">
                                      {cause}
                                    </TableCell>
                                    <TableCell className="border p-2 whitespace-pre-wrap">
                                      {relay}
                                    </TableCell>
                                    <TableCell className="text-center border p-2">
                                      {data.breakdown_declared || ''}
                                    </TableCell>
                                    <TableCell className="text-center border p-2">
                                      {data.fault_identified_during_patrolling || ''}
                                    </TableCell>
                                    <TableCell className="text-center border p-2">
                                      {data.fault_location || ''}
                                    </TableCell>
                                    <TableCell className="border p-2 whitespace-pre-wrap">
                                      {remarksCombined}
                                    </TableCell>
                                  </TableRow>
                                );
                              })}
                          </>
                        ));
                      })()}
                    </TableBody>
                  </Table>
                </div>
              </TabsContent>
              <TabsContent value="220kv">
                <div className="border rounded-md overflow-x-auto">
                  <Table>
                    <TableHeader>
                      <TableRow>
                        <TableHead colSpan={11} className="text-center border bg-muted h-auto py-2">
                          {`Interruptions of 400KV SS Shankarpally 220KV Feeders for the Month of ${MONTH_LABELS[month - 1] || month}-${year}`}
                        </TableHead>
                      </TableRow>
                      <TableRow>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 w-[60px]">Sl. No</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[120px]">Date</TableHead>
                        <TableHead colSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[140px]">Time</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[80px]">Duration</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[260px]">Cause of Interruption</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[260px]">Relay Indications / LC Work carried out</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[130px]">Break down declared or Not</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[150px]">Fault identified in patrolling</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[120px]">Fault Location</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[200px]">Remarks and action taken</TableHead>
                      </TableRow>
                      <TableRow>
                        <TableHead className="text-center border bg-muted h-auto py-2 min-w-[70px]">From</TableHead>
                        <TableHead className="text-center border bg-muted h-auto py-2 min-w-[70px]">To</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {(() => {
                        const groups = buildReportGroups('feeder_220kv');
                        if (!groups.length) {
                          return (
                            <TableRow>
                              <TableCell colSpan={11} className="text-center py-4">
                                No interruptions found for 220KV feeders in this period.
                              </TableCell>
                            </TableRow>
                          );
                        }
                        return groups.map((group) => (
                          <>
                            <TableRow key={`${group.name}-header`}>
                              <TableCell colSpan={11} className="border font-semibold text-center">
                                {group.name}
                              </TableCell>
                            </TableRow>
                            {group.entries.map((entry, index) => {
                                const data = entry.data || {};
                                const desc = data.description || '';
                                let cause = data.cause_of_interruption || '';
                                let relay = data.relay_indications_lc_work || '';
                                if ((!cause || cause === desc) && desc) {
                                  const parsed = splitCauseAndRelay(desc);
                                  cause = parsed.cause || desc;
                                  if (!relay) {
                                    relay = parsed.relay || '';
                                  }
                                }
                                const endDateStr = data.end_date || entry.date;
                                const endTimeStr = data.end_time || '';
                                const toDisplay =
                                  !endTimeStr
                                    ? ''
                                    : endDateStr === entry.date
                                    ? endTimeStr
                                    : `${formatDate(endDateStr)} ${endTimeStr}`;
                                const remarksCombined = [data.remarks, data.action_taken].filter(Boolean).join('\n');
                                return (
                                  <TableRow key={entry.id || `${group.name}-${index}`}>
                                    <TableCell className="text-center border p-2">{index + 1}</TableCell>
                                    <TableCell className="text-center border p-2 whitespace-nowrap">
                                      {formatDate(entry.date)}
                                    </TableCell>
                                    <TableCell className="text-center border p-2">
                                      {data.start_time || ''}
                                    </TableCell>
                                    <TableCell className="text-center border p-2">
                                      {toDisplay}
                                    </TableCell>
                                    <TableCell className="text-center border p-2">
                                      {formatDurationMinutes(data.duration_minutes)}
                                    </TableCell>
                                    <TableCell className="border p-2 whitespace-pre-wrap">
                                      {cause}
                                    </TableCell>
                                    <TableCell className="border p-2 whitespace-pre-wrap">
                                      {relay}
                                    </TableCell>
                                    <TableCell className="text-center border p-2">
                                      {data.breakdown_declared || ''}
                                    </TableCell>
                                    <TableCell className="text-center border p-2">
                                      {data.fault_identified_during_patrolling || ''}
                                    </TableCell>
                                    <TableCell className="text-center border p-2">
                                      {data.fault_location || ''}
                                    </TableCell>
                                    <TableCell className="border p-2 whitespace-pre-wrap">
                                      {remarksCombined}
                                    </TableCell>
                                  </TableRow>
                                );
                              })}
                          </>
                        ));
                      })()}
                    </TableBody>
                  </Table>
                </div>
              </TabsContent>
              <TabsContent value="ict_reactor">
                <div className="border rounded-md overflow-x-auto">
                  <Table>
                    <TableHeader>
                      <TableRow>
                        <TableHead colSpan={11} className="text-center border bg-muted h-auto py-2">
                          {`Interruptions of 400KV SS Shankarpally ICTs / Bays / Reactor for the Month of ${MONTH_LABELS[month - 1] || month}-${year}`}
                        </TableHead>
                      </TableRow>
                      <TableRow>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 w-[60px]">Sl. No</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[120px]">Date</TableHead>
                        <TableHead colSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[140px]">Time</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[80px]">Duration</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[260px]">Cause of Interruption</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[260px]">Relay Indications / LC Work carried out</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[130px]">Break down declared or Not</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[150px]">Fault identified in patrolling</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[120px]">Fault Location</TableHead>
                        <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[200px]">Remarks and action taken</TableHead>
                      </TableRow>
                      <TableRow>
                        <TableHead className="text-center border bg-muted h-auto py-2 min-w-[70px]">From</TableHead>
                        <TableHead className="text-center border bg-muted h-auto py-2 min-w-[70px]">To</TableHead>
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {(() => {
                        const groups = buildReportGroups(['ict_feeder', 'reactor_feeder']);
                        if (!groups.length) {
                          return (
                            <TableRow>
                              <TableCell colSpan={11} className="text-center py-4">
                                No interruptions found for ICTs / Reactor in this period.
                              </TableCell>
                            </TableRow>
                          );
                        }
                        return groups.map((group, groupIndex) => (
                          <>
                            <TableRow key={`${group.name}-header`}>
                              <TableCell colSpan={11} className="border font-semibold text-center">
                                {group.name}
                              </TableCell>
                            </TableRow>
                            {group.entries.map((entry, index) => {
                                const data = entry.data || {};
                                const desc = data.description || '';
                                let cause = data.cause_of_interruption || '';
                                let relay = data.relay_indications_lc_work || '';
                                if ((!cause || cause === desc) && desc) {
                                  const parsed = splitCauseAndRelay(desc);
                                  cause = parsed.cause || desc;
                                  if (!relay) {
                                    relay = parsed.relay || '';
                                  }
                                }
                                const endDateStr = data.end_date || entry.date;
                                const endTimeStr = data.end_time || '';
                                const toDisplay =
                                  !endTimeStr
                                    ? ''
                                    : endDateStr === entry.date
                                    ? endTimeStr
                                    : `${formatDate(endDateStr)} ${endTimeStr}`;
                                const remarksCombined = [data.remarks, data.action_taken].filter(Boolean).join('\n');
                                return (
                                  <TableRow key={entry.id || `${group.name}-${index}`}>
                                    <TableCell className="text-center border p-2">{index + 1}</TableCell>
                                    <TableCell className="text-center border p-2 whitespace-nowrap">
                                      {formatDate(entry.date)}
                                    </TableCell>
                                    <TableCell className="text-center border p-2">
                                      {data.start_time || ''}
                                    </TableCell>
                                    <TableCell className="text-center border p-2">
                                      {toDisplay}
                                    </TableCell>
                                    <TableCell className="text-center border p-2">
                                      {formatDurationMinutes(data.duration_minutes)}
                                    </TableCell>
                                    <TableCell className="border p-2 whitespace-pre-wrap">
                                      {cause}
                                    </TableCell>
                                    <TableCell className="border p-2 whitespace-pre-wrap">
                                      {relay}
                                    </TableCell>
                                    <TableCell className="text-center border p-2">
                                      {data.breakdown_declared || ''}
                                    </TableCell>
                                    <TableCell className="text-center border p-2">
                                      {data.fault_identified_during_patrolling || ''}
                                    </TableCell>
                                    <TableCell className="text-center border p-2">
                                      {data.fault_location || ''}
                                    </TableCell>
                                    <TableCell className="border p-2 whitespace-pre-wrap">
                                      {remarksCombined}
                                    </TableCell>
                                  </TableRow>
                                );
                              })}
                          </>
                        ));
                      })()}
                    </TableBody>
                  </Table>
                </div>
              </TabsContent>
            </Tabs>
          )}
        </DialogContent>
      </Dialog>

      <Dialog open={!!deleteConfirmEntry} onOpenChange={closeDeleteConfirm}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Confirm Delete</DialogTitle>
          </DialogHeader>
          <p className="text-sm text-slate-600">
            Are you sure you want to delete this interruption entry? This action cannot be undone.
          </p>
          <DialogFooter>
            <Button variant="outline" onClick={closeDeleteConfirm} disabled={deletingEntry}>
              Cancel
            </Button>
            <Button variant="destructive" onClick={handleConfirmDelete} disabled={deletingEntry}>
              {deletingEntry ? 'Deleting...' : 'Delete'}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {importPreviewOpen && (
        <Card className="mt-4 shadow-lg border-0 ring-1 ring-slate-100 overflow-hidden bg-white">
          <CardHeader className="bg-slate-50/50 border-b border-slate-100 py-4 flex flex-row items-center justify-between">
            <CardTitle className="text-lg font-bold text-slate-700">
              Import Preview
            </CardTitle>
            <div className="flex gap-2">
              <Button
                variant="outline"
                onClick={() => setImportPreviewOpen(false)}
              >
                Cancel
              </Button>
              <Button
                onClick={handleImportConfirm}
              >
                Confirm Import
              </Button>
            </div>
          </CardHeader>
          <CardContent className="p-0 overflow-x-auto max-h-[400px]">
            <Table>
              <TableHeader>
                <TableRow className="bg-slate-100 border-b border-slate-200">
                  <TableHead className="w-[50px] text-center">
                    <input
                      type="checkbox"
                      className="h-4 w-4"
                      checked={
                        importData.length > 0 &&
                        importData.every((_, index) => selectedImportRows[index])
                      }
                      onChange={(e) => {
                        const checked = e.target.checked;
                        const next = {};
                        importData.forEach((_, index) => {
                          next[index] = checked;
                        });
                        setSelectedImportRows(next);
                      }}
                    />
                  </TableHead>
                  <TableHead className="text-center">Date</TableHead>
                  <TableHead className="text-center">Time From</TableHead>
                  <TableHead className="text-center">Time To</TableHead>
                  <TableHead className="text-center">Duration</TableHead>
                  <TableHead>Cause of Interruption</TableHead>
                  <TableHead>Relay Indications / LC Work carried out</TableHead>
                  <TableHead className="text-center">Existing</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {importData.length === 0 ? (
                  <TableRow>
                    <TableCell colSpan={7} className="text-center py-6 text-muted-foreground">
                      No interruption events detected in the uploaded chat for this feeder and period.
                    </TableCell>
                  </TableRow>
                ) : (
                  importData.map((row, idx) => {
                    const endDateStr = row.end_date || row.date;
                    const endTimeStr = row.end_time || '';
                    const toDisplay =
                      !endTimeStr
                        ? ''
                        : endDateStr === row.date
                        ? endTimeStr
                        : `${formatDate(endDateStr)} ${endTimeStr}`;
                    return (
                      <TableRow key={idx}>
                        <TableCell className="text-center">
                          <input
                            type="checkbox"
                            className="h-4 w-4"
                            checked={!!selectedImportRows[idx]}
                            onChange={(e) => {
                              const checked = e.target.checked;
                              setSelectedImportRows(prev => ({
                                ...prev,
                                [idx]: checked,
                              }));
                            }}
                          />
                        </TableCell>
                        <TableCell className="text-center">{formatDate(row.date)}</TableCell>
                        <TableCell className="text-center">{row.start_time}</TableCell>
                        <TableCell className="text-center">{toDisplay}</TableCell>
                        <TableCell className="text-center">{formatDurationMinutes(row.duration_minutes)}</TableCell>
                        <TableCell>
                          <span className="block whitespace-pre-wrap">
                            {(() => {
                              const desc = row.description || '';
                              let cause = row.cause_of_interruption || '';
                              let relay = row.relay_indications_lc_work || '';
                              if ((!cause || cause === desc) && desc) {
                                const parsed = splitCauseAndRelay(desc);
                                cause = parsed.cause || desc;
                                if (!relay) {
                                  relay = parsed.relay || '';
                                }
                              }
                              return cause;
                            })()}
                          </span>
                        </TableCell>
                        <TableCell>
                          <span className="block whitespace-pre-wrap">
                            {(() => {
                              const desc = row.description || '';
                              let cause = row.cause_of_interruption || '';
                              let relay = row.relay_indications_lc_work || '';
                              if ((!cause || cause === desc) && desc) {
                                const parsed = splitCauseAndRelay(desc);
                                cause = parsed.cause || desc;
                                if (!relay) {
                                  relay = parsed.relay || '';
                                }
                              }
                              return relay;
                            })()}
                          </span>
                        </TableCell>
                        <TableCell className="text-center">
                          {row.exists ? (
                            <span className="text-xs font-semibold text-amber-600 bg-amber-100 px-2 py-1 rounded-full">
                              Existing
                            </span>
                          ) : (
                            <span className="text-xs font-semibold text-emerald-600 bg-emerald-100 px-2 py-1 rounded-full">
                              New
                            </span>
                          )}
                        </TableCell>
                      </TableRow>
                    );
                  })
                )}
              </TableBody>
            </Table>
          </CardContent>
        </Card>
      )}
    </div>
  );
}
