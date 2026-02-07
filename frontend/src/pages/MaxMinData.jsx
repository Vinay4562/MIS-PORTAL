 import React, { useState, useEffect, Fragment, useRef } from 'react';
import axios from 'axios';
import { toast } from 'sonner';
import { Button } from '@/components/ui/button';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
 import { Calendar, Download, Plus, Edit, Trash2, RefreshCcw, Upload, ChevronLeft, ChevronRight } from 'lucide-react';
import { formatDate } from '@/lib/utils';
import MaxMinEntryModal from '@/components/MaxMinEntryModal';
import MaxMinAnalytics from '@/components/MaxMinAnalytics';
import MaxMinImportPreviewModal from '@/components/MaxMinImportPreviewModal';
import { FullPageLoader, BlockLoader } from '@/components/ui/loader';

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

const formatTime = (t) => {
  if (!t) return '';
  const str = String(t).trim();
  if (str.includes(':')) {
    const parts = str.split(':');
    if (parts.length >= 2) {
      return `${parts[0].padStart(2, '0')}:${parts[1].padStart(2, '0')}`;
    }
  }
  return str;
};

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

export default function MaxMinData() {
  const [feeders, setFeeders] = useState([]);
  const [selectedFeeder, setSelectedFeeder] = useState(null);
  const [entries, setEntries] = useState([]);
  const [year, setYear] = useState(new Date().getFullYear());
  const [month, setMonth] = useState(new Date().getMonth() + 1);
  const [showDateSelector, setShowDateSelector] = useState(true);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingEntry, setEditingEntry] = useState(null);
  const [loading, setLoading] = useState(false);
  const [initialized, setInitialized] = useState(false);
  const [ictEntries, setIctEntries] = useState({});
  const [importPreviewOpen, setImportPreviewOpen] = useState(false);
  const [importData, setImportData] = useState([]);
  const fileInputRef = useRef(null);

  useEffect(() => {
    initializeModule();
  }, []);

  const initializeModule = async () => {
    try {
      await axios.post(`${API}/max-min/init`);
      const response = await axios.get(`${API}/max-min/feeders`);
      setFeeders(response.data);
      setInitialized(true);
    } catch (error) {
      console.error('Failed to initialize max-min module:', error);
      toast.error('Failed to load feeders');
    }
  };

  const getSortedFeeders = () => {
    if (!feeders) return [];
    return [...feeders].sort((a, b) => {
      const indexA = FEEDER_ORDER.indexOf(a.name);
      const indexB = FEEDER_ORDER.indexOf(b.name);
      if (indexA !== -1 && indexB !== -1) return indexA - indexB;
      if (indexA !== -1) return -1;
      if (indexB !== -1) return 1;
      return a.name.localeCompare(b.name);
    });
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

  const handleSubmitDateSelection = () => {
    if (feeders.length > 0) {
      const sorted = getSortedFeeders();
      const feederToUse = selectedFeeder || sorted[0];
      setSelectedFeeder(feederToUse);
      setShowDateSelector(false);
      fetchEntries(feederToUse.id, year, month);
    } else {
      toast.error("No feeders available");
    }
  };

  const fetchEntries = async (feederId, selectedYear, selectedMonth) => {
    setLoading(true);
    try {
      const response = await axios.get(`${API}/max-min/entries/${feederId}`, {
        params: { year: selectedYear, month: selectedMonth }
      });
      // Sort entries by date
      const sortedEntries = response.data.sort((a, b) => a.date.localeCompare(b.date));
      setEntries(sortedEntries);
    } catch (error) {
      console.error('Failed to fetch entries:', error);
      toast.error('Failed to load entries');
    } finally {
      setLoading(false);
    }
  };

  // Fetch ICT entries for Station Load calculation if on Bus Station page
  useEffect(() => {
    const fetchICTDataForStationLoad = async () => {
      const ictFeeders = feeders.filter(f => f.type === 'ict_feeder');
      const dataMap = {};
      
      await Promise.all(ictFeeders.map(async (feeder) => {
        try {
          const response = await axios.get(`${API}/max-min/entries/${feeder.id}`, {
            params: { year, month }
          });
          dataMap[feeder.id] = response.data;
        } catch (error) {
          console.error(`Failed to fetch ICT entries for ${feeder.name}`, error);
        }
      }));
      
      setIctEntries(dataMap);
    };

    if (selectedFeeder?.type === 'bus_station' && !showDateSelector) {
      fetchICTDataForStationLoad();
    }
  }, [selectedFeeder, year, month, feeders, showDateSelector]);

  const handleFeederChange = (feederId) => {
    const feeder = feeders.find(f => f.id === feederId);
    if (feeder) {
      setSelectedFeeder(feeder);
      fetchEntries(feeder.id, year, month);
    }
  };

  const handleSaveEntry = async (entryData) => {
    // entryData comes from Modal: { date, data: {...} }
    
    // Auto-calculate averages before saving
    const dataToSave = { ...entryData.data };
    
    if (selectedFeeder.type === 'ict_feeder' || selectedFeeder.type.startsWith('feeder_')) {
       const maxAmps = parseFloat(dataToSave.max?.amps || 0);
       const minAmps = parseFloat(dataToSave.min?.amps || 0);
       const maxMW = parseFloat(dataToSave.max?.mw || 0);
       const minMW = parseFloat(dataToSave.min?.mw || 0);
       
       if (!isNaN(maxAmps) && !isNaN(minAmps)) {
           if (!dataToSave.avg) dataToSave.avg = {};
           dataToSave.avg.amps = ((maxAmps + minAmps) / 2).toFixed(2);
       }
       if (!isNaN(maxMW) && !isNaN(minMW)) {
           if (!dataToSave.avg) dataToSave.avg = {};
           dataToSave.avg.mw = ((maxMW + minMW) / 2).toFixed(2);
       }
    }

    // For Bus Station, handle Station Load if needed.
    if (selectedFeeder.type === 'bus_station') {
        const calculatedLoad = getStationLoad(entryData.date);
        if (calculatedLoad) {
            dataToSave.station_load = {
                max_mw: calculatedLoad.max_mw,
                mvar: calculatedLoad.mvar,
                time: calculatedLoad.time
            };
        }
    }

    try {
      // Check if entry exists for this date
      const existingEntry = entries.find(e => e.date === entryData.date);
      let response;
      
      if (existingEntry) {
          response = await axios.put(`${API}/max-min/entries/${existingEntry.id}`, {
            feeder_id: selectedFeeder.id,
            date: entryData.date,
            data: dataToSave
          });
          toast.success('Data Updated');
      } else {
          response = await axios.post(`${API}/max-min/entries`, {
            feeder_id: selectedFeeder.id,
            date: entryData.date,
            data: dataToSave
          });
          toast.success('Data Saved');
      }
      
      // Update local state
      const existingIndex = entries.findIndex(e => e.date === entryData.date);
      let newEntries;
      if (existingIndex !== -1) {
        newEntries = [...entries];
        newEntries[existingIndex] = response.data;
      } else {
        newEntries = [...entries, response.data];
      }
      setEntries(newEntries.sort((a, b) => a.date.localeCompare(b.date)));
      setEditingEntry(null);
    } catch (error) {
      console.error('Failed to save entry:', error);
      toast.error('Failed to save data');
      throw error; // Re-throw for modal to handle loading state
    }
  };

  const handleDeleteEntry = async (entryId) => {
    if (!confirm('Are you sure you want to delete this entry?')) return;
    try {
      await axios.delete(`${API}/max-min/entries/${entryId}`);
      setEntries(entries.filter(e => e.id !== entryId));
      toast.success('Entry deleted');
    } catch (error) {
      console.error('Failed to delete entry:', error);
      toast.error('Failed to delete entry');
    }
  };

  const handleExport = async () => {
    if (!selectedFeeder) return;
    try {
      const response = await axios.get(
        `${API}/max-min/export/${selectedFeeder.id}/${year}/${month}`,
        { responseType: 'blob' }
      );
      
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `${selectedFeeder.name}_${year}_${month.toString().padStart(2, '0')}.xlsx`);
      document.body.appendChild(link);
      link.click();
      link.remove();
    } catch (error) {
      toast.error('Failed to export data');
    }
  };
  
  const handleRefresh = () => {
    if (!selectedFeeder) return;
    fetchEntries(selectedFeeder.id, year, month);
    toast.success('Data refreshed');
  };

  const handleExportAll = async () => {
    try {
      const response = await axios.get(
        `${API}/max-min/export-all/${year}/${month}`,
        { responseType: 'blob' }
      );
      
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `MaxMin_All_${year}_${month.toString().padStart(2, '0')}.xlsx`);
      document.body.appendChild(link);
      link.click();
      link.remove();
    } catch (error) {
      console.error('Failed to export all data:', error);
      toast.error('Failed to export all data');
    }
  };
 
  const handleImportClick = () => {
    if (!selectedFeeder) return;
    fileInputRef.current?.click();
  };
 
  const handleFileSelect = async (event) => {
    const file = event.target.files?.[0];
    if (!file || !selectedFeeder) return;
    event.target.value = '';
    const formData = new FormData();
    formData.append('file', file);
    setLoading(true);
    try {
      const response = await axios.post(`${API}/max-min/preview-import/${selectedFeeder.id}`, formData, {
        headers: { 'Content-Type': 'multipart/form-data' }
      });
      setImportData(response.data);
      setImportPreviewOpen(true);
    } catch (error) {
      console.error('Import preview failed:', error);
      toast.error(error.response?.data?.detail || 'Failed to preview import');
    } finally {
      setLoading(false);
    }
  };
 
  const handleImportConfirm = async () => {
    if (!selectedFeeder) return;
    setLoading(true);
    try {
      await axios.post(`${API}/max-min/import-entries`, {
        feeder_id: selectedFeeder.id,
        entries: importData
      });
      toast.success('Data imported successfully');
      setImportPreviewOpen(false);
      fetchEntries(selectedFeeder.id, year, month);
    } catch (error) {
      console.error('Import failed:', error);
      toast.error('Failed to import data');
    } finally {
      setLoading(false);
    }
  };

  const getStationLoad = (date) => {
    let totalMaxMW = 0;
    let totalMVAR = 0;
    let commonTime = '-';
    let hasData = false;
    
    Object.values(ictEntries).forEach(feederEntries => {
      const entry = feederEntries.find(e => e.date === date);
      const maxMW = entry?.data?.max?.mw;
      
      if (maxMW && maxMW !== 'N/S') {
        totalMaxMW += parseFloat(maxMW) || 0;
        hasData = true;
        
        // Capture the common time from the first valid entry
        if (entry.data.max.time && entry.data.max.time !== 'N/S' && commonTime === '-') {
            commonTime = entry.data.max.time;
        }
      }
      
      const maxMVAR = entry?.data?.max?.mvar;
      if (maxMVAR && maxMVAR !== 'N/S') {
        totalMVAR += parseFloat(maxMVAR) || 0;
      }
    });
    
    if (!hasData) return null;
    return { max_mw: totalMaxMW.toFixed(2), mvar: totalMVAR.toFixed(2), time: commonTime };
  };

  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];

  // Helper to safely get nested values
  const getVal = (obj, path) => {
      return path.split('.').reduce((acc, part) => acc && acc[part], obj);
  };

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
      {/* Header */}
      <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-6 bg-white dark:bg-slate-950 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-800">
        <div>
          <h1 className="text-3xl md:text-4xl font-heading font-extrabold text-transparent bg-clip-text bg-gradient-to-r from-indigo-600 via-purple-600 to-pink-600">
            Max Min Data
          </h1>
          <p className="text-sm md:text-base text-slate-500 font-medium mt-2 flex items-center gap-2">
            <span className="inline-block w-2.5 h-2.5 rounded-full bg-emerald-500 ring-4 ring-emerald-50"></span>
            {monthNames[month - 1]} {year}
          </p>
        </div>
        
        <div className="flex flex-wrap items-center gap-3 w-full lg:w-auto">
          <Button 
            variant="outline" 
            onClick={handleRefresh}
            disabled={!selectedFeeder}
            className="flex-1 lg:flex-none border-slate-200 hover:bg-slate-50 text-slate-700 font-medium"
            data-testid="refresh-button"
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
            onClick={() => {
                setEditingEntry(null);
                setIsModalOpen(true);
            }}
            className="flex-1 lg:flex-none bg-gradient-to-r from-indigo-600 to-purple-600 hover:from-indigo-700 hover:to-purple-700 text-white shadow-lg shadow-indigo-200 border-0 font-medium transition-all hover:scale-105"
          >
            <Plus className="w-4 h-4 mr-2" />
            New Entry
          </Button>
          <Button 
            onClick={handleImportClick}
            variant="outline"
            className="flex-1 lg:flex-none border-slate-200 hover:bg-slate-50 text-slate-700 font-medium"
          >
            <Upload className="w-4 h-4 mr-2 text-indigo-500" />
            Import
          </Button>
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileSelect}
            className="hidden"
            accept=".xlsx,.xls"
          />
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
      </div>

      {/* Feeder Selection */}
      <Card className="w-full max-w-xl mb-8 border-0 shadow-md ring-1 ring-slate-100 overflow-hidden">
        <div className="h-1 w-full bg-gradient-to-r from-indigo-500 to-purple-500"></div>
        <CardContent className="pt-6 pb-6">
            <label className="text-xs font-bold uppercase tracking-wider text-slate-500 mb-3 block flex items-center gap-2">
              <div className="w-1.5 h-1.5 rounded-full bg-indigo-500"></div>
              Select Feeder
            </label>
            <div className="flex items-center gap-2">
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
        </CardContent>
      </Card>

      {/* Data Table */}
      {selectedFeeder && (
        <Card className="shadow-lg border-0 ring-1 ring-slate-100 overflow-hidden bg-white relative">
            {loading && <BlockLoader text="Loading entries..." className="z-20" />}
            <CardHeader className="bg-slate-50/50 border-b border-slate-100 py-5">
                <div className="flex items-center justify-between">
                    <CardTitle className="text-lg font-bold text-slate-700 flex items-center gap-2">
                        <div className="p-2 bg-indigo-100 text-indigo-600 rounded-lg">
                            <Calendar className="w-5 h-5" />
                        </div>
                        Daily Records
                    </CardTitle>
                    <span className="text-sm font-semibold text-slate-500 bg-white px-4 py-1.5 rounded-full border shadow-sm">
                        {entries.length} Entries
                    </span>
                </div>
            </CardHeader>
            <CardContent className="p-0 overflow-x-auto">
                <Table>
                    <TableHeader>
                        <TableRow className="bg-slate-100 border-b border-slate-200">
                            <TableHead className="w-[120px] font-bold text-slate-700 text-center">Date</TableHead>
                            {selectedFeeder.type === 'bus_station' ? (
                                <>
                                    <TableHead colSpan={3} className="text-center border-l border-slate-200 text-rose-700 font-bold">Maximum Voltages</TableHead>
                                    <TableHead colSpan={3} className="text-center border-l border-slate-200 text-blue-700 font-bold">Minimum Voltages</TableHead>
                                    <TableHead colSpan={3} className="text-center border-l border-slate-200 text-purple-700 font-bold">Station Load</TableHead>
                                </>
                            ) : (
                                <>
                                    <TableHead colSpan={selectedFeeder.type === 'ict_feeder' ? 4 : 3} className="text-center border-l border-slate-200 text-rose-700 font-bold">Maximum</TableHead>
                                    <TableHead colSpan={selectedFeeder.type === 'ict_feeder' ? 4 : 3} className="text-center border-l border-slate-200 text-blue-700 font-bold">Minimum</TableHead>
                                    <TableHead colSpan={2} className="text-center border-l border-slate-200 text-emerald-700 font-bold">Average</TableHead>
                                </>
                            )}
                            <TableHead className="text-center border-l border-slate-200 font-bold text-slate-700">Actions</TableHead>
                        </TableRow>
                        <TableRow className="bg-slate-50 border-b border-slate-200 text-xs uppercase tracking-wider text-slate-500 font-semibold">
                            <TableHead className="text-center"></TableHead>
                            {selectedFeeder.type === 'bus_station' ? (
                                <>
                                    <TableHead className="border-l border-slate-200 text-rose-600 text-center">400KV</TableHead><TableHead className="text-rose-600 text-center">220KV</TableHead><TableHead className="text-rose-600 text-center">Time</TableHead>{/*
                                    */}<TableHead className="border-l border-slate-200 text-blue-600 text-center">400KV</TableHead><TableHead className="text-blue-600 text-center">220KV</TableHead><TableHead className="text-blue-600 text-center">Time</TableHead>{/*
                                    */}<TableHead className="border-l border-slate-200 text-purple-600 text-center">Max MW</TableHead><TableHead className="text-purple-600 text-center">MVAR</TableHead><TableHead className="text-purple-600 text-center">Time</TableHead>
                                </>
                            ) : (
                                <>
                                    <TableHead className="border-l border-slate-200 text-rose-600 text-center">Amps</TableHead><TableHead className="text-rose-600 text-center">MW</TableHead>{/*
                                    */}{selectedFeeder.type === 'ict_feeder' && <TableHead className="text-rose-600 text-center">MVAR</TableHead>}{/*
                                    */}<TableHead className="text-rose-600 text-center">Time</TableHead>{/*
                                    
                                    */}<TableHead className="border-l border-slate-200 text-blue-600 text-center">Amps</TableHead><TableHead className="text-blue-600 text-center">MW</TableHead>{/*
                                    */}{selectedFeeder.type === 'ict_feeder' && <TableHead className="text-blue-600 text-center">MVAR</TableHead>}{/*
                                    */}<TableHead className="text-blue-600 text-center">Time</TableHead>{/*
                                    
                                    */}<TableHead className="border-l border-slate-200 text-emerald-600 text-center">Amps</TableHead><TableHead className="text-emerald-600 text-center">MW</TableHead>
                                </>
                            )}
                            <TableHead className="border-l border-slate-200 text-center"></TableHead>
                        </TableRow>
                    </TableHeader>
                    <TableBody>
                        {entries.length === 0 ? (
                            <TableRow>
                                <TableCell colSpan={20} className="text-center py-8 text-muted-foreground">
                                    No entries found for this month. Click "Entry" to add data.
                                </TableCell>
                            </TableRow>
                        ) : (
                            entries.map((entry) => (
                                <TableRow key={entry.id}>
                                    <TableCell className="font-medium whitespace-nowrap text-center">{formatDate(entry.date)}</TableCell>
                                    {selectedFeeder.type === 'bus_station' ? (
                                        <>
                                            <TableCell className="border-l text-center">{getVal(entry.data, 'max_bus_voltage_400kv.value')}</TableCell>{/*
                                            */}<TableCell className="text-center">{getVal(entry.data, 'max_bus_voltage_220kv.value')}</TableCell>{/*
                                            */}<TableCell className="text-center">{formatTime(getVal(entry.data, 'max_bus_voltage.time') || getVal(entry.data, 'max_bus_voltage_400kv.time') || getVal(entry.data, 'max_bus_voltage_220kv.time'))}</TableCell>{/*
                                            
                                            */}<TableCell className="border-l text-center">{getVal(entry.data, 'min_bus_voltage_400kv.value')}</TableCell>{/*
                                            */}<TableCell className="text-center">{getVal(entry.data, 'min_bus_voltage_220kv.value')}</TableCell>{/*
                                            */}<TableCell className="text-center">{formatTime(getVal(entry.data, 'min_bus_voltage.time') || getVal(entry.data, 'min_bus_voltage_400kv.time') || getVal(entry.data, 'min_bus_voltage_220kv.time'))}</TableCell>{/*
                                            
                                            */}<TableCell className="border-l text-center">{getStationLoad(entry.date)?.max_mw || getVal(entry.data, 'station_load.max_mw')}</TableCell>{/*
                                            */}<TableCell className="text-center">{getStationLoad(entry.date)?.mvar || getVal(entry.data, 'station_load.mvar')}</TableCell>{/*
                                            */}<TableCell className="text-center">{formatTime(getStationLoad(entry.date)?.time || getVal(entry.data, 'station_load.time'))}</TableCell>
                                        </>
                                    ) : (
                                        <>
                                            <TableCell className="border-l text-center">{getVal(entry.data, 'max.amps')}</TableCell>{/*
                                            */}<TableCell className="text-center">{getVal(entry.data, 'max.mw')}</TableCell>{/*
                                            */}{selectedFeeder.type === 'ict_feeder' && <TableCell className="text-center">{getVal(entry.data, 'max.mvar')}</TableCell>}{/*
                                            */}<TableCell className="text-center">{formatTime(getVal(entry.data, 'max.time'))}</TableCell>{/*
                                            
                                            */}<TableCell className="border-l text-center">{getVal(entry.data, 'min.amps')}</TableCell>{/*
                                            */}<TableCell className="text-center">{getVal(entry.data, 'min.mw')}</TableCell>{/*
                                            */}{selectedFeeder.type === 'ict_feeder' && <TableCell className="text-center">{getVal(entry.data, 'min.mvar')}</TableCell>}{/*
                                            */}<TableCell className="text-center">{formatTime(getVal(entry.data, 'min.time'))}</TableCell>{/*
                                            
                                            */}<TableCell className="border-l text-center">{getVal(entry.data, 'avg.amps')}</TableCell>{/*
                                            */}<TableCell className="text-center">{getVal(entry.data, 'avg.mw')}</TableCell>
                                        </>
                                    )}
                                    <TableCell className="text-center border-l">
                                        <div className="flex justify-center gap-2">
                                            <Button variant="ghost" size="icon" onClick={() => {
                                                setEditingEntry(entry);
                                                setIsModalOpen(true);
                                            }}>
                                                <Edit className="w-4 h-4" />
                                            </Button>
                                            <Button variant="ghost" size="icon" className="text-red-500 hover:text-red-600" onClick={() => handleDeleteEntry(entry.id)}>
                                                <Trash2 className="w-4 h-4" />
                                            </Button>
                                        </div>
                                    </TableCell>
                                </TableRow>
                            ))
                        )}
                    </TableBody>
                </Table>
            </CardContent>
        </Card>
      )}
 
      <MaxMinImportPreviewModal
        isOpen={importPreviewOpen}
        onClose={() => setImportPreviewOpen(false)}
        data={importData}
        feederType={selectedFeeder?.type}
        onConfirm={handleImportConfirm}
        loading={loading}
      />

      {/* Summary Table */}
      {selectedFeeder && entries.length > 0 && (
          <SummaryTable entries={entries} selectedFeeder={selectedFeeder} year={year} month={month} getStationLoad={getStationLoad} />
      )}
      
      {selectedFeeder && entries.length > 0 && (
        <MaxMinAnalytics entries={entries} selectedFeeder={selectedFeeder} />
      )}

      {/* Modal */}
      {isModalOpen && (
        <MaxMinEntryModal
            isOpen={isModalOpen}
            onClose={() => setIsModalOpen(false)}
            onSave={handleSaveEntry}
            feeder={selectedFeeder}
            year={year}
            month={month}
            entries={entries}
            initialData={editingEntry}
            defaultDate={(() => {
              if (entries.length === 0) {
                return `${year}-${String(month).padStart(2, '0')}-01`;
              }
              const sorted = [...entries].sort((a, b) => a.date.localeCompare(b.date));
              const lastDate = new Date(sorted[sorted.length - 1].date);
              lastDate.setDate(lastDate.getDate() + 1);
              return lastDate.toISOString().split('T')[0];
            })()}
            onPrevFeeder={goToPrevFeeder}
            onNextFeeder={goToNextFeeder}
        />
      )}
    </div>
  );
}

// Summary Table Component
function SummaryTable({ entries, selectedFeeder, year, month, getStationLoad }) {
      const [backendSummary, setBackendSummary] = useState(null);

      useEffect(() => {
          setBackendSummary(null);
          if (selectedFeeder && selectedFeeder.type !== 'bus_station') {
              axios.get(`${API}/max-min/summary/${selectedFeeder.id}/${year}/${month}`)
                  .then(res => {
                      if (res.data) setBackendSummary(res.data);
                  })
                  .catch(() => setBackendSummary(null));
          }
      }, [selectedFeeder, year, month]);

      const p1Start = `${year}-${month.toString().padStart(2, '0')}-01`;
      const p1End = `${year}-${month.toString().padStart(2, '0')}-15`;
      const p2Start = `${year}-${month.toString().padStart(2, '0')}-16`;
      const p2End = `${year}-${month.toString().padStart(2, '0')}-${new Date(year, month, 0).getDate()}`;
      const monthStart = p1Start;
      const monthEnd = p2End;
      
      const periods = [
          { name: "1st to 15th", start: p1Start, end: p1End },
          { name: "16th to End", start: p2Start, end: p2End },
          { name: "Full Month", start: monthStart, end: monthEnd }
      ];

      if (selectedFeeder.type === 'bus_station') {
          return (
            <Card className="mt-8 shadow-lg border-0 ring-1 ring-slate-100 overflow-hidden">
                <CardHeader className="bg-gradient-to-r from-slate-50 to-slate-100 border-b border-slate-200 py-4">
                    <CardTitle className="text-lg font-bold text-slate-700 flex items-center gap-2">
                        <div className="p-2 bg-rose-100 text-rose-600 rounded-lg">
                            <Calendar className="w-5 h-5" />
                        </div>
                        Monthly Summary Report
                    </CardTitle>
                </CardHeader>
                <CardContent className="p-0">
                    <Table>
                        <TableHeader>
                            <TableRow className="bg-slate-100 border-b border-slate-200">
                                <TableHead className="font-bold text-slate-700 text-center">Period</TableHead>
                                <TableHead className="font-bold text-slate-700 text-center">Parameter</TableHead>
                                <TableHead className="text-center border-l border-slate-200 text-rose-700 font-bold">Maximum Value</TableHead>
                                <TableHead className="text-rose-700 font-bold text-center">Date</TableHead>
                                <TableHead className="text-rose-700 font-bold text-center">Time</TableHead>
                                <TableHead className="text-center border-l border-slate-200 text-blue-700 font-bold">Minimum Value</TableHead>
                                <TableHead className="text-blue-700 font-bold text-center">Date</TableHead>
                                <TableHead className="text-blue-700 font-bold text-center">Time</TableHead>
                            </TableRow>
                        </TableHeader>
                        <TableBody>
                            {periods.map(p => {
                                // 1. Filter entries for this period
                                const pEntries = entries.filter(e => e.date >= p.start && e.date <= p.end);
                                
                                // 2. Find Max/Min Values
                                let max400Val = -Infinity, min400Val = Infinity;
                                let max220Val = -Infinity, min220Val = Infinity;
                                let maxLoadVal = -Infinity;
                                
                                pEntries.forEach(e => {
                                    const v400Max = parseFloat(e.data?.max_bus_voltage_400kv?.value);
                                    if (!isNaN(v400Max) && v400Max > max400Val) max400Val = v400Max;
                                    
                                    const v400Min = parseFloat(e.data?.min_bus_voltage_400kv?.value);
                                    if (!isNaN(v400Min) && v400Min < min400Val) min400Val = v400Min;
                                    
                                    const v220Max = parseFloat(e.data?.max_bus_voltage_220kv?.value);
                                    if (!isNaN(v220Max) && v220Max > max220Val) max220Val = v220Max;
                                    
                                    const v220Min = parseFloat(e.data?.min_bus_voltage_220kv?.value);
                                    if (!isNaN(v220Min) && v220Min < min220Val) min220Val = v220Min;
                                    
                                    const loadVal = parseFloat(getStationLoad(e.date)?.max_mw || e.data?.station_load?.max_mw);
                                    if (!isNaN(loadVal) && loadVal > maxLoadVal) maxLoadVal = loadVal;
                                });

                                // 3. Collect Candidates
                                const candsMax400 = [], candsMin400 = [];
                                const candsMax220 = [], candsMin220 = [];
                                let maxLoadEntry = { date: '-', time: '-' };

                                pEntries.forEach(e => {
                                    const v400Max = parseFloat(e.data?.max_bus_voltage_400kv?.value);
                                    if (v400Max === max400Val) candsMax400.push({ date: e.date, time: (e.data?.max_bus_voltage_400kv?.time || '').trim() });
                                    
                                    const v400Min = parseFloat(e.data?.min_bus_voltage_400kv?.value);
                                    if (v400Min === min400Val) candsMin400.push({ date: e.date, time: (e.data?.min_bus_voltage_400kv?.time || '').trim() });

                                    const v220Max = parseFloat(e.data?.max_bus_voltage_220kv?.value);
                                    if (v220Max === max220Val) candsMax220.push({ date: e.date, time: (e.data?.max_bus_voltage_220kv?.time || '').trim() });

                                    const v220Min = parseFloat(e.data?.min_bus_voltage_220kv?.value);
                                    if (v220Min === min220Val) candsMin220.push({ date: e.date, time: (e.data?.min_bus_voltage_220kv?.time || '').trim() });

                                    const loadVal = parseFloat(getStationLoad(e.date)?.max_mw || e.data?.station_load?.max_mw);
                                    if (loadVal === maxLoadVal) {
                                        maxLoadEntry = { 
                                            date: e.date, 
                                            time: getStationLoad(e.date)?.time || e.data?.station_load?.time || '-'
                                        };
                                    }
                                });

                                // 4. Find Best Matches (Common Time Priority)
                                const findBest = (listA, listB) => {
                                    for (const a of listA) {
                                        for (const b of listB) {
                                            if (a.date === b.date && a.time === b.time) return [a, b];
                                        }
                                    }
                                    return [listA[0] || { date: '-', time: '-' }, listB[0] || { date: '-', time: '-' }];
                                };

                                const [finalMax400, finalMax220] = findBest(candsMax400, candsMax220);
                                const [finalMin400, finalMin220] = findBest(candsMin400, candsMin220);

                                return (
                                    <Fragment key={p.name}>
                                        <TableRow key={`${p.name}-400`} className="hover:bg-slate-50/50">
                                            <TableCell rowSpan={3} className="font-medium text-center bg-slate-50/30">{p.name}</TableCell>
                                            <TableCell className="font-medium text-slate-600 text-center">400KV Bus Voltage</TableCell>
                                            <TableCell className="border-l text-center font-medium text-rose-600">{max400Val === -Infinity ? '-' : max400Val}</TableCell>
                                            <TableCell className="text-center text-slate-500">{formatDate(finalMax400.date)}</TableCell>
                                            <TableCell className="text-center text-slate-500">{formatTime(finalMax400.time)}</TableCell>
                                            <TableCell className="border-l text-center font-medium text-blue-600">{min400Val === Infinity ? '-' : min400Val}</TableCell>
                                            <TableCell className="text-center text-slate-500">{formatDate(finalMin400.date)}</TableCell>
                                            <TableCell className="text-center text-slate-500">{formatTime(finalMin400.time)}</TableCell>
                                        </TableRow>
                                        <TableRow key={`${p.name}-220`} className="hover:bg-slate-50/50">
                                            <TableCell className="font-medium text-slate-600 text-center">220KV Bus Voltage</TableCell>
                                            <TableCell className="border-l text-center font-medium text-rose-600">{max220Val === -Infinity ? '-' : max220Val}</TableCell>
                                            <TableCell className="text-center text-slate-500">{formatDate(finalMax220.date)}</TableCell>
                                            <TableCell className="text-center text-slate-500">{formatTime(finalMax220.time)}</TableCell>
                                            <TableCell className="border-l text-center font-medium text-blue-600">{min220Val === Infinity ? '-' : min220Val}</TableCell>
                                            <TableCell className="text-center text-slate-500">{formatDate(finalMin220.date)}</TableCell>
                                            <TableCell className="text-center text-slate-500">{formatTime(finalMin220.time)}</TableCell>
                                        </TableRow>
                                        <TableRow key={`${p.name}-load`} className="hover:bg-slate-50/50 border-b border-slate-100">
                                            <TableCell className="font-medium text-slate-600 text-center">Station Load (MW)</TableCell>
                                            <TableCell className="border-l text-center font-medium text-rose-600">{maxLoadVal === -Infinity ? '-' : maxLoadVal}</TableCell>
                                            <TableCell className="text-center text-slate-500">{formatDate(maxLoadEntry.date)}</TableCell>
                                            <TableCell className="text-center text-slate-500">{formatTime(maxLoadEntry.time)}</TableCell>
                                            <TableCell colSpan={3} className="border-l text-center text-muted-foreground">-</TableCell>
                                        </TableRow>
                                    </Fragment>
                                );
                            })}
                        </TableBody>
                    </Table>
                </CardContent>
            </Card>
          );
      } else {
        // Feeder Summary
        return (
            <Card className="mt-8 shadow-lg border-0 ring-1 ring-slate-100 overflow-hidden">
                <CardHeader className="bg-gradient-to-r from-slate-50 to-slate-100 border-b border-slate-200 py-4">
                    <CardTitle className="text-lg font-bold text-slate-700 flex items-center gap-2">
                        <div className="p-2 bg-emerald-100 text-emerald-600 rounded-lg">
                            <Calendar className="w-5 h-5" />
                        </div>
                        Monthly Summary Report
                    </CardTitle>
                </CardHeader>
                <CardContent className="p-0">
                    <Table>
                        <TableHeader>
                            <TableRow className="bg-slate-100 border-b border-slate-200">
                                <TableHead className="font-bold text-slate-700 text-center">Period</TableHead>
                                <TableHead className="font-bold text-slate-700 text-center">Parameter</TableHead>
                                <TableHead className="text-center border-l border-slate-200 text-rose-700 font-bold">Maximum Value</TableHead>
                                <TableHead className="text-rose-700 font-bold text-center">Date</TableHead>
                                <TableHead className="text-rose-700 font-bold text-center">Time</TableHead>
                                <TableHead className="text-center border-l border-slate-200 text-blue-700 font-bold">Minimum Value</TableHead>
                                <TableHead className="text-blue-700 font-bold text-center">Date</TableHead>
                                <TableHead className="text-blue-700 font-bold text-center">Time</TableHead>
                                <TableHead className="text-center border-l border-slate-200 text-emerald-700 font-bold">Average Value</TableHead>
                            </TableRow>
                        </TableHeader>
                        <TableBody>
                            {periods.map(p => {
                                let maxMW = -Infinity, minMW = Infinity, maxMWDate, maxMWTime, minMWDate, minMWTime;
                                let totalMW = 0, countMW = 0;
                                let maxAmps = -Infinity, minAmps = Infinity, maxAmpsDate, maxAmpsTime, minAmpsDate, minAmpsTime;
                                let totalAmps = 0, countAmps = 0;
                                
                                let maxMWEntry = null;
                                let minMWEntry = null;

                                entries.forEach(e => {
                                    if (e.date >= p.start && e.date <= p.end) {
                                        // MW Calculations
                                        const valMaxMW = parseFloat(e.data?.max?.mw);
                                        const valMinMW = parseFloat(e.data?.min?.mw);
                                        let valAvgMW = parseFloat(e.data?.avg?.mw);
                                        
                                        // Auto-calculate average if missing but max/min exist
                                        if (isNaN(valAvgMW) && !isNaN(valMaxMW) && !isNaN(valMinMW)) {
                                            valAvgMW = (valMaxMW + valMinMW) / 2;
                                        }

                                        if (!isNaN(valMaxMW) && valMaxMW > maxMW) { 
                                            maxMW = valMaxMW; 
                                            maxMWEntry = e;
                                        }
                                        if (!isNaN(valMinMW) && valMinMW < minMW) { 
                                            minMW = valMinMW; 
                                            minMWEntry = e;
                                        }
                                        if (!isNaN(valAvgMW)) {
                                            totalMW += valAvgMW;
                                            countMW++;
                                        }

                                        // Amps Calculations (Only for Averages)
                                        const valMaxAmps = parseFloat(e.data?.max?.amps);
                                        const valMinAmps = parseFloat(e.data?.min?.amps);
                                        let valAvgAmps = parseFloat(e.data?.avg?.amps);

                                        if (isNaN(valAvgAmps) && !isNaN(valMaxAmps) && !isNaN(valMinAmps)) {
                                            valAvgAmps = (valMaxAmps + valMinAmps) / 2;
                                        }
                                        
                                        if (!isNaN(valAvgAmps)) {
                                            totalAmps += valAvgAmps;
                                            countAmps++;
                                        }
                                    }
                                });

                                // Apply Max MW Logic
                                if (maxMWEntry) {
                                    maxMWDate = maxMWEntry.date;
                                    maxMWTime = maxMWEntry.data?.max?.time;
                                    
                                    const valMaxAmps = parseFloat(maxMWEntry.data?.max?.amps);
                                    if (!isNaN(valMaxAmps)) {
                                        maxAmps = valMaxAmps;
                                        maxAmpsDate = maxMWEntry.date;
                                        maxAmpsTime = maxMWEntry.data?.max?.time;
                                    }
                                }

                                // Apply Min MW Logic
                                if (minMWEntry) {
                                    minMWDate = minMWEntry.date;
                                    minMWTime = minMWEntry.data?.min?.time;
                                    
                                    const valMinAmps = parseFloat(minMWEntry.data?.min?.amps);
                                    if (!isNaN(valMinAmps)) {
                                        minAmps = valMinAmps;
                                        minAmpsDate = minMWEntry.date;
                                        minAmpsTime = minMWEntry.data?.min?.time;
                                    }
                                }

                                let avgMW = countMW > 0 ? (totalMW / countMW).toFixed(2) : '-';
                                let avgAmps = countAmps > 0 ? (totalAmps / countAmps).toFixed(2) : '-';

                                if (backendSummary) {
                                    const s = backendSummary.find(x => x.name === p.name);
                                    if (s) {
                                        maxMW = s.max_mw; maxMWDate = s.max_mw_date; maxMWTime = s.max_mw_time;
                                        minMW = s.min_mw; minMWDate = s.min_mw_date; minMWTime = s.min_mw_time;
                                        
                                        maxAmps = s.max_amps; maxAmpsDate = s.max_amps_date; maxAmpsTime = s.max_amps_time;
                                        minAmps = s.min_amps; minAmpsDate = s.min_amps_date; minAmpsTime = s.min_amps_time;
                                        
                                        avgMW = s.avg_mw;
                                        avgAmps = s.avg_amps;
                                    }
                                }

                                return (
                                    <Fragment key={p.name}>
                                        <TableRow key={`${p.name}-amps`} className="hover:bg-slate-50/50">
                                            <TableCell rowSpan={2} className="font-medium text-center bg-slate-50/30">{p.name}</TableCell>
                                            <TableCell className="font-medium text-slate-600 text-center">Amps</TableCell>
                                            <TableCell className="border-l text-center font-medium text-rose-600">{maxAmps === -Infinity ? '-' : maxAmps}</TableCell>
                                            <TableCell className="text-center text-slate-500">{formatDate(maxAmpsDate)}</TableCell>
                                            <TableCell className="text-center text-slate-500">{formatTime(maxAmpsTime)}</TableCell>
                                            <TableCell className="border-l text-center font-medium text-blue-600">{minAmps === Infinity ? '-' : minAmps}</TableCell>
                                            <TableCell className="text-center text-slate-500">{formatDate(minAmpsDate)}</TableCell>
                                            <TableCell className="text-center text-slate-500">{formatTime(minAmpsTime)}</TableCell>
                                            <TableCell className="border-l text-center font-bold text-emerald-600">{avgAmps}</TableCell>
                                        </TableRow>
                                        <TableRow key={`${p.name}-mw`} className="hover:bg-slate-50/50 border-b border-slate-100">
                                            <TableCell className="font-medium text-slate-600 text-center">MW</TableCell>
                                            <TableCell className="border-l text-center font-medium text-rose-600">{maxMW === -Infinity ? '-' : maxMW}</TableCell>
                                            <TableCell className="text-center text-slate-500">{formatDate(maxMWDate)}</TableCell>
                                            <TableCell className="text-center text-slate-500">{formatTime(maxMWTime)}</TableCell>
                                            <TableCell className="border-l text-center font-medium text-blue-600">{minMW === Infinity ? '-' : minMW}</TableCell>
                                            <TableCell className="text-center text-slate-500">{formatDate(minMWDate)}</TableCell>
                                            <TableCell className="text-center text-slate-500">{formatTime(minMWTime)}</TableCell>
                                            <TableCell className="border-l text-center font-bold text-emerald-600">{avgMW}</TableCell>
                                        </TableRow>
                                    </Fragment>
                                );
                            })}
                        </TableBody>
                    </Table>
                </CardContent>
            </Card>
        );
      }
}
