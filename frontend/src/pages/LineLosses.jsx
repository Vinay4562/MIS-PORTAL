import { useState, useEffect, useRef } from 'react';
import axios from 'axios';
import { toast } from 'sonner';
import { Button } from '@/components/ui/button';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import DataEntryModal from '@/components/DataEntryModal';
import ImportPreviewModal from '@/components/ImportPreviewModal';
import FeederTable from '@/components/FeederTable';
import AnalyticsCharts from '@/components/AnalyticsCharts';
import { Download, Plus, Calendar, RefreshCcw, Upload, ChevronLeft, ChevronRight } from 'lucide-react';

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

const FEEDER_ORDER = [
  "400 KV Shanakrapally-MHRM-2",
  "400 KV Shanakrapally-MHRM-1",
  "400 KV Shanakrapally-Narsapur-1",
  "400 KV Shanakrapally-Narsapur-2",
  "400 KV KethiReddyPally-1",
  "400 KV KethiReddyPally-2",
  "400 KV Nizamabad-1&2",
  "220 KV Parigi-1",
  "220 KV Parigi-2",
  "220 KV Tandur",
  "220 KV Gachibowli-1",
  "220 KV Gachibowli-2",
  "220 KV KethiReddyPally",
  "220 KV Yeddumailaram-1",
  "220 KV Yeddumailaram-2",
  "220 KV Sadasivapet-1",
  "220 KV Sadasivapet-2"
];

export default function LineLosses() {
  const [feeders, setFeeders] = useState([]);
  const [selectedFeeder, setSelectedFeeder] = useState(null);
  const [entries, setEntries] = useState([]);
  const [year, setYear] = useState(new Date().getFullYear());
  const [month, setMonth] = useState(new Date().getMonth() + 1);
  const [showDateSelector, setShowDateSelector] = useState(true);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [loading, setLoading] = useState(false);
  const [initialized, setInitialized] = useState(false);
  const [importPreviewOpen, setImportPreviewOpen] = useState(false);
  const [importData, setImportData] = useState([]);
  const fileInputRef = useRef(null);

  useEffect(() => {
    initializeFeeders();
  }, []);

  const initializeFeeders = async () => {
    try {
      const response = await axios.get(`${API}/feeders`);
      if (response.data.length === 0) {
        await axios.post(`${API}/init-feeders`);
        const feedersResponse = await axios.get(`${API}/feeders`);
        setFeeders(feedersResponse.data);
      } else {
        setFeeders(response.data);
      }
      setInitialized(true);
    } catch (error) {
      console.error('Failed to initialize feeders:', error);
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
      handleFeederChange(sorted[idx - 1]);
    }
  };

  const goToNextFeeder = () => {
    const sorted = getSortedFeeders();
    const idx = sorted.findIndex(f => f.id === selectedFeeder?.id);
    if (idx !== -1 && idx < sorted.length - 1) {
      handleFeederChange(sorted[idx + 1]);
    }
  };

  const handleSubmitDateSelection = () => {
    if (feeders.length > 0) {
      const sorted = getSortedFeeders();
      setSelectedFeeder(sorted[0]);
      setShowDateSelector(false);
      fetchEntries(sorted[0].id, year, month);
    }
  };

  const fetchEntries = async (feederId, selectedYear, selectedMonth) => {
    setLoading(true);
    try {
      const response = await axios.get(`${API}/entries`, {
        params: { feeder_id: feederId, year: selectedYear, month: selectedMonth }
      });
      setEntries(response.data);
    } catch (error) {
      console.error('Failed to fetch entries:', error);
      toast.error('Failed to load entries');
    } finally {
      setLoading(false);
    }
  };

  const handleFeederChange = (feeder) => {
    setSelectedFeeder(feeder);
    fetchEntries(feeder.id, year, month);
  };

  const handleEntryCreated = (newEntry) => {
    setEntries([...entries, newEntry].sort((a, b) => a.date.localeCompare(b.date)));
    toast.success('Data Saved');
  };

  const handleEntryUpdated = (updatedEntry) => {
    setEntries(entries.map(e => e.id === updatedEntry.id ? updatedEntry : e));
    toast.success('Entry updated successfully');
  };

  const handleEntryDeleted = async (entryId) => {
    try {
      await axios.delete(`${API}/entries/${entryId}`);
      setEntries(entries.filter(e => e.id !== entryId));
      toast.success('Entry deleted successfully');
    } catch (error) {
      toast.error('Failed to delete entry');
    }
  };

  const handleExport = async () => {
    try {
      const response = await axios.get(
        `${API}/export/${selectedFeeder.id}/${year}/${month}`,
        { responseType: 'blob' }
      );
      
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `${selectedFeeder.name}_${year}_${month}.xlsx`);
      document.body.appendChild(link);
      link.click();
      link.remove();
      
      toast.success('Export completed successfully');
    } catch (error) {
      console.error('Export failed:', error);
      toast.error('Failed to export data');
    }
  };

  const handleExportAll = async () => {
    try {
      const response = await axios.get(
        `${API}/line-losses/export-all/${year}/${month}`,
        { responseType: 'blob' }
      );
      
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `Line_Losses_All_${month}-${year}.xlsx`);
      document.body.appendChild(link);
      link.click();
      link.remove();
      
      toast.success('Export All completed successfully');
    } catch (error) {
      console.error('Export All failed:', error);
      toast.error('Failed to export all data');
    }
  };
  
  const handleRefresh = () => {
    if (!selectedFeeder) return;
    fetchEntries(selectedFeeder.id, year, month);
    toast.success('Data refreshed');
  };

  const handleImportClick = () => {
    fileInputRef.current?.click();
  };

  const handleFileSelect = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    
    // Reset input
    event.target.value = '';
    
    const formData = new FormData();
    formData.append('file', file);
    
    setLoading(true);
    try {
      const response = await axios.post(`${API}/preview-import/${selectedFeeder.id}`, formData, {
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
    setLoading(true);
    try {
      await axios.post(`${API}/import-entries`, {
        feeder_id: selectedFeeder.id,
        entries: importData
      });
      toast.success('Data imported successfully');
      setImportPreviewOpen(false);
      fetchEntries(selectedFeeder.id, year, month); // Refresh
    } catch (error) {
      console.error('Import failed:', error);
      toast.error('Failed to import data');
    } finally {
      setLoading(false);
    }
  };

  /* Removed duplicate code */

  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];

  if (!initialized) {
    return (
      <div className="flex items-center justify-center h-full">
        <p className="text-lg">Loading...</p>
      </div>
    );
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
              <label className="text-sm font-medium" data-testid="year-label">Year</label>
              <Select value={year.toString()} onValueChange={(v) => setYear(parseInt(v))}>
                <SelectTrigger data-testid="year-selector">
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
              <label className="text-sm font-medium" data-testid="month-label">Month</label>
              <Select value={month.toString()} onValueChange={(v) => setMonth(parseInt(v))}>
                <SelectTrigger data-testid="month-selector">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  {monthNames.map((name, idx) => (
                    <SelectItem key={idx + 1} value={(idx + 1).toString()}>{name}</SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            <Button 
              onClick={handleSubmitDateSelection} 
              className="w-full" 
              size="lg"
              data-testid="submit-date-button"
            >
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
      <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4">
        <div>
          <h1 className="text-2xl md:text-3xl font-heading font-bold text-slate-900 dark:text-slate-100">
            Line Losses Management
          </h1>
          <p className="text-sm md:text-base text-slate-600 dark:text-slate-400 mt-1">
            {monthNames[month - 1]} {year}
          </p>
        </div>
        
        <div className="flex flex-wrap items-center gap-2 md:gap-3 w-full lg:w-auto">
          <Button 
            variant="outline" 
            onClick={handleRefresh}
            disabled={!selectedFeeder}
            data-testid="refresh-button"
            className="flex-1 lg:flex-none"
          >
            <RefreshCcw className="w-4 h-4 mr-2" />
            Refresh
          </Button>
          <Button 
            variant="outline" 
            onClick={() => setShowDateSelector(true)}
            data-testid="change-period-button"
            className="flex-1 lg:flex-none"
          >
            <Calendar className="w-4 h-4 mr-2" />
            Period
          </Button>
          <Button 
            onClick={() => setIsModalOpen(true)}
            data-testid="new-entry-button"
            className="flex-1 lg:flex-none"
          >
            <Plus className="w-4 h-4 mr-2" />
            Entry
          </Button>
          <Button 
            onClick={handleImportClick}
            variant="outline"
            className="flex-1 lg:flex-none"
          >
            <Upload className="w-4 h-4 mr-2" />
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
            data-testid="export-button"
            className="flex-1 lg:flex-none"
          >
            <Download className="w-4 h-4 mr-2" />
            Export
          </Button>
          <Button 
            variant="secondary" 
            onClick={handleExportAll}
            disabled={feeders.length === 0}
            className="flex-1 lg:flex-none"
          >
            <Download className="w-4 h-4 mr-2" />
            Export All
          </Button>
        </div>
      </div>

      {/* Feeder Selection Dropdown */}
      <div className="w-full max-w-xl mb-6">
        <label className="text-sm font-medium mb-2 block text-slate-600 dark:text-slate-400">
          Select Feeder
        </label>
        <div className="flex items-center gap-2">
          <Button
            variant="outline"
            size="icon"
            onClick={goToPrevFeeder}
            disabled={!selectedFeeder}
            title="Previous Feeder"
          >
            <ChevronLeft className="w-4 h-4" />
          </Button>
          <Select
            value={selectedFeeder?.id || ''}
            onValueChange={(value) => {
              const feeder = feeders.find(f => f.id === value);
              if (feeder) handleFeederChange(feeder);
            }}
          >
            <SelectTrigger className="w-full">
              <SelectValue placeholder="Select a feeder" />
            </SelectTrigger>
            <SelectContent className="max-h-[300px]">
              {getSortedFeeders().map(feeder => (
                <SelectItem key={feeder.id} value={feeder.id}>
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
          >
            <ChevronRight className="w-4 h-4" />
          </Button>
        </div>
      </div>

      {/* Data Table */}
      {selectedFeeder && (
        <FeederTable
          feeder={selectedFeeder}
          entries={entries}
          loading={loading}
          onUpdate={handleEntryUpdated}
          onDelete={handleEntryDeleted}
        />
      )}

      {/* Analytics */}
      {selectedFeeder && entries.length > 0 && (
        <AnalyticsCharts entries={entries} feeder={selectedFeeder} />
      )}

      {/* Data Entry Modal */}
      {isModalOpen && selectedFeeder && (
        <DataEntryModal
          isOpen={isModalOpen}
          onClose={() => setIsModalOpen(false)}
          feeder={selectedFeeder}
          year={year}
          month={month}
          defaultDate={(() => {
            if (entries.length === 0) {
              return `${year}-${String(month).padStart(2, '0')}-01`;
            }
            const sorted = [...entries].sort((a, b) => a.date.localeCompare(b.date));
            const lastDate = new Date(sorted[sorted.length - 1].date);
            lastDate.setDate(lastDate.getDate() + 1);
            return lastDate.toISOString().split('T')[0];
          })()}
          onEntryCreated={handleEntryCreated}
          onEntryUpdated={handleEntryUpdated}
          entries={entries}
          onPrevFeeder={goToPrevFeeder}
          onNextFeeder={goToNextFeeder}
        />
      )}
      
      {/* Import Preview Modal */}
      {importPreviewOpen && (
        <ImportPreviewModal
          isOpen={importPreviewOpen}
          onClose={() => setImportPreviewOpen(false)}
          data={importData}
          onConfirm={handleImportConfirm}
          loading={loading}
        />
      )}
    </div>
  );
}
