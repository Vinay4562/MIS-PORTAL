import { useState, useEffect } from 'react';
import axios from 'axios';
import { toast } from 'sonner';
import { Button } from '@/components/ui/button';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Download, Plus, Activity, Calendar, RefreshCcw } from 'lucide-react';
import EnergyTable from '@/components/EnergyTable';
import EnergyEntryModal from '@/components/EnergyEntryModal';
import EnergyAnalytics from '@/components/EnergyAnalytics';

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

export default function EnergyConsumption() {
  const [sheets, setSheets] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState(null);
  const [entries, setEntries] = useState([]);
  const [year, setYear] = useState(new Date().getFullYear());
  const [month, setMonth] = useState(new Date().getMonth() + 1);
  const [showDateSelector, setShowDateSelector] = useState(true);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [loading, setLoading] = useState(false);
  const [initialized, setInitialized] = useState(false);
  const [editingEntry, setEditingEntry] = useState(null);

  useEffect(() => {
    initializeModule();
  }, []);

  const initializeModule = async () => {
    try {
      // First check/init
      await axios.post(`${API}/energy/init`);
      // Then fetch sheets
      const response = await axios.get(`${API}/energy/sheets`);
      setSheets(response.data);
      setInitialized(true);
    } catch (error) {
      console.error('Failed to initialize energy module:', error);
      toast.error('Failed to load energy sheets');
    }
  };

  const handleSubmitDateSelection = () => {
    if (sheets.length > 0) {
      // Default to first sheet if none selected, or keep existing
      const sheetToUse = selectedSheet || sheets[0];
      setSelectedSheet(sheetToUse);
      setShowDateSelector(false);
      fetchEntries(sheetToUse.id, year, month);
    } else {
        toast.error("No sheets available");
    }
  };

  const fetchEntries = async (sheetId, selectedYear, selectedMonth) => {
    setLoading(true);
    try {
      const response = await axios.get(`${API}/energy/entries/${sheetId}`, {
        params: { year: selectedYear, month: selectedMonth }
      });
      setEntries(response.data);
    } catch (error) {
      console.error('Failed to fetch entries:', error);
      toast.error('Failed to load entries');
    } finally {
      setLoading(false);
    }
  };

  const handleSheetChange = (sheet) => {
    setSelectedSheet(sheet);
    fetchEntries(sheet.id, year, month);
  };

  const goToPrevSheet = () => {
    const idx = sheets.findIndex(s => s.id === selectedSheet?.id);
    if (idx > 0) {
      handleSheetChange(sheets[idx - 1]);
    }
  };

  const goToNextSheet = () => {
    const idx = sheets.findIndex(s => s.id === selectedSheet?.id);
    if (idx !== -1 && idx < sheets.length - 1) {
      handleSheetChange(sheets[idx + 1]);
    }
  };

  const handleExport = async () => {
    if (!selectedSheet) return;
    try {
      const response = await axios.get(
        `${API}/energy/export/${selectedSheet.id}/${year}/${month}`,
        { responseType: 'blob' }
      );
      
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `${selectedSheet.name}_${year}_${month.toString().padStart(2, '0')}.xlsx`);
      document.body.appendChild(link);
      link.click();
      link.remove();
    } catch (error) {
      toast.error('Failed to export data');
    }
  };
  
  const handleRefresh = () => {
    if (!selectedSheet) return;
    fetchEntries(selectedSheet.id, year, month);
    toast.success('Data refreshed');
  };

  const handleEntryCreated = (newEntry) => {
    fetchEntries(selectedSheet.id, year, month);
    toast.success('Data Saved');
    setEditingEntry(null);
    
  };

  const handleEdit = (entry) => {
    setEditingEntry(entry);
    setIsModalOpen(true);
  };

  const handleDelete = async (entry) => {
    if (!window.confirm('Are you sure you want to delete this entry?')) return;
    
    try {
        await axios.delete(`${API}/energy/entries/${entry.id}`);
        toast.success('Entry deleted successfully');
        fetchEntries(selectedSheet.id, year, month);
    } catch (error) {
        console.error('Failed to delete entry:', error);
        toast.error('Failed to delete entry');
    }
  };

  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];

  if (!initialized) {
    return <div className="p-8 text-center">Loading Energy Module...</div>;
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

            <Button 
              onClick={handleSubmitDateSelection} 
              className="w-full" 
              size="lg"
            >
              Load Data
            </Button>
          </CardContent>
        </Card>
      </div>
    );
  }

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4">
        <div>
          <h1 className="text-2xl md:text-3xl font-heading font-bold text-slate-900 dark:text-slate-100 flex items-center gap-2">
            <Activity className="w-6 h-6 md:w-8 md:h-8 text-blue-600" />
            Energy Consumption
          </h1>
          <p className="text-sm md:text-base text-slate-500 dark:text-slate-400 mt-1">
            {monthNames[month - 1]} {year}
          </p>
        </div>
        
        <div className="flex flex-wrap items-center gap-2 md:gap-3 w-full lg:w-auto">
            <Button 
                variant="outline" 
                onClick={handleRefresh}
                disabled={!selectedSheet}
                className="flex-1 lg:flex-none"
                data-testid="refresh-button"
            >
                <RefreshCcw className="w-4 h-4 mr-2" />
                Refresh
            </Button>
            <Button 
                variant="outline" 
                onClick={() => setShowDateSelector(true)}
                className="flex-1 lg:flex-none"
            >
                <Calendar className="w-4 h-4 mr-2" />
                Period
            </Button>
            <Button 
                onClick={() => setIsModalOpen(true)} 
                disabled={!selectedSheet}
                className="flex-1 lg:flex-none"
            >
              <Plus className="w-4 h-4 mr-2" />
              Entry
            </Button>
            <Button 
                variant="secondary" 
                onClick={handleExport} 
                disabled={!selectedSheet || entries.length === 0}
                className="flex-1 lg:flex-none"
            >
              <Download className="w-4 h-4 mr-2" />
              Export
            </Button>
        </div>
      </div>

      {/* Sheet Selection */}
      <div className="w-full max-w-xl mb-6">
        <label className="text-sm font-medium mb-2 block text-slate-600 dark:text-slate-400">
          Select Sheet
        </label>
        <Select
          value={selectedSheet?.id || ''}
          onValueChange={(value) => {
              const sheet = sheets.find(s => s.id === value);
              if (sheet) handleSheetChange(sheet);
          }}
        >
          <SelectTrigger>
            <SelectValue placeholder="Select sheet" />
          </SelectTrigger>
          <SelectContent>
            {sheets.map(sheet => (
              <SelectItem key={sheet.id} value={sheet.id}>
                {sheet.name}
              </SelectItem>
            ))}
          </SelectContent>
        </Select>
      </div>

      {/* Data Table */}
      {selectedSheet && (
        <EnergyTable 
            sheet={selectedSheet} 
            entries={entries} 
            loading={loading}
            onEdit={handleEdit}
            onDelete={handleDelete}
        />
      )}

      {/* Data Entry Modal */}
      {isModalOpen && selectedSheet && (
        <EnergyEntryModal
          isOpen={isModalOpen}
          onClose={() => {
            setIsModalOpen(false);
            setEditingEntry(null);
          }}
          sheet={selectedSheet}
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
          entry={editingEntry}
          onPrevSheet={goToPrevSheet}
          onNextSheet={goToNextSheet}
        />
      )}
      
      {selectedSheet && entries.length > 0 && (
        <EnergyAnalytics entries={entries} sheet={selectedSheet} />
      )}
    </div>
  );
}
