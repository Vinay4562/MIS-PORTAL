import { useState, useEffect } from 'react';
import axios from 'axios';
import { toast } from 'sonner';
import { Button } from '@/components/ui/button';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Download, Plus, Activity, Calendar } from 'lucide-react';
import EnergyTable from '@/components/EnergyTable';
import EnergyEntryModal from '@/components/EnergyEntryModal';

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

  const handleEntryCreated = (newEntry) => {
    fetchEntries(selectedSheet.id, year, month);
    toast.success('Entry saved successfully');
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
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h1 className="text-3xl font-heading font-bold text-slate-900 dark:text-slate-100 flex items-center gap-2">
            <Activity className="w-8 h-8 text-blue-600" />
            Energy Consumption
          </h1>
          <p className="text-slate-500 dark:text-slate-400 mt-1">
            {monthNames[month - 1]} {year}
          </p>
        </div>
        
        <div className="flex items-center gap-3">
            <Button 
                variant="outline" 
                onClick={() => setShowDateSelector(true)}
            >
                <Calendar className="w-4 h-4 mr-2" />
                Change Period
            </Button>
            <Button onClick={() => setIsModalOpen(true)} disabled={!selectedSheet}>
              <Plus className="w-4 h-4 mr-2" />
              New Entry
            </Button>
            <Button variant="secondary" onClick={handleExport} disabled={!selectedSheet || entries.length === 0}>
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

      {selectedSheet && (
        <EnergyEntryModal
            isOpen={isModalOpen}
            onClose={() => {
                setIsModalOpen(false);
                setEditingEntry(null);
            }}
            sheet={selectedSheet}
            onEntryCreated={handleEntryCreated}
            entry={editingEntry}
        />
      )}
    </div>
  );
}
