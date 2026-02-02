import { useState, useEffect } from 'react';
import axios from 'axios';
import { toast } from 'sonner';
import { Button } from '@/components/ui/button';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import DataEntryModal from '@/components/DataEntryModal';
import FeederTable from '@/components/FeederTable';
import AnalyticsCharts from '@/components/AnalyticsCharts';
import { Download, Plus, Calendar } from 'lucide-react';

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

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

  const handleSubmitDateSelection = () => {
    if (feeders.length > 0) {
      setSelectedFeeder(feeders[0]);
      setShowDateSelector(false);
      fetchEntries(feeders[0].id, year, month);
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
    toast.success('Entry created successfully');
    setIsModalOpen(false);
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
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div>
          <h1 className="text-3xl font-heading font-bold text-slate-900 dark:text-slate-100">
            Line Losses Management
          </h1>
          <p className="text-slate-600 dark:text-slate-400 mt-1">
            {monthNames[month - 1]} {year}
          </p>
        </div>
        
        <div className="flex gap-2">
          <Button 
            variant="outline" 
            onClick={() => setShowDateSelector(true)}
            data-testid="change-period-button"
          >
            <Calendar className="w-4 h-4 mr-2" />
            Change Period
          </Button>
          <Button 
            onClick={() => setIsModalOpen(true)}
            data-testid="new-entry-button"
          >
            <Plus className="w-4 h-4 mr-2" />
            New Entry
          </Button>
          <Button 
            variant="secondary" 
            onClick={handleExport}
            disabled={!selectedFeeder || entries.length === 0}
            data-testid="export-button"
          >
            <Download className="w-4 h-4 mr-2" />
            Export
          </Button>
        </div>
      </div>

      {/* Feeder Tabs */}
      <div className="feeder-tabs">
        {feeders.map(feeder => (
          <button
            key={feeder.id}
            onClick={() => handleFeederChange(feeder)}
            className={`feeder-tab ${selectedFeeder?.id === feeder.id ? 'active' : ''}`}
            data-testid={`feeder-tab-${feeder.id}`}
          >
            {feeder.name}
          </button>
        ))}
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
          onEntryCreated={handleEntryCreated}
        />
      )}
    </div>
  );
}
