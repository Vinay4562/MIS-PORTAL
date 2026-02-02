import { useState, useEffect } from 'react';
import axios from 'axios';
import { Dialog, DialogContent, DialogHeader, DialogTitle } from '@/components/ui/dialog';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { toast } from 'sonner';
import { Edit2 } from 'lucide-react';

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

export default function DataEntryModal({ isOpen, onClose, feeder, year, month, onEntryCreated }) {
  const [selectedDate, setSelectedDate] = useState('');
  const [end1ImportFinal, setEnd1ImportFinal] = useState('');
  const [end1ExportFinal, setEnd1ExportFinal] = useState('');
  const [end2ImportFinal, setEnd2ImportFinal] = useState('');
  const [end2ExportFinal, setEnd2ExportFinal] = useState('');
  const [previousEntry, setPreviousEntry] = useState(null);
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    if (isOpen) {
      const today = new Date();
      const dateStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
      setSelectedDate(dateStr);
      fetchPreviousEntry(dateStr);
    }
  }, [isOpen]);

  useEffect(() => {
    if (selectedDate) {
      fetchPreviousEntry(selectedDate);
    }
  }, [selectedDate]);

  const fetchPreviousEntry = async (date) => {
    try {
      const dateObj = new Date(date);
      const prevDate = new Date(dateObj);
      prevDate.setDate(prevDate.getDate() - 1);
      const prevDateStr = prevDate.toISOString().split('T')[0];

      const response = await axios.get(`${API}/entries`, {
        params: { feeder_id: feeder.id }
      });

      const prevEntry = response.data.find(e => e.date === prevDateStr);
      setPreviousEntry(prevEntry || null);
    } catch (error) {
      console.error('Failed to fetch previous entry:', error);
    }
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setLoading(true);

    try {
      const response = await axios.post(`${API}/entries`, {
        feeder_id: feeder.id,
        date: selectedDate,
        end1_import_final: parseFloat(end1ImportFinal),
        end1_export_final: parseFloat(end1ExportFinal),
        end2_import_final: parseFloat(end2ImportFinal),
        end2_export_final: parseFloat(end2ExportFinal)
      });

      onEntryCreated(response.data);
      onClose();
    } catch (error) {
      toast.error(error.response?.data?.detail || 'Failed to create entry');
    } finally {
      setLoading(false);
    }
  };

  const getInitialValue = (field) => {
    if (!previousEntry) return '0';
    return previousEntry[field]?.toString() || '0';
  };

  return (
    <Dialog open={isOpen} onOpenChange={onClose}>
      <DialogContent className="max-w-2xl max-h-[90vh] overflow-y-auto" data-testid="data-entry-modal" aria-describedby="entry-form-description">
        <DialogHeader>
          <DialogTitle className="text-2xl font-heading">New Entry - {feeder.name}</DialogTitle>
          <p id="entry-form-description" className="sr-only">
            Enter daily readings for {feeder.name}. Initial values are auto-filled from previous day.
          </p>
        </DialogHeader>

        <form onSubmit={handleSubmit} className="space-y-6">
          <div className="space-y-2">
            <Label htmlFor="date">Date</Label>
            <Input
              id="date"
              type="date"
              value={selectedDate}
              onChange={(e) => setSelectedDate(e.target.value)}
              required
              data-testid="date-input"
            />
          </div>

          {/* End 1 */}
          <div className="bg-slate-50 dark:bg-slate-800 p-4 rounded-lg space-y-4">
            <h3 className="font-semibold text-lg">{feeder.end1_name} End</h3>
            
            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-2">
                <Label>Import Initial (MWH)</Label>
                <div className="flex items-center gap-2">
                  <Input
                    type="number"
                    step="0.01"
                    value={getInitialValue('end1_import_final')}
                    disabled
                    className="bg-slate-100 dark:bg-slate-700"
                  />
                  <span className="text-xs text-slate-500">Auto-filled</span>
                </div>
              </div>

              <div className="space-y-2">
                <Label htmlFor="end1ImportFinal">Import Final (MWH)</Label>
                <Input
                  id="end1ImportFinal"
                  type="number"
                  step="0.01"
                  value={end1ImportFinal}
                  onChange={(e) => setEnd1ImportFinal(e.target.value)}
                  required
                  data-testid="end1-import-final"
                />
              </div>

              <div className="space-y-2">
                <Label>Export Initial (MWH)</Label>
                <div className="flex items-center gap-2">
                  <Input
                    type="number"
                    step="0.01"
                    value={getInitialValue('end1_export_final')}
                    disabled
                    className="bg-slate-100 dark:bg-slate-700"
                  />
                  <span className="text-xs text-slate-500">Auto-filled</span>
                </div>
              </div>

              <div className="space-y-2">
                <Label htmlFor="end1ExportFinal">Export Final (MWH)</Label>
                <Input
                  id="end1ExportFinal"
                  type="number"
                  step="0.01"
                  value={end1ExportFinal}
                  onChange={(e) => setEnd1ExportFinal(e.target.value)}
                  required
                  data-testid="end1-export-final"
                />
              </div>
            </div>
            <div className="text-xs text-slate-500">
              MF: Import = {feeder.end1_import_mf}, Export = {feeder.end1_export_mf}
            </div>
          </div>

          {/* End 2 */}
          <div className="bg-slate-50 dark:bg-slate-800 p-4 rounded-lg space-y-4">
            <h3 className="font-semibold text-lg">{feeder.end2_name} End</h3>
            
            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-2">
                <Label>Import Initial (MWH)</Label>
                <div className="flex items-center gap-2">
                  <Input
                    type="number"
                    step="0.01"
                    value={getInitialValue('end2_import_final')}
                    disabled
                    className="bg-slate-100 dark:bg-slate-700"
                  />
                  <span className="text-xs text-slate-500">Auto-filled</span>
                </div>
              </div>

              <div className="space-y-2">
                <Label htmlFor="end2ImportFinal">Import Final (MWH)</Label>
                <Input
                  id="end2ImportFinal"
                  type="number"
                  step="0.01"
                  value={end2ImportFinal}
                  onChange={(e) => setEnd2ImportFinal(e.target.value)}
                  required
                  data-testid="end2-import-final"
                />
              </div>

              <div className="space-y-2">
                <Label>Export Initial (MWH)</Label>
                <div className="flex items-center gap-2">
                  <Input
                    type="number"
                    step="0.01"
                    value={getInitialValue('end2_export_final')}
                    disabled
                    className="bg-slate-100 dark:bg-slate-700"
                  />
                  <span className="text-xs text-slate-500">Auto-filled</span>
                </div>
              </div>

              <div className="space-y-2">
                <Label htmlFor="end2ExportFinal">Export Final (MWH)</Label>
                <Input
                  id="end2ExportFinal"
                  type="number"
                  step="0.01"
                  value={end2ExportFinal}
                  onChange={(e) => setEnd2ExportFinal(e.target.value)}
                  required
                  data-testid="end2-export-final"
                />
              </div>
            </div>
            <div className="text-xs text-slate-500">
              MF: Import = {feeder.end2_import_mf}, Export = {feeder.end2_export_mf}
            </div>
          </div>

          <div className="flex gap-2 justify-end">
            <Button type="button" variant="outline" onClick={onClose} data-testid="cancel-button">
              Cancel
            </Button>
            <Button type="submit" disabled={loading} data-testid="save-entry-button">
              {loading ? 'Saving...' : 'Save Entry'}
            </Button>
          </div>
        </form>
      </DialogContent>
    </Dialog>
  );
}
