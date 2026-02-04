import { Dialog, DialogContent, DialogHeader, DialogTitle } from '@/components/ui/dialog';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { useState, useEffect, useCallback, useRef } from 'react';
import axios from 'axios';
import { toast } from 'sonner';
import { ChevronLeft, ChevronRight } from 'lucide-react';

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

export default function EnergyEntryModal({ isOpen, onClose, sheet, year, month, defaultDate, onEntryCreated, entry, entries, onPrevSheet, onNextSheet }) {
  const [date, setDate] = useState(() => {
      // Use defaultDate if provided (calculated by parent based on existing entries)
      if (defaultDate) return defaultDate;

      // Fallback logic
      const today = new Date();
      if (today.getFullYear() === year && (today.getMonth() + 1) === month) {
          return `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
      }
      return `${year}-${String(month).padStart(2, '0')}-01`;
  });
  const [readings, setReadings] = useState({}); // { meterId: finalValue }
  const [previousEntry, setPreviousEntry] = useState(null);
  const [loading, setLoading] = useState(false);
  const [currentEntry, setCurrentEntry] = useState(entry);
  const firstInputRef = useRef(null);

  // Calculate constraints
  const strMonth = month.toString().padStart(2, '0');
  const minDate = `${year}-${strMonth}-01`;
  const maxDay = new Date(year, month, 0).getDate();
  const maxDate = `${year}-${strMonth}-${maxDay}`;

  useEffect(() => {
    if (isOpen) {
        if (entry) {
            setDate(entry.date);
            setCurrentEntry(entry);
            const entryReadings = {};
            entry.readings.forEach(r => {
                entryReadings[r.meter_id] = r.final;
            });
            setReadings(entryReadings);
        } else {
            // New entry logic
            if (defaultDate) {
                setDate(defaultDate);
            } else {
                const today = new Date();
                if (today.getFullYear() === year && (today.getMonth() + 1) === month) {
                    setDate(`${year}-${String(month).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`);
                } else {
                    setDate(`${year}-${String(month).padStart(2, '0')}-01`);
                }
            }
            setCurrentEntry(null);
            setReadings({});
        }
        setPreviousEntry(null);
    }
  }, [isOpen, entry, sheet.id, year, month]);
  
  // Handle date navigation and data loading
  useEffect(() => {
      if (isOpen && entries && date) {
          // Don't overwrite if date matches currentEntry (initial load)
          if (currentEntry && currentEntry.date === date) return;

          const found = entries.find(e => e.date === date);
          if (found) {
              setCurrentEntry(found);
              const entryReadings = {};
              found.readings.forEach(r => {
                  entryReadings[r.meter_id] = r.final;
              });
              setReadings(entryReadings);
          } else {
              setCurrentEntry(null);
              setReadings({});
          }
      }
  }, [date, entries, isOpen]); // Removed currentEntry from dependency to avoid loop

  // Focus effect on sheet change
  useEffect(() => {
      if (!currentEntry) {
        setTimeout(() => firstInputRef.current?.focus(), 0);
      }
  }, [sheet.id, currentEntry]);

  const fetchPreviousEntry = useCallback(async (currentDate) => {
    try {
        const dateObj = new Date(currentDate);
        const prevDate = new Date(dateObj);
        prevDate.setDate(prevDate.getDate() - 1);
        const prevDateStr = prevDate.toISOString().split('T')[0];

        // Fetch entries to find previous one. 
        // We fetch recent entries for the sheet.
        const response = await axios.get(`${API}/energy/entries/${sheet.id}`);
        const prevEntry = response.data.find(e => e.date === prevDateStr);
        setPreviousEntry(prevEntry || null);
    } catch (error) {
        console.error('Failed to fetch previous entry:', error);
    }
  }, [sheet]);

  useEffect(() => {
    if (date && sheet) {
        fetchPreviousEntry(date);
    }
  }, [date, sheet, fetchPreviousEntry]);

  const handlePrevDate = () => {
    const d = new Date(date);
    d.setDate(d.getDate() - 1);
    const newDate = d.toISOString().split('T')[0];
    if (newDate >= minDate) {
        setDate(newDate);
    } else {
        toast.error("Date out of range");
    }
  };

  const handleNextDate = () => {
    const d = new Date(date);
    d.setDate(d.getDate() + 1);
    const newDate = d.toISOString().split('T')[0];
    if (newDate <= maxDate) {
        setDate(newDate);
    } else {
        toast.error("Date out of range");
    }
  };

  const getInitialValue = (meterId) => {
    // If editing, use the actual initial value from the entry itself if available, 
    // OR fetch from previous entry. 
    // Actually, consistency matters: The current entry's initial SHOULD match previous final.
    // So logic remains same: derive from previous entry.
    // EXCEPT: What if previous entry is missing? Then 0.
    // BUT: If we are editing, we might want to see what was actually saved as initial?
    // The UI shows "Initial (Prev. Final)". So showing previous final is correct.
    
    if (!previousEntry) return '0';
    const reading = previousEntry.readings.find(r => r.meter_id === meterId);
    return reading ? reading.final : '0';
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setLoading(true);
    
    try {
        const payload = {
            sheet_id: sheet.id,
            date: date,
            readings: Object.entries(readings).map(([meterId, final]) => ({
                meter_id: meterId,
                final: parseFloat(final)
            }))
        };

        let response;
        if (currentEntry) {
            response = await axios.put(`${API}/energy/entries/${currentEntry.id}`, payload);
        } else {
            response = await axios.post(`${API}/energy/entries`, payload);
        }
        
        onEntryCreated(response.data);
    } catch (error) {
        toast.error(currentEntry ? 'Failed to update entry' : 'Failed to save entry');
        console.error(error);
    } finally {
        setLoading(false);
    }
  };

  return (
    <Dialog open={isOpen} onOpenChange={onClose}>
      <DialogContent className="max-w-2xl max-h-[90vh] overflow-y-auto">
        <DialogHeader>
          <div className="flex items-center justify-between">
            <DialogTitle>{currentEntry ? 'Edit Entry' : 'New Entry'} - {sheet.name}</DialogTitle>
            <div className="flex gap-2">
              {onPrevSheet && (
                <Button type="button" variant="outline" onClick={onPrevSheet}>Previous Sheet</Button>
              )}
              {onNextSheet && (
                <Button type="button" variant="outline" onClick={onNextSheet}>Next Sheet</Button>
              )}
            </div>
          </div>
        </DialogHeader>
        <form onSubmit={handleSubmit} className="space-y-6">
            <div className="space-y-2">
                <Label>Date</Label>
                <div className="flex items-center gap-2">
                    <Button type="button" variant="outline" size="icon" onClick={handlePrevDate} title="Previous Date">
                        <ChevronLeft className="h-4 w-4" />
                    </Button>
                    <Input 
                        type="date" 
                        value={date} 
                        min={minDate}
                        max={maxDate}
                        onChange={e => setDate(e.target.value)} 
                        required 
                        className="flex-1"
                    />
                    <Button type="button" variant="outline" size="icon" onClick={handleNextDate} title="Next Date">
                        <ChevronRight className="h-4 w-4" />
                    </Button>
                </div>
            </div>
            
            <div className="space-y-4">
                <div className="flex items-center justify-between border-b pb-2">
                    <h3 className="font-medium text-sm text-slate-500">Meter Readings</h3>
                    <span className="text-xs text-slate-400">Previous day's final becomes today's initial</span>
                </div>
                
                <div className="grid grid-cols-12 gap-4 font-medium text-sm text-slate-500 mb-2 px-1">
                    <div className="col-span-4">Meter Name</div>
                    <div className="col-span-4">Initial (Prev. Final)</div>
                    <div className="col-span-4">Final Reading</div>
                </div>

                <div className="space-y-4">
                    {sheet.meters.map((meter, index) => (
                        <div key={meter.id} className="grid grid-cols-12 gap-4 items-center bg-slate-50 dark:bg-slate-800/50 p-3 rounded-lg">
                            <div className="col-span-4">
                                <Label className="truncate block font-medium" title={meter.name}>
                                    {meter.name}
                                </Label>
                                <span className="text-xs text-slate-400">{meter.unit} (MF: {meter.mf})</span>
                            </div>
                            
                            <div className="col-span-4">
                                <Input 
                                    type="number" 
                                    value={getInitialValue(meter.id)}
                                    disabled
                                    className="bg-slate-100 dark:bg-slate-700 text-slate-500"
                                />
                            </div>

                            <div className="col-span-4">
                                <Input 
                                    ref={index === 0 ? firstInputRef : null}
                                    type="number" 
                                    step="0.01" 
                                    placeholder="Final"
                                    value={readings[meter.id] || ''}
                                    onChange={e => setReadings(prev => ({...prev, [meter.id]: e.target.value}))}
                                    required
                                    className="font-medium"
                                />
                            </div>
                        </div>
                    ))}
                </div>
            </div>

            <div className="pt-4 flex justify-end gap-2 border-t">
                <Button type="button" variant="outline" onClick={onClose}>Cancel</Button>
                <Button type="submit" disabled={loading}>Save Entry</Button>
            </div>
        </form>
      </DialogContent>
    </Dialog>
  );
}
