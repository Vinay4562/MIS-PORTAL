import React, { useState, useEffect, useRef } from 'react';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { toast } from 'sonner';
import { Edit, Calendar, Ban, ChevronLeft, ChevronRight } from 'lucide-react';

export default function MaxMinEntryModal({ isOpen, onClose, onSave, feeder, year, month, initialData, defaultDate, onPrevFeeder, onNextFeeder, entries }) {
  const [formData, setFormData] = useState({});
  const [selectedDate, setSelectedDate] = useState('');
  const [loading, setLoading] = useState(false);
  const firstInputRef = useRef(null);

  // Calculate min and max date for the date picker
  const strMonth = month.toString().padStart(2, '0');
  const minDate = `${year}-${strMonth}-01`;
  const maxDay = new Date(year, month, 0).getDate();
  const maxDate = `${year}-${strMonth}-${maxDay}`;

  useEffect(() => {
    if (isOpen) {
      if (initialData) {
        setFormData(initialData.data || {});
        setSelectedDate(initialData.date);
      } else {
        setFormData({});
        // Use defaultDate if provided
        if (defaultDate) {
            setSelectedDate(defaultDate);
        }
      }
      
      // Focus first input
      setTimeout(() => firstInputRef.current?.focus(), 0);
    }
  }, [isOpen, initialData, defaultDate, feeder?.id]);

  // Load data when date changes
  useEffect(() => {
    if (!isOpen || !selectedDate || !entries) return;

    // Avoid overwriting if we are on the initial edit entry
    if (initialData && initialData.date === selectedDate) {
        return;
    }

    const found = entries.find(e => e.date === selectedDate);
    if (found) {
        setFormData(found.data || {});
    } else {
        // Clear form for new date
        setFormData({});
    }
  }, [selectedDate, entries, isOpen, initialData]);

  const handleChange = (path, value) => {
    setFormData(prev => {
      const newData = { ...prev };
      const keys = path.split('.');
      let current = newData;
      for (let i = 0; i < keys.length - 1; i++) {
        if (!current[keys[i]]) current[keys[i]] = {};
        current = current[keys[i]];
      }
      current[keys[keys.length - 1]] = value;
      return newData;
    });
  };

  const getValue = (path) => {
    const keys = path.split('.');
    let current = formData;
    for (const key of keys) {
      if (!current) return '';
      current = current[key];
    }
    return current || '';
  };

  const handlePrevDate = () => {
    const d = new Date(selectedDate);
    d.setDate(d.getDate() - 1);
    const newDate = d.toISOString().split('T')[0];
    if (newDate >= minDate) {
        setSelectedDate(newDate);
    } else {
        toast.error("Date out of range");
    }
  };

  const handleNextDate = () => {
    const d = new Date(selectedDate);
    d.setDate(d.getDate() + 1);
    const newDate = d.toISOString().split('T')[0];
    if (newDate <= maxDate) {
        setSelectedDate(newDate);
    } else {
        toast.error("Date out of range");
    }
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!selectedDate) {
      toast.error('Please select a date');
      return;
    }
    setLoading(true);
    try {
      await onSave({
        date: selectedDate,
        data: formData
      });
    } catch (error) {
      console.error(error);
      toast.error('Failed to save entry');
    } finally {
      setLoading(false);
    }
  };

  if (!feeder) return null;

  const handleSharedTimeChange = (type, value) => {
    if (type === 'max') {
      handleChange('max_bus_voltage_400kv.time', value);
      handleChange('max_bus_voltage_220kv.time', value);
    } else {
      handleChange('min_bus_voltage_400kv.time', value);
      handleChange('min_bus_voltage_220kv.time', value);
    }
  };

  const handleNotInService = () => {
    setFormData(prev => {
      const newData = { ...prev };
      
      const setVal = (path, val) => {
        const keys = path.split('.');
        let current = newData;
        for (let i = 0; i < keys.length - 1; i++) {
          if (!current[keys[i]]) current[keys[i]] = {};
          current = current[keys[i]];
        }
        current[keys[keys.length - 1]] = val;
      };

      const fields = [
          'max.amps', 'max.mw', 'max.time',
          'min.amps', 'min.mw', 'min.time',
          'avg.amps', 'avg.mw'
      ];
      
      if (feeder.type === 'ict_feeder') {
          fields.push('max.mvar', 'min.mvar');
      }

      fields.forEach(f => setVal(f, 'N/S'));
      
      return newData;
    });
  };

  const renderField = (label, path, type = "number", step = "0.01") => {
    const val = getValue(path);
    const isNS = val === 'N/S';
    return (
    <div className="grid grid-cols-4 items-center gap-4">
      <Label htmlFor={path} className="text-right text-slate-600 font-medium">
        {label}
      </Label>
      <Input
        id={path}
        type={isNS ? "text" : type}
        step={step}
        value={val}
        onChange={(e) => handleChange(path, e.target.value)}
        className={`col-span-3 border-slate-200 focus:ring-2 focus:ring-indigo-500/20 ${isNS ? 'bg-slate-50 text-slate-500' : ''}`}
        disabled={isNS}
        required
      />
    </div>
    );
  };

  const renderTimeField = (label, path) => {
    const val = getValue(path);
    const isNS = val === 'N/S';
    return (
    <div className="grid grid-cols-4 items-center gap-4">
      <Label htmlFor={path} className="text-right text-slate-600 font-medium">
        {label}
      </Label>
      <Input
        id={path}
        type={isNS ? "text" : "time"}
        value={val}
        onChange={(e) => handleChange(path, e.target.value)}
        className={`col-span-3 border-slate-200 focus:ring-2 focus:ring-indigo-500/20 ${isNS ? 'bg-slate-50 text-slate-500' : ''}`}
        disabled={isNS}
        required
      />
    </div>
    );
  };
 
  const getGroupTime = (kind) => {
    const explicit = getValue(`${kind}_bus_voltage.time`);
    if (explicit) return explicit;
    return (
      getValue(`${kind}_bus_voltage_400kv.time`) ||
      getValue(`${kind}_bus_voltage_220kv.time`) ||
      ''
    );
  };
 
  const setGroupTime = (kind, value) => {
    handleChange(`${kind}_bus_voltage.time`, value);
    handleChange(`${kind}_bus_voltage_400kv.time`, value);
    handleChange(`${kind}_bus_voltage_220kv.time`, value);
  };

  return (
    <Dialog open={isOpen} onOpenChange={onClose}>
      <DialogContent className="max-w-2xl max-h-[90vh] p-0 gap-0 border-0 shadow-2xl overflow-hidden flex flex-col">
        <DialogHeader className="bg-gradient-to-r from-indigo-600 to-purple-600 p-6 text-white flex-shrink-0">
          <DialogTitle className="flex items-center justify-between gap-2 text-xl">
            <div className="flex items-center gap-2">
                <div className="p-2 bg-white/20 rounded-lg backdrop-blur-sm">
                    {initialData ? <Edit className="w-5 h-5 text-white" /> : <Calendar className="w-5 h-5 text-white" />}
                </div>
                <span className="flex items-center gap-2">
                  {initialData ? 'Edit Entry' : 'New Entry'} - {feeder.name}
                </span>
            </div>
            <div className="flex items-center gap-2">
              {onPrevFeeder && (
                <Button type="button" variant="secondary" onClick={onPrevFeeder} className="bg-white/20 text-white hover:bg-white/30 text-xs h-8">
                  Prev Feeder
                </Button>
              )}
              {onNextFeeder && (
                <Button type="button" variant="secondary" onClick={onNextFeeder} className="bg-white/20 text-white hover:bg-white/30 text-xs h-8">
                  Next Feeder
                </Button>
              )}
            </div>
          </DialogTitle>
        </DialogHeader>

        <form onSubmit={handleSubmit} className="flex flex-col flex-1 overflow-hidden">
          <div className="flex-1 overflow-y-auto p-6 space-y-6">
            
            <div className="space-y-4">
                {/* Date Selection with Navigation */}
                <div className="grid grid-cols-4 items-center gap-4">
                    <Label htmlFor="date" className="text-right text-slate-600 font-medium">Date</Label>
                    <div className="col-span-3 flex items-center gap-2">
                        <Button type="button" variant="outline" size="icon" onClick={handlePrevDate} title="Previous Date">
                            <ChevronLeft className="h-4 w-4" />
                        </Button>
                        <Input
                          id="date"
                          type="date"
                          value={selectedDate}
                          min={minDate}
                          max={maxDate}
                          onChange={(e) => setSelectedDate(e.target.value)}
                          required
                          className="flex-1 border-slate-200 focus:ring-2 focus:ring-indigo-500/20"
                        />
                        <Button type="button" variant="outline" size="icon" onClick={handleNextDate} title="Next Date">
                            <ChevronRight className="h-4 w-4" />
                        </Button>
                    </div>
                </div>

                {feeder.type !== 'bus_station' && (
                    <div className="flex justify-end">
                        <Button 
                            type="button" 
                            variant="destructive" 
                            size="sm" 
                            onClick={handleNotInService}
                            className="bg-rose-100 text-rose-700 hover:bg-rose-200 border-rose-200"
                        >
                            <Ban className="w-4 h-4 mr-2" />
                            Mark as Not In Service
                        </Button>
                    </div>
                )}
            </div>

            <div className="space-y-6">
              {feeder.type === 'bus_station' ? (
                <>
                  <div className="bg-slate-50 p-4 rounded-xl border border-slate-100 space-y-4">
                    <h3 className="font-semibold text-slate-800 flex items-center gap-2">
                        <span className="w-1.5 h-1.5 rounded-full bg-rose-500"></span>
                        Maximum Voltages
                    </h3>
                    
                    <div className="space-y-4">
                        <div>
                            <Label className="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-2 block">400KV Bus Voltage</Label>
                            {renderField("Max Value (kV)", "max_bus_voltage_400kv.value")}
                        </div>
                        
                        <div className="pt-2 border-t border-slate-200/50">
                             <Label className="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-2 block mt-2">220KV Bus Voltage</Label>
                            {renderField("Max Value (kV)", "max_bus_voltage_220kv.value")}
                        </div>
 
                        <div className="grid grid-cols-4 items-center gap-4">
                          <Label htmlFor="max_bus_voltage.time" className="text-right text-slate-600 font-medium">
                            Time
                          </Label>
                          <Input
                            id="max_bus_voltage.time"
                            type={getGroupTime('max') === 'N/S' ? 'text' : 'time'}
                            value={getGroupTime('max')}
                            onChange={(e) => setGroupTime('max', e.target.value)}
                            className={`col-span-3 border-slate-200 focus:ring-2 focus:ring-indigo-500/20 ${getGroupTime('max') === 'N/S' ? 'bg-slate-50 text-slate-500' : ''}`}
                            disabled={getGroupTime('max') === 'N/S'}
                            required
                          />
                        </div>
                    </div>
                  </div>

                  <div className="bg-slate-50 p-4 rounded-xl border border-slate-100 space-y-4">
                    <h3 className="font-semibold text-slate-800 flex items-center gap-2">
                        <span className="w-1.5 h-1.5 rounded-full bg-blue-500"></span>
                        Minimum Voltages
                    </h3>
                    
                    <div className="space-y-4">
                        <div>
                            <Label className="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-2 block">400KV Bus Voltage</Label>
                            {renderField("Min Value (kV)", "min_bus_voltage_400kv.value")}
                        </div>
                        
                        <div className="pt-2 border-t border-slate-200/50">
                            <Label className="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-2 block mt-2">220KV Bus Voltage</Label>
                            {renderField("Min Value (kV)", "min_bus_voltage_220kv.value")}
                        </div>
 
                        <div className="grid grid-cols-4 items-center gap-4">
                          <Label htmlFor="min_bus_voltage.time" className="text-right text-slate-600 font-medium">
                            Time
                          </Label>
                          <Input
                            id="min_bus_voltage.time"
                            type={getGroupTime('min') === 'N/S' ? 'text' : 'time'}
                            value={getGroupTime('min')}
                            onChange={(e) => setGroupTime('min', e.target.value)}
                            className={`col-span-3 border-slate-200 focus:ring-2 focus:ring-indigo-500/20 ${getGroupTime('min') === 'N/S' ? 'bg-slate-50 text-slate-500' : ''}`}
                            disabled={getGroupTime('min') === 'N/S'}
                            required
                          />
                        </div>
                    </div>
                  </div>
                </>
              ) : (
                <>
                  <div className="bg-slate-50 p-4 rounded-xl border border-slate-100 space-y-4">
                    <h3 className="font-semibold text-slate-800 flex items-center gap-2">
                        <span className="w-1.5 h-1.5 rounded-full bg-rose-500"></span>
                        Maximum Values
                    </h3>
                    {renderField("Max Amps", "max.amps")}
                    {renderField("Max MW", "max.mw")}
                    {feeder.type === 'ict_feeder' && renderField("Max MVAR", "max.mvar")}
                    {renderTimeField("Time", "max.time")}
                  </div>

                  <div className="bg-slate-50 p-4 rounded-xl border border-slate-100 space-y-4">
                    <h3 className="font-semibold text-slate-800 flex items-center gap-2">
                        <span className="w-1.5 h-1.5 rounded-full bg-blue-500"></span>
                        Minimum Values
                    </h3>
                    {renderField("Min Amps", "min.amps")}
                    {renderField("Min MW", "min.mw")}
                    {feeder.type === 'ict_feeder' && renderField("Min MVAR", "min.mvar")}
                    {renderTimeField("Time", "min.time")}
                  </div>
                </>
              )}
            </div>

          </div>

          <DialogFooter className="p-6 bg-slate-50 border-t border-slate-100 flex-shrink-0">
            <Button type="button" variant="outline" onClick={onClose}>
              Cancel
            </Button>
            <Button type="submit" disabled={loading} className="bg-indigo-600 hover:bg-indigo-700 text-white">
              {loading ? 'Saving...' : 'Save Entry'}
            </Button>
          </DialogFooter>
        </form>
      </DialogContent>
    </Dialog>
  );
}
