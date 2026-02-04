import React, { useState, useEffect, useRef } from 'react';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { toast } from 'sonner';
import { Edit, Calendar, Ban } from 'lucide-react';

export default function MaxMinEntryModal({ isOpen, onClose, onSave, feeder, year, month, initialData, onPrevFeeder, onNextFeeder }) {
  const [formData, setFormData] = useState({});
  const [selectedDate, setSelectedDate] = useState('');
  const [loading, setLoading] = useState(false);
  const firstInputRef = useRef(null);

  useEffect(() => {
    if (isOpen) {
      if (initialData) {
        setFormData(initialData.data || {});
        setSelectedDate(initialData.date);
      } else {
        setFormData({});
        // Preserve selectedDate if already set (for consecutive entries)
        // If not set, user must pick date.
      }
      
      // Focus first input
      setTimeout(() => firstInputRef.current?.focus(), 0);
    }
  }, [isOpen, initialData, feeder?.id]);

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

  // Calculate min and max date for the date picker
  // Better string construction to avoid timezone issues:
  const strMonth = month.toString().padStart(2, '0');
  const minDateStr = `${year}-${strMonth}-01`;
  const maxDay = new Date(year, month, 0).getDate();
  const maxDateStr = `${year}-${strMonth}-${maxDay}`;

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

  return (
    <Dialog open={isOpen} onOpenChange={onClose}>
      <DialogContent className="max-w-2xl max-h-[90vh] p-0 gap-0 border-0 shadow-2xl overflow-hidden flex flex-col">
        <DialogHeader className="bg-gradient-to-r from-indigo-600 to-purple-600 p-6 text-white flex-shrink-0">
          <DialogTitle className="flex items-center justify-between gap-2 text-xl">
            <div className="p-2 bg-white/20 rounded-lg backdrop-blur-sm">
                {initialData ? <Edit className="w-5 h-5 text-white" /> : <Calendar className="w-5 h-5 text-white" />}
            </div>
            <span className="flex items-center gap-2">
              {initialData ? 'Edit Entry' : 'New Entry'} - {feeder.name}
            </span>
            <span className="flex items-center gap-2">
              {onPrevFeeder && (
                <Button type="button" variant="secondary" onClick={onPrevFeeder} className="bg-white/20 text-white hover:bg-white/30">
                  Previous
                </Button>
              )}
              {onNextFeeder && (
                <Button type="button" variant="secondary" onClick={onNextFeeder} className="bg-white/20 text-white hover:bg-white/30">
                  Next
                </Button>
              )}
            </span>
          </DialogTitle>
        </DialogHeader>
        
        <div className="overflow-y-auto flex-1">
        <form onSubmit={handleSubmit} className="p-6 space-y-6">
          <div className="grid grid-cols-4 items-center gap-4 bg-slate-50 p-4 rounded-xl border border-slate-100">
            <Label htmlFor="date" className="text-right font-bold text-slate-700">Date</Label>
            {initialData ? (
               <Input id="date" value={formatDate(selectedDate)} disabled className="col-span-3 bg-white" />
            ) : (
              <Input 
                id="date" 
                type="date"
                min={minDateStr}
                max={maxDateStr}
                value={selectedDate}
                onChange={(e) => setSelectedDate(e.target.value)}
                className="col-span-3 border-slate-200 focus:ring-2 focus:ring-indigo-500/20"
                required
              />
            )}
          </div>

          {feeder.type !== 'bus_station' && (
            <div className="flex justify-end">
                <Button
                    type="button"
                    variant="outline"
                    onClick={handleNotInService}
                    className="bg-red-50 text-red-600 hover:bg-red-100 border-red-200 hover:border-red-300"
                >
                    <Ban className="w-4 h-4 mr-2" />
                    Not in Service
                </Button>
            </div>
          )}

          {feeder.type === 'bus_station' && (
            <div className="space-y-6">
              <div className="space-y-4 rounded-xl border border-rose-100 bg-rose-50/30 p-4">
                <h3 className="font-bold text-rose-700 flex items-center gap-2 border-b border-rose-100 pb-2">
                    <div className="w-2 h-8 rounded-full bg-rose-500"></div>
                    Max Bus Voltages
                </h3>
                {renderField("400KV Value", "max_bus_voltage_400kv.value")}
                {renderField("220KV Value", "max_bus_voltage_220kv.value")}
                <div className="grid grid-cols-4 items-center gap-4">
                  <Label htmlFor="max_time" className="text-right text-slate-600 font-medium">Time</Label>
                  <Input
                    id="max_time"
                    type="time"
                    value={getValue("max_bus_voltage_400kv.time")}
                    onChange={(e) => handleSharedTimeChange('max', e.target.value)}
                    className="col-span-3 border-slate-200 focus:ring-2 focus:ring-indigo-500/20"
                    required
                  />
                </div>
              </div>
              
              <div className="space-y-4 rounded-xl border border-blue-100 bg-blue-50/30 p-4">
                <h3 className="font-bold text-blue-700 flex items-center gap-2 border-b border-blue-100 pb-2">
                    <div className="w-2 h-8 rounded-full bg-blue-500"></div>
                    Min Bus Voltages
                </h3>
                {renderField("400KV Value", "min_bus_voltage_400kv.value")}
                {renderField("220KV Value", "min_bus_voltage_220kv.value")}
                <div className="grid grid-cols-4 items-center gap-4">
                  <Label htmlFor="min_time" className="text-right text-slate-600 font-medium">Time</Label>
                  <Input
                    id="min_time"
                    type="time"
                    value={getValue("min_bus_voltage_400kv.time")}
                    onChange={(e) => handleSharedTimeChange('min', e.target.value)}
                    className="col-span-3 border-slate-200 focus:ring-2 focus:ring-indigo-500/20"
                    required
                  />
                </div>
              </div>
            </div>
          )}

          {(feeder.type === 'ict_feeder' || feeder.type.startsWith('feeder_')) && (
            <div className="space-y-6">
              <div className="space-y-4 rounded-xl border border-rose-100 bg-rose-50/30 p-4">
                <h3 className="font-bold text-rose-700 flex items-center gap-2 border-b border-rose-100 pb-2">
                    <div className="w-2 h-8 rounded-full bg-rose-500"></div>
                    Max Values
                </h3>
                {renderField("Amps", "max.amps")}
                {renderField("MW", "max.mw")}
                {feeder.type === 'ict_feeder' && renderField("MVAR", "max.mvar")}
                {renderTimeField("Time", "max.time")}
              </div>

              <div className="space-y-4 rounded-xl border border-blue-100 bg-blue-50/30 p-4">
                <h3 className="font-bold text-blue-700 flex items-center gap-2 border-b border-blue-100 pb-2">
                    <div className="w-2 h-8 rounded-full bg-blue-500"></div>
                    Min Values
                </h3>
                {renderField("Amps", "min.amps")}
                {renderField("MW", "min.mw")}
                {feeder.type === 'ict_feeder' && renderField("MVAR", "min.mvar")}
                {renderTimeField("Time", "min.time")}
              </div>
            </div>
          )}

          <DialogFooter className="bg-slate-50 p-6 -mx-6 -mb-6 border-t border-slate-100">
            <Button type="button" variant="outline" onClick={onClose} disabled={loading} className="border-slate-200 hover:bg-slate-100">
              Cancel
            </Button>
            <Button type="submit" disabled={loading} className="bg-gradient-to-r from-indigo-600 to-purple-600 hover:from-indigo-700 hover:to-purple-700 text-white shadow-lg shadow-indigo-200 border-0">
              {loading ? 'Saving...' : 'Save changes'}
            </Button>
          </DialogFooter>
        </form>
        </div>
      </DialogContent>
    </Dialog>
  );
}
