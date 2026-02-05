import React from 'react';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { ScrollArea } from "@/components/ui/scroll-area";
 
export default function MaxMinImportPreviewModal({ isOpen, onClose, data, feederType, onConfirm, loading }) {
  const renderHeaders = () => {
    if (feederType === 'bus_station') {
      return (
        <TableRow>
          <TableHead>Date</TableHead>
          <TableHead>Max 400KV</TableHead>
          <TableHead>Max 220KV</TableHead>
          <TableHead>Max Time</TableHead>
          <TableHead>Min 400KV</TableHead>
          <TableHead>Min 220KV</TableHead>
          <TableHead>Min Time</TableHead>
          <TableHead>Station Max MW</TableHead>
          <TableHead>Station Time</TableHead>
          <TableHead>MVAR</TableHead>
          <TableHead>Status</TableHead>
        </TableRow>
      );
    }
    if (feederType === 'ict_feeder') {
      return (
        <TableRow>
          <TableHead>Date</TableHead>
          <TableHead>Max Amps</TableHead>
          <TableHead>Max MW</TableHead>
          <TableHead>Max MVAR</TableHead>
          <TableHead>Max Time</TableHead>
          <TableHead>Min Amps</TableHead>
          <TableHead>Min MW</TableHead>
          <TableHead>Min MVAR</TableHead>
          <TableHead>Min Time</TableHead>
          <TableHead>Avg Amps</TableHead>
          <TableHead>Avg MW</TableHead>
          <TableHead>Status</TableHead>
        </TableRow>
      );
    }
    return (
      <TableRow>
        <TableHead>Date</TableHead>
        <TableHead>Max Amps</TableHead>
        <TableHead>Max MW</TableHead>
        <TableHead>Max Time</TableHead>
        <TableHead>Min Amps</TableHead>
        <TableHead>Min MW</TableHead>
        <TableHead>Min Time</TableHead>
        <TableHead>Avg Amps</TableHead>
        <TableHead>Avg MW</TableHead>
        <TableHead>Status</TableHead>
      </TableRow>
    );
  };
 
  const renderRow = (row, idx) => {
    const d = row.data || {};
    if (feederType === 'bus_station') {
      const maxTime = d.max_bus_voltage_400kv?.time || d.max_bus_voltage_220kv?.time || '-';
      const minTime = d.min_bus_voltage_400kv?.time || d.min_bus_voltage_220kv?.time || '-';
      return (
        <TableRow key={idx} className={row.exists ? "bg-yellow-50 dark:bg-yellow-900/20 opacity-70" : ""}>
          <TableCell>{row.date}</TableCell>
          <TableCell>{d.max_bus_voltage_400kv?.value ?? ''}</TableCell>
          <TableCell>{d.max_bus_voltage_220kv?.value ?? ''}</TableCell>
          <TableCell>{maxTime}</TableCell>
          <TableCell>{d.min_bus_voltage_400kv?.value ?? ''}</TableCell>
          <TableCell>{d.min_bus_voltage_220kv?.value ?? ''}</TableCell>
          <TableCell>{minTime}</TableCell>
          <TableCell>{d.station_load?.max_mw ?? ''}</TableCell>
          <TableCell>{d.station_load?.time ?? ''}</TableCell>
          <TableCell>{d.station_load?.mvar ?? ''}</TableCell>
          <TableCell>{row.exists ? <span className="text-yellow-600 font-medium text-xs">Existing</span> : <span className="text-green-600 font-medium text-xs">New</span>}</TableCell>
        </TableRow>
      );
    }
    if (feederType === 'ict_feeder') {
      return (
        <TableRow key={idx} className={row.exists ? "bg-yellow-50 dark:bg-yellow-900/20 opacity-70" : ""}>
          <TableCell>{row.date}</TableCell>
          <TableCell>{d.max?.amps ?? ''}</TableCell>
          <TableCell>{d.max?.mw ?? ''}</TableCell>
          <TableCell>{d.max?.mvar ?? ''}</TableCell>
          <TableCell>{d.max?.time ?? ''}</TableCell>
          <TableCell>{d.min?.amps ?? ''}</TableCell>
          <TableCell>{d.min?.mw ?? ''}</TableCell>
          <TableCell>{d.min?.mvar ?? ''}</TableCell>
          <TableCell>{d.min?.time ?? ''}</TableCell>
          <TableCell>{d.avg?.amps ?? ''}</TableCell>
          <TableCell>{d.avg?.mw ?? ''}</TableCell>
          <TableCell>{row.exists ? <span className="text-yellow-600 font-medium text-xs">Existing</span> : <span className="text-green-600 font-medium text-xs">New</span>}</TableCell>
        </TableRow>
      );
    }
    return (
      <TableRow key={idx} className={row.exists ? "bg-yellow-50 dark:bg-yellow-900/20 opacity-70" : ""}>
        <TableCell>{row.date}</TableCell>
        <TableCell>{d.max?.amps ?? ''}</TableCell>
        <TableCell>{d.max?.mw ?? ''}</TableCell>
        <TableCell>{d.max?.time ?? ''}</TableCell>
        <TableCell>{d.min?.amps ?? ''}</TableCell>
        <TableCell>{d.min?.mw ?? ''}</TableCell>
        <TableCell>{d.min?.time ?? ''}</TableCell>
        <TableCell>{d.avg?.amps ?? ''}</TableCell>
        <TableCell>{d.avg?.mw ?? ''}</TableCell>
        <TableCell>{row.exists ? <span className="text-yellow-600 font-medium text-xs">Existing</span> : <span className="text-green-600 font-medium text-xs">New</span>}</TableCell>
      </TableRow>
    );
  };
 
  return (
    <Dialog open={isOpen} onOpenChange={onClose}>
      <DialogContent className="max-w-5xl max-h-[90vh] flex flex-col">
        <DialogHeader>
          <DialogTitle>Import Preview</DialogTitle>
        </DialogHeader>
        <div className="flex-1 overflow-hidden flex flex-col min-h-0">
          <div className="text-sm text-slate-500 mb-2 flex justify-between shrink-0">
            <span>Found {data.length} records. {data.filter(r => r.exists).length > 0 ? `${data.filter(r => r.exists).length} existing records will be skipped.` : ""}</span>
          </div>
          <ScrollArea className="h-[60vh] border rounded-md">
            <Table>
              <TableHeader>
                {renderHeaders()}
              </TableHeader>
              <TableBody>
                {data.map((row, idx) => renderRow(row, idx))}
                {data.length === 0 && (
                  <TableRow>
                    <TableCell colSpan={12} className="text-center h-24 text-slate-500">
                      No new data found to import.
                    </TableCell>
                  </TableRow>
                )}
              </TableBody>
            </Table>
          </ScrollArea>
        </div>
        <DialogFooter>
          <Button variant="outline" onClick={onClose} disabled={loading}>Cancel</Button>
          <Button onClick={onConfirm} disabled={loading || data.length === 0}>
            {loading ? "Importing..." : "Import"}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
