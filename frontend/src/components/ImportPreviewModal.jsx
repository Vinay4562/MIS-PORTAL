
import React from 'react';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { ScrollArea } from "@/components/ui/scroll-area";

export default function ImportPreviewModal({ isOpen, onClose, data, onConfirm, loading }) {
  return (
    <Dialog open={isOpen} onOpenChange={onClose}>
      <DialogContent className="max-w-4xl max-h-[90vh] flex flex-col">
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
                <TableRow>
                  <TableHead>Date</TableHead>
                  <TableHead>End1 Import Final</TableHead>
                  <TableHead>End1 Export Final</TableHead>
                  <TableHead>End2 Import Final</TableHead>
                  <TableHead>End2 Export Final</TableHead>
                  <TableHead>Status</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {data.map((row, index) => (
                  <TableRow key={index} className={row.exists ? "bg-yellow-50 dark:bg-yellow-900/20 opacity-70" : ""}>
                    <TableCell>{row.date}</TableCell>
                    <TableCell>{row.end1_import_final}</TableCell>
                    <TableCell>{row.end1_export_final}</TableCell>
                    <TableCell>{row.end2_import_final}</TableCell>
                    <TableCell>{row.end2_export_final}</TableCell>
                    <TableCell>
                      {row.exists ? (
                        <span className="text-yellow-600 font-medium text-xs">Existing</span>
                      ) : (
                        <span className="text-green-600 font-medium text-xs">New</span>
                      )}
                    </TableCell>
                  </TableRow>
                ))}
                {data.length === 0 && (
                   <TableRow>
                     <TableCell colSpan={5} className="text-center h-24 text-slate-500">
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
