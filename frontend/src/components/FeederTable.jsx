import { useState } from 'react';
import axios from 'axios';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Dialog, DialogContent, DialogHeader, DialogTitle } from '@/components/ui/dialog';
import { toast } from 'sonner';
import { Edit2, Trash2, Save, X } from 'lucide-react';

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

export default function FeederTable({ feeder, entries, loading, onUpdate, onDelete }) {
  const [editingId, setEditingId] = useState(null);
  const [editValues, setEditValues] = useState({});
  const [deleteConfirmId, setDeleteConfirmId] = useState(null);

  const startEdit = (entry) => {
    setEditingId(entry.id);
    setEditValues({
      end1_import_initial: entry.end1_import_initial,
      end1_import_final: entry.end1_import_final,
      end1_export_initial: entry.end1_export_initial,
      end1_export_final: entry.end1_export_final,
      end2_import_initial: entry.end2_import_initial,
      end2_import_final: entry.end2_import_final,
      end2_export_initial: entry.end2_export_initial,
      end2_export_final: entry.end2_export_final
    });
  };

  const cancelEdit = () => {
    setEditingId(null);
    setEditValues({});
  };

  const saveEdit = async (entryId) => {
    try {
      const response = await axios.put(`${API}/entries/${entryId}`, editValues);
      onUpdate(response.data);
      setEditingId(null);
      setEditValues({});
    } catch (error) {
      toast.error('Failed to update entry');
    }
  };

  const confirmDelete = async () => {
    if (deleteConfirmId) {
      await onDelete(deleteConfirmId);
      setDeleteConfirmId(null);
    }
  };

  const formatNumber = (num) => {
    return typeof num === 'number' ? num.toFixed(2) : '0.00';
  };

  const getLossColor = (percent) => {
    if (percent > 5) return 'loss-positive';
    if (percent < 0) return 'loss-negative';
    return '';
  };

  if (loading) {
    return (
      <div className="flex items-center justify-center p-12">
        <p className="text-lg">Loading entries...</p>
      </div>
    );
  }

  if (entries.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center p-12 bg-slate-50 dark:bg-slate-800 rounded-lg border-2 border-dashed">
        <p className="text-lg font-medium">No entries found</p>
        <p className="text-sm text-slate-500 mt-1">Click "New Entry" to add your first record</p>
      </div>
    );
  }

  return (
    <>
      <div className="table-container">
        <table className="data-table">
          <thead>
            <tr>
              <th rowSpan="2">Date</th>
              <th colSpan="4" className="text-center">{feeder.end1_name} Import</th>
              <th colSpan="4" className="text-center">{feeder.end1_name} Export</th>
              <th colSpan="4" className="text-center">{feeder.end2_name} Import</th>
              <th colSpan="4" className="text-center">{feeder.end2_name} Export</th>
              <th rowSpan="2">% Loss</th>
              <th rowSpan="2">Actions</th>
            </tr>
            <tr>
              <th>Initial</th>
              <th>Final</th>
              <th>MF</th>
              <th>Consumption</th>
              <th>Initial</th>
              <th>Final</th>
              <th>MF</th>
              <th>Consumption</th>
              <th>Initial</th>
              <th>Final</th>
              <th>MF</th>
              <th>Consumption</th>
              <th>Initial</th>
              <th>Final</th>
              <th>MF</th>
              <th>Consumption</th>
            </tr>
          </thead>
          <tbody>
            {entries.map((entry) => {
              const isEditing = editingId === entry.id;
              
              return (
                <tr key={entry.id} data-testid={`entry-row-${entry.id}`}>
                  <td>{entry.date}</td>
                  
                  {/* End 1 Import */}
                  <td className="numeric-cell">
                    {isEditing ? (
                      <Input
                        type="number"
                        step="0.01"
                        value={editValues.end1_import_initial}
                        onChange={(e) => setEditValues({...editValues, end1_import_initial: parseFloat(e.target.value)})}
                        className="w-24 h-8 text-xs"
                      />
                    ) : formatNumber(entry.end1_import_initial)}
                  </td>
                  <td className="numeric-cell">
                    {isEditing ? (
                      <Input
                        type="number"
                        step="0.01"
                        value={editValues.end1_import_final}
                        onChange={(e) => setEditValues({...editValues, end1_import_final: parseFloat(e.target.value)})}
                        className="w-24 h-8 text-xs"
                      />
                    ) : formatNumber(entry.end1_import_final)}
                  </td>
                  <td className="numeric-cell">{feeder.end1_import_mf}</td>
                  <td className="numeric-cell font-medium">{formatNumber(entry.end1_import_consumption)}</td>
                  
                  {/* End 1 Export */}
                  <td className="numeric-cell">
                    {isEditing ? (
                      <Input
                        type="number"
                        step="0.01"
                        value={editValues.end1_export_initial}
                        onChange={(e) => setEditValues({...editValues, end1_export_initial: parseFloat(e.target.value)})}
                        className="w-24 h-8 text-xs"
                      />
                    ) : formatNumber(entry.end1_export_initial)}
                  </td>
                  <td className="numeric-cell">
                    {isEditing ? (
                      <Input
                        type="number"
                        step="0.01"
                        value={editValues.end1_export_final}
                        onChange={(e) => setEditValues({...editValues, end1_export_final: parseFloat(e.target.value)})}
                        className="w-24 h-8 text-xs"
                      />
                    ) : formatNumber(entry.end1_export_final)}
                  </td>
                  <td className="numeric-cell">{feeder.end1_export_mf}</td>
                  <td className="numeric-cell font-medium">{formatNumber(entry.end1_export_consumption)}</td>
                  
                  {/* End 2 Import */}
                  <td className="numeric-cell">
                    {isEditing ? (
                      <Input
                        type="number"
                        step="0.01"
                        value={editValues.end2_import_initial}
                        onChange={(e) => setEditValues({...editValues, end2_import_initial: parseFloat(e.target.value)})}
                        className="w-24 h-8 text-xs"
                      />
                    ) : formatNumber(entry.end2_import_initial)}
                  </td>
                  <td className="numeric-cell">
                    {isEditing ? (
                      <Input
                        type="number"
                        step="0.01"
                        value={editValues.end2_import_final}
                        onChange={(e) => setEditValues({...editValues, end2_import_final: parseFloat(e.target.value)})}
                        className="w-24 h-8 text-xs"
                      />
                    ) : formatNumber(entry.end2_import_final)}
                  </td>
                  <td className="numeric-cell">{feeder.end2_import_mf}</td>
                  <td className="numeric-cell font-medium">{formatNumber(entry.end2_import_consumption)}</td>
                  
                  {/* End 2 Export */}
                  <td className="numeric-cell">
                    {isEditing ? (
                      <Input
                        type="number"
                        step="0.01"
                        value={editValues.end2_export_initial}
                        onChange={(e) => setEditValues({...editValues, end2_export_initial: parseFloat(e.target.value)})}
                        className="w-24 h-8 text-xs"
                      />
                    ) : formatNumber(entry.end2_export_initial)}
                  </td>
                  <td className="numeric-cell">
                    {isEditing ? (
                      <Input
                        type="number"
                        step="0.01"
                        value={editValues.end2_export_final}
                        onChange={(e) => setEditValues({...editValues, end2_export_final: parseFloat(e.target.value)})}
                        className="w-24 h-8 text-xs"
                      />
                    ) : formatNumber(entry.end2_export_final)}
                  </td>
                  <td className="numeric-cell">{feeder.end2_export_mf}</td>
                  <td className="numeric-cell font-medium">{formatNumber(entry.end2_export_consumption)}</td>
                  
                  {/* Loss % */}
                  <td className={`numeric-cell font-bold ${getLossColor(entry.loss_percent)}`}>
                    {formatNumber(entry.loss_percent)}%
                  </td>
                  
                  {/* Actions */}
                  <td>
                    <div className="flex gap-1">
                      {isEditing ? (
                        <>
                          <Button 
                            size="icon" 
                            variant="ghost" 
                            className="h-8 w-8"
                            onClick={() => saveEdit(entry.id)}
                            data-testid="save-edit-button"
                          >
                            <Save className="w-4 h-4" />
                          </Button>
                          <Button 
                            size="icon" 
                            variant="ghost" 
                            className="h-8 w-8"
                            onClick={cancelEdit}
                            data-testid="cancel-edit-button"
                          >
                            <X className="w-4 h-4" />
                          </Button>
                        </>
                      ) : (
                        <>
                          <Button 
                            size="icon" 
                            variant="ghost" 
                            className="h-8 w-8"
                            onClick={() => startEdit(entry)}
                            data-testid="edit-button"
                          >
                            <Edit2 className="w-4 h-4" />
                          </Button>
                          <Button 
                            size="icon" 
                            variant="ghost" 
                            className="h-8 w-8 text-red-600 hover:text-red-700"
                            onClick={() => setDeleteConfirmId(entry.id)}
                            data-testid="delete-button"
                          >
                            <Trash2 className="w-4 h-4" />
                          </Button>
                        </>
                      )}
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* Delete Confirmation Dialog */}
      <Dialog open={!!deleteConfirmId} onOpenChange={() => setDeleteConfirmId(null)}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Confirm Delete</DialogTitle>
          </DialogHeader>
          <p className="text-sm text-slate-600 dark:text-slate-400">
            Are you sure you want to delete this entry? This action cannot be undone.
          </p>
          <div className="flex gap-2 justify-end mt-4">
            <Button variant="outline" onClick={() => setDeleteConfirmId(null)}>
              Cancel
            </Button>
            <Button variant="destructive" onClick={confirmDelete}>
              Delete
            </Button>
          </div>
        </DialogContent>
      </Dialog>
    </>
  );
}
