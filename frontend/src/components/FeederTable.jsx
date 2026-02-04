import { useState } from 'react';
import axios from 'axios';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Dialog, DialogContent, DialogHeader, DialogTitle } from '@/components/ui/dialog';
import { toast } from 'sonner';
import { Edit2, Trash2, Save, X } from 'lucide-react';
import { formatDate } from '@/lib/utils';

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
    if (percent > 5) return 'text-red-600 font-bold';
    if (percent < 0) return 'text-green-600 font-bold';
    return 'text-slate-900 dark:text-slate-100';
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
      <div className="flex flex-col items-center justify-center p-12 bg-slate-50 dark:bg-slate-800 rounded-lg border-2 border-dashed border-slate-300 dark:border-slate-700">
        <p className="text-lg font-medium text-slate-700 dark:text-slate-300">No entries found</p>
        <p className="text-sm text-slate-500 mt-1">Click "New Entry" to add your first record</p>
      </div>
    );
  }

  const renderCell = (value, isEditing, fieldName, type = 'number') => {
    if (isEditing) {
      return (
        <Input
          type={type}
          step="0.01"
          value={editValues[fieldName]}
          onChange={(e) => setEditValues({ ...editValues, [fieldName]: parseFloat(e.target.value) })}
          className="h-7 w-24 text-xs px-2 bg-white dark:bg-slate-900 border-slate-200"
        />
      );
    }
    return <span className="font-medium">{formatNumber(value)}</span>;
  };

  const renderSectionHeader = (title) => (
    <th colSpan={4} className="px-4 py-3 font-medium text-center border-r border-slate-200 dark:border-slate-700 bg-slate-100 dark:bg-slate-800">
      {title}
    </th>
  );

  const renderSubHeaders = () => (
    <>
      <th className="px-2 py-2 text-center font-normal border-r border-slate-200 dark:border-slate-700 min-w-[80px] text-slate-500">Initial</th>
      <th className="px-2 py-2 text-center font-normal border-r border-slate-200 dark:border-slate-700 min-w-[80px] text-slate-500">Final</th>
      <th className="px-2 py-2 text-center font-normal border-r border-slate-200 dark:border-slate-700 min-w-[50px] text-slate-500">MF</th>
      <th className="px-2 py-2 text-center font-normal border-r border-slate-200 dark:border-slate-700 min-w-[80px] text-slate-500">Cons.</th>
    </>
  );

  return (
    <>
      <div className="rounded-md border border-slate-200 dark:border-slate-700 overflow-hidden shadow-sm bg-white dark:bg-slate-900">
        <div className="overflow-x-auto">
          <table className="w-full text-sm text-left border-collapse">
            <thead className="bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 border-b border-slate-200 dark:border-slate-700">
              <tr>
                <th rowSpan={2} className="px-4 py-3 font-medium text-center border-r border-slate-200 dark:border-slate-700 sticky left-0 bg-slate-100 dark:bg-slate-800 z-20">
                  Date
                </th>
                {renderSectionHeader(`${feeder.end1_name} Import`)}
                {renderSectionHeader(`${feeder.end1_name} Export`)}
                {renderSectionHeader(`${feeder.end2_name} Import`)}
                {renderSectionHeader(`${feeder.end2_name} Export`)}
                <th rowSpan={2} className="px-4 py-3 font-medium text-center border-r border-slate-200 dark:border-slate-700 bg-slate-100 dark:bg-slate-800">
                  % Loss
                </th>
                <th rowSpan={2} className="px-4 py-3 font-medium text-center sticky right-0 bg-slate-100 dark:bg-slate-800 z-20">
                  Actions
                </th>
              </tr>
              <tr className="bg-slate-50 dark:bg-slate-800/50 text-xs">
                {renderSubHeaders()}
                {renderSubHeaders()}
                {renderSubHeaders()}
                {renderSubHeaders()}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200 dark:divide-slate-700">
              {entries.map((entry) => {
                const isEditing = editingId === entry.id;
                return (
                  <tr key={entry.id} className="hover:bg-slate-50 dark:hover:bg-slate-800/50 group transition-colors" data-testid={`entry-row-${entry.id}`}>
                    <td className="px-4 py-3 text-center font-medium whitespace-nowrap border-r border-slate-200 dark:border-slate-700 sticky left-0 bg-white dark:bg-slate-900 group-hover:bg-slate-50 dark:group-hover:bg-slate-800/50 z-10">
                      {formatDate(entry.date)}
                    </td>

                    {/* End 1 Import */}
                    <td className="px-2 py-3 text-center border-r border-slate-200 dark:border-slate-700 text-slate-500">
                      {renderCell(entry.end1_import_initial, isEditing, 'end1_import_initial')}
                    </td>
                    <td className="px-2 py-3 text-center border-r border-slate-200 dark:border-slate-700">
                      {renderCell(entry.end1_import_final, isEditing, 'end1_import_final')}
                    </td>
                    <td className="px-2 py-3 text-center border-r border-slate-200 dark:border-slate-700 text-slate-400">
                      {feeder.end1_import_mf}
                    </td>
                    <td className="px-2 py-3 text-center border-r border-slate-200 dark:border-slate-700 font-bold text-blue-600 dark:text-blue-400">
                      {formatNumber(entry.end1_import_consumption)}
                    </td>

                    {/* End 1 Export */}
                    <td className="px-2 py-3 text-center border-r border-slate-200 dark:border-slate-700 text-slate-500">
                      {renderCell(entry.end1_export_initial, isEditing, 'end1_export_initial')}
                    </td>
                    <td className="px-2 py-3 text-center border-r border-slate-200 dark:border-slate-700">
                      {renderCell(entry.end1_export_final, isEditing, 'end1_export_final')}
                    </td>
                    <td className="px-2 py-3 text-center border-r border-slate-200 dark:border-slate-700 text-slate-400">
                      {feeder.end1_export_mf}
                    </td>
                    <td className="px-2 py-3 text-center border-r border-slate-200 dark:border-slate-700 font-bold text-blue-600 dark:text-blue-400">
                      {formatNumber(entry.end1_export_consumption)}
                    </td>

                    {/* End 2 Import */}
                    <td className="px-2 py-3 text-center border-r border-slate-200 dark:border-slate-700 text-slate-500">
                      {renderCell(entry.end2_import_initial, isEditing, 'end2_import_initial')}
                    </td>
                    <td className="px-2 py-3 text-center border-r border-slate-200 dark:border-slate-700">
                      {renderCell(entry.end2_import_final, isEditing, 'end2_import_final')}
                    </td>
                    <td className="px-2 py-3 text-center border-r border-slate-200 dark:border-slate-700 text-slate-400">
                      {feeder.end2_import_mf}
                    </td>
                    <td className="px-2 py-3 text-center border-r border-slate-200 dark:border-slate-700 font-bold text-blue-600 dark:text-blue-400">
                      {formatNumber(entry.end2_import_consumption)}
                    </td>

                    {/* End 2 Export */}
                    <td className="px-2 py-3 text-center border-r border-slate-200 dark:border-slate-700 text-slate-500">
                      {renderCell(entry.end2_export_initial, isEditing, 'end2_export_initial')}
                    </td>
                    <td className="px-2 py-3 text-center border-r border-slate-200 dark:border-slate-700">
                      {renderCell(entry.end2_export_final, isEditing, 'end2_export_final')}
                    </td>
                    <td className="px-2 py-3 text-center border-r border-slate-200 dark:border-slate-700 text-slate-400">
                      {feeder.end2_export_mf}
                    </td>
                    <td className="px-2 py-3 text-center border-r border-slate-200 dark:border-slate-700 font-bold text-blue-600 dark:text-blue-400">
                      {formatNumber(entry.end2_export_consumption)}
                    </td>

                    {/* Loss % */}
                    <td className={`px-4 py-3 text-center font-bold border-r border-slate-200 dark:border-slate-700 ${getLossColor(entry.loss_percent)}`}>
                      {formatNumber(entry.loss_percent)}%
                    </td>

                    {/* Actions */}
                    <td className="px-2 py-2 text-center sticky right-0 bg-white dark:bg-slate-900 group-hover:bg-slate-50 dark:group-hover:bg-slate-800/50 z-10">
                      <div className="flex items-center justify-center gap-1">
                        {isEditing ? (
                          <>
                            <Button
                              size="icon"
                              variant="ghost"
                              className="h-8 w-8 text-green-600 hover:text-green-700 hover:bg-green-50"
                              onClick={() => saveEdit(entry.id)}
                              data-testid="save-edit-button"
                            >
                              <Save className="w-4 h-4" />
                            </Button>
                            <Button
                              size="icon"
                              variant="ghost"
                              className="h-8 w-8 text-slate-500 hover:text-slate-700 hover:bg-slate-50"
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
                              className="h-8 w-8 text-blue-600 hover:text-blue-700 hover:bg-blue-50"
                              onClick={() => startEdit(entry)}
                              data-testid="edit-button"
                            >
                              <Edit2 className="w-4 h-4" />
                            </Button>
                            <Button
                              size="icon"
                              variant="ghost"
                              className="h-8 w-8 text-red-600 hover:text-red-700 hover:bg-red-50"
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
      </div>

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
