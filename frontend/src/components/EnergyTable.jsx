import { Edit2, Trash2 } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { formatDate } from '@/lib/utils';

export default function EnergyTable({ sheet, entries, loading, onEdit, onDelete }) {
  if (loading) {
    return <div className="p-8 text-center">Loading entries...</div>;
  }

  if (entries.length === 0) {
    return <div className="p-8 text-center text-slate-500">No entries found for this period</div>;
  }

  return (
    <div className="rounded-md border border-slate-200 dark:border-slate-700 overflow-hidden">
      <div className="overflow-x-auto">
        <table className="w-full text-sm text-center">
            <thead className="bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 border-b border-slate-200 dark:border-slate-700">
                <tr>
                    <th className="px-4 py-3 font-medium text-center border-r border-slate-200 dark:border-slate-700 sticky left-0 bg-slate-100 dark:bg-slate-800 z-10">Date</th>
                    {sheet.meters.map(meter => (
                        <th key={meter.id} colSpan={4} className="px-4 py-3 font-medium text-center border-r border-slate-200 dark:border-slate-700">
                            {meter.name} ({meter.unit})
                        </th>
                    ))}
                    <th className="px-4 py-3 font-medium text-center border-r">Total</th>
                    <th className="px-4 py-3 font-medium text-center sticky right-0 bg-slate-100 dark:bg-slate-800 z-10">Actions</th>
                </tr>
                <tr className="text-xs text-slate-500 dark:text-slate-400 bg-slate-50 dark:bg-slate-800/50">
                    <th className="sticky left-0 bg-slate-50 dark:bg-slate-800/50 border-r z-10"></th>
                    {sheet.meters.map(meter => (
                        <>
                            <th className="px-2 py-2 text-center font-normal border-r min-w-[80px]">Initial</th>
                            <th className="px-2 py-2 text-center font-normal border-r min-w-[80px]">Final</th>
                            <th className="px-2 py-2 text-center font-normal border-r min-w-[50px]">MF</th>
                            <th className="px-2 py-2 text-center font-normal border-r min-w-[80px]">Cons.</th>
                        </>
                    ))}
                    <th className="px-4 py-2 text-center font-normal border-r">Consumption</th>
                    <th className="sticky right-0 bg-slate-50 dark:bg-slate-800/50 z-10"></th>
                </tr>
            </thead>
            <tbody className="divide-y divide-slate-200 dark:divide-slate-700 bg-white dark:bg-slate-900">
                {entries.map(entry => (
                    <tr key={entry.id} className="hover:bg-slate-50 dark:hover:bg-slate-800/50 group">
                        <td className="px-4 py-3 font-medium whitespace-nowrap border-r sticky left-0 bg-white dark:bg-slate-900 group-hover:bg-slate-50 dark:group-hover:bg-slate-800/50 z-10">
                            {formatDate(entry.date)}
                        </td>
                        {sheet.meters.map(meter => {
                            const reading = entry.readings.find(r => r.meter_id === meter.id);
                            return (
                                <>
                                    <td className="px-2 py-3 text-center border-r text-slate-500">{reading?.initial?.toFixed(2) || '-'}</td>
                                    <td className="px-2 py-3 text-center border-r font-medium">{reading?.final?.toFixed(2) || '-'}</td>
                                    <td className="px-2 py-3 text-center border-r text-slate-400">{meter.mf}</td>
                                    <td className="px-2 py-3 text-center border-r font-bold text-blue-600 dark:text-blue-400">
                                        {reading?.consumption?.toFixed(2) || '-'}
                                    </td>
                                </>
                            );
                        })}
                        <td className="px-4 py-3 text-center font-bold text-slate-900 dark:text-white border-r">
                            {entry.total_consumption?.toFixed(2)}
                        </td>
                        <td className="px-2 py-2 text-center sticky right-0 bg-white dark:bg-slate-900 group-hover:bg-slate-50 dark:group-hover:bg-slate-800/50 z-10">
                            <div className="flex items-center justify-center gap-1">
                                <Button size="icon" variant="ghost" className="h-8 w-8 text-blue-600 hover:text-blue-700 hover:bg-blue-50" onClick={() => onEdit(entry)}>
                                    <Edit2 className="w-4 h-4" />
                                </Button>
                                <Button size="icon" variant="ghost" className="h-8 w-8 text-red-600 hover:text-red-700 hover:bg-red-50" onClick={() => onDelete(entry)}>
                                    <Trash2 className="w-4 h-4" />
                                </Button>
                            </div>
                        </td>
                    </tr>
                ))}
            </tbody>
        </table>
      </div>
    </div>
  );
}
