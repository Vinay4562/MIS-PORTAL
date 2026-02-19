import React from 'react';
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { BlockLoader } from "@/components/ui/loader";

import { ChevronLeft, ChevronRight } from 'lucide-react';
import { Button } from '@/components/ui/button';

export function ReportPreviewModal({ isOpen, onClose, title, data, loading, year, month, subtitle, onPrev, onNext, hasNext, date, reportId }) {
  
  const renderContent = () => {
    if (loading) {
      return <BlockLoader text="Loading preview..." />;
    }

    if (!data) return <div className="text-center py-4">No data available</div>;

    // Specific rendering for Boundary Meter Report
    if (title.includes("Boundary Meter")) {
        const { report_data, prev_month_str, current_month_end_str } = data;
        
        return (
            <div className="border rounded-md overflow-x-auto">
                <Table>
                    <TableHeader>
                        <TableRow>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">Sl. No.</TableHead>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[200px]">Name of Feeder</TableHead>
                            <TableHead colSpan={2} className="text-center border bg-muted h-auto py-2">Timing</TableHead>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">Initial Reading</TableHead>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">Final Reading</TableHead>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">Diff.</TableHead>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">M.F.</TableHead>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">Consumption</TableHead>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">Remarks</TableHead>
                        </TableRow>
                        <TableRow>
                            <TableHead className="text-center border bg-muted h-auto py-2">From</TableHead>
                            <TableHead className="text-center border bg-muted h-auto py-2">To</TableHead>
                        </TableRow>
                    </TableHeader>
                    <TableBody>
                        {report_data && report_data.length > 0 ? (
                            report_data.map((row, index) => (
                                <TableRow key={index}>
                                    <TableCell className="text-center border p-2">{index + 1}</TableCell>
                                    <TableCell className="border p-2 whitespace-pre-wrap">{row.name}</TableCell>
                                    <TableCell className="text-center border p-2 whitespace-pre-wrap">{prev_month_str}{'\n'}12:00 Hrs</TableCell>
                                    <TableCell className="text-center border p-2 whitespace-pre-wrap">{current_month_end_str}{'\n'}12:00 Hrs</TableCell>
                                    <TableCell className="text-center border p-2">{row.initial?.toFixed(2)}</TableCell>
                                    <TableCell className="text-center border p-2">{row.final?.toFixed(2)}</TableCell>
                                    <TableCell className="text-center border p-2">{row.diff?.toFixed(2)}</TableCell>
                                    <TableCell className="text-center border p-2">{row.mf}</TableCell>
                                    <TableCell className="text-center border p-2 font-medium">{row.consumption?.toFixed(2)}</TableCell>
                                    <TableCell className="text-center border p-2">-</TableCell>
                                </TableRow>
                            ))
                        ) : (
                            <TableRow>
                                <TableCell colSpan={10} className="text-center py-4">No entries found for this period</TableCell>
                            </TableRow>
                        )}
                    </TableBody>
                </Table>
            </div>
        );
    }

    if (title.includes("Daily Max MVA")) {
        return (
            <div className="border rounded-md overflow-x-auto">
                <Table>
                    <TableHeader>
                        <TableRow>
                            <TableHead className="text-center border bg-muted h-auto py-2">Date</TableHead>
                            <TableHead className="text-center border bg-muted h-auto py-2">MW</TableHead>
                            <TableHead className="text-center border bg-muted h-auto py-2">MVAR</TableHead>
                            <TableHead className="text-center border bg-muted h-auto py-2">Time</TableHead>
                            <TableHead className="text-center border bg-muted h-auto py-2">MVA</TableHead>
                        </TableRow>
                    </TableHeader>
                    <TableBody>
                        {data && data.length > 0 ? (
                            data.map((row, index) => (
                                <TableRow key={index}>
                                    <TableCell className="text-center border p-2">{row.date}</TableCell>
                                    <TableCell className="text-center border p-2">{row.mw}</TableCell>
                                    <TableCell className="text-center border p-2">{row.mvar}</TableCell>
                                    <TableCell className="text-center border p-2">{row.time}</TableCell>
                                    <TableCell className="text-center border p-2">{row.mva}</TableCell>
                                </TableRow>
                            ))
                        ) : (
                            <TableRow>
                                <TableCell colSpan={5} className="text-center py-4">No entries found for this period</TableCell>
                            </TableRow>
                        )}
                    </TableBody>
                </Table>
            </div>
        );
    }

    if (title === "KPI") {
        const { lines, icts } = data || {};
        
        return (
            <div className="space-y-8">
                {/* Sheet 1: Over Loading of Lines */}
                <div className="border rounded-md overflow-x-auto">
                    <h3 className="font-bold p-2 bg-muted border-b">Over Loading of Lines</h3>
                    <Table>
                        <TableHeader>
                            <TableRow>
                                <TableHead className="text-center border bg-muted h-auto py-2">Sl. No.</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Name of Zone</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Circle</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Name of feeder</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Length in KM</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Length in CKM</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Conductor</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Capacity</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Avg.Loading (Amps)</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Max.Line loading (Amps)</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">% line loading</TableHead>
                            </TableRow>
                        </TableHeader>
                        <TableBody>
                            {lines && lines.length > 0 ? (
                                lines.map((row, index) => (
                                    <TableRow key={index}>
                                        <TableCell className="text-center border p-2">{row.sl_no}</TableCell>
                                        <TableCell className="text-center border p-2">{row.zone}</TableCell>
                                        <TableCell className="text-center border p-2">{row.circle}</TableCell>
                                        <TableCell className="text-center border p-2">{row.feeder_name}</TableCell>
                                        <TableCell className="text-center border p-2">{row.length_km}</TableCell>
                                        <TableCell className="text-center border p-2">{row.length_ckm}</TableCell>
                                        <TableCell className="text-center border p-2">{row.conductor}</TableCell>
                                        <TableCell className="text-center border p-2">{row.capacity}</TableCell>
                                        <TableCell className="text-center border p-2">{row.avg_loading}</TableCell>
                                        <TableCell className="text-center border p-2">{row.max_loading}</TableCell>
                                        <TableCell className="text-center border p-2">{row.pct_loading}</TableCell>
                                    </TableRow>
                                ))
                            ) : (
                                <TableRow>
                                    <TableCell colSpan={11} className="text-center py-4">No data available</TableCell>
                                </TableRow>
                            )}
                        </TableBody>
                    </Table>
                </div>

                {/* Sheet 2: ICT'S */}
                <div className="border rounded-md overflow-x-auto">
                    <h3 className="font-bold p-2 bg-muted border-b">ICT'S</h3>
                    <Table>
                        <TableHeader>
                            <TableRow>
                                <TableHead className="text-center border bg-muted h-auto py-2">Sl.No</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Name of Zone</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">TL&ss Circle</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Name of Substation</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Existing ICT capacity (MVA)</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Proposed augmentation</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Max. Demand (MW)</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Average load in MW</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Average Percentage of ICT loading</TableHead>
                            </TableRow>
                        </TableHeader>
                        <TableBody>
                            {icts && icts.length > 0 ? (
                                icts.map((row, index) => (
                                    <TableRow key={index}>
                                        <TableCell className="text-center border p-2">{row.sl_no}</TableCell>
                                        <TableCell className="text-center border p-2">{row.zone}</TableCell>
                                        <TableCell className="text-center border p-2">{row.circle}</TableCell>
                                        <TableCell className="text-center border p-2">{row.substation}</TableCell>
                                        <TableCell className="text-center border p-2">{row.capacity_name}</TableCell>
                                        <TableCell className="text-center border p-2">{row.proposed}</TableCell>
                                        <TableCell className="text-center border p-2">{row.max_demand}</TableCell>
                                        <TableCell className="text-center border p-2">{row.avg_load}</TableCell>
                                        <TableCell className="text-center border p-2">{row.pct_loading}</TableCell>
                                    </TableRow>
                                ))
                            ) : (
                                <TableRow>
                                    <TableCell colSpan={9} className="text-center py-4">No data available</TableCell>
                                </TableRow>
                            )}
                        </TableBody>
                    </Table>
                </div>
            </div>
        );
    }

    if (title === "Fortnight" || title === "Max-Min Daily Report") {
        if (!data || !data.periods) return <div className="text-center py-4">No data available</div>;
        
        return (
            <div className="w-full">
                <Tabs defaultValue={data.periods[0]?.name || "1-15"} className="w-full">
                    <TabsList className="grid w-full grid-cols-3 mb-4">
                        {data.periods.map(p => (
                            <TabsTrigger key={p.name} value={p.name}>{p.name}</TabsTrigger>
                        ))}
                    </TabsList>
                    {data.periods.map(p => (
                        <TabsContent key={p.name} value={p.name} className="space-y-4">
                            <div className="border rounded-md overflow-x-auto">
                                <Table>
                                    <TableHeader>
                                        <TableRow>
                                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">Sl. No.</TableHead>
                                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[200px]">Name of Feeder</TableHead>
                                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">Rating</TableHead>
                                            <TableHead colSpan={4} className="text-center border bg-muted h-auto py-2">Max Demand Reached During</TableHead>
                                            <TableHead colSpan={4} className="text-center border bg-muted h-auto py-2">Min Demand Reached</TableHead>
                                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">Remarks</TableHead>
                                        </TableRow>
                                        <TableRow>
                                            <TableHead className="text-center border bg-muted h-auto py-2">Amps</TableHead>
                                            <TableHead className="text-center border bg-muted h-auto py-2">MW</TableHead>
                                            <TableHead className="text-center border bg-muted h-auto py-2">Date</TableHead>
                                            <TableHead className="text-center border bg-muted h-auto py-2">Time</TableHead>
                                            <TableHead className="text-center border bg-muted h-auto py-2">Amps</TableHead>
                                            <TableHead className="text-center border bg-muted h-auto py-2">MW</TableHead>
                                            <TableHead className="text-center border bg-muted h-auto py-2">Date</TableHead>
                                            <TableHead className="text-center border bg-muted h-auto py-2">Time</TableHead>
                                        </TableRow>
                                    </TableHeader>
                                    <TableBody>
                                        {/* Main Feeders */}
                                        {p.main_feeders.map((row) => (
                                            <TableRow key={row.sl_no}>
                                                <TableCell className="text-center border p-2">{row.sl_no}</TableCell>
                                                <TableCell className="border p-2 whitespace-nowrap">{row.name}</TableCell>
                                                <TableCell className="text-center border p-2">{row.rating}</TableCell>
                                                <TableCell className="text-center border p-2">{row.max_amps}</TableCell>
                                                <TableCell className="text-center border p-2">{row.max_mw}</TableCell>
                                                <TableCell className="text-center border p-2 whitespace-nowrap">{row.max_mw_date}</TableCell>
                                                <TableCell className="text-center border p-2 whitespace-nowrap">{row.max_mw_time}</TableCell>
                                                <TableCell className="text-center border p-2">{row.min_amps}</TableCell>
                                                <TableCell className="text-center border p-2">{row.min_mw}</TableCell>
                                                <TableCell className="text-center border p-2 whitespace-nowrap">{row.min_mw_date}</TableCell>
                                                <TableCell className="text-center border p-2 whitespace-nowrap">{row.min_mw_time}</TableCell>
                                                <TableCell className="text-center border p-2"></TableCell>
                                            </TableRow>
                                        ))}
                                        
                                        {/* ICT Header */}
                                        <TableRow>
                                            <TableCell colSpan={12} className="text-center font-bold bg-muted p-2">ICT'S</TableCell>
                                        </TableRow>
                                        
                                        {/* ICT Feeders */}
                                        {p.ict_feeders.map((row) => (
                                            <TableRow key={row.sl_no}>
                                                <TableCell className="text-center border p-2">{row.sl_no}</TableCell>
                                                <TableCell className="border p-2 whitespace-nowrap">{row.name}</TableCell>
                                                <TableCell className="text-center border p-2">{row.rating}</TableCell>
                                                <TableCell className="text-center border p-2">{row.max_amps}</TableCell>
                                                <TableCell className="text-center border p-2">{row.max_mw}</TableCell>
                                                <TableCell className="text-center border p-2 whitespace-nowrap">{row.max_mw_date}</TableCell>
                                                <TableCell className="text-center border p-2 whitespace-nowrap">{row.max_mw_time}</TableCell>
                                                <TableCell className="text-center border p-2">{row.min_amps}</TableCell>
                                                <TableCell className="text-center border p-2">{row.min_mw}</TableCell>
                                                <TableCell className="text-center border p-2 whitespace-nowrap">{row.min_mw_date}</TableCell>
                                                <TableCell className="text-center border p-2 whitespace-nowrap">{row.min_mw_time}</TableCell>
                                                <TableCell className="text-center border p-2"></TableCell>
                                            </TableRow>
                                        ))}
                                    </TableBody>
                                </Table>
                            </div>
                            
                            {/* Station Load */}
                            {p.station_load && (
                                <div className="border rounded-md overflow-x-auto w-fit">
                                    <Table>
                                        <TableHeader>
                                            <TableRow>
                                                <TableHead className="text-center border bg-muted h-auto py-2">Station Load in MW</TableHead>
                                                <TableHead className="text-center border bg-muted h-auto py-2">Time</TableHead>
                                                <TableHead className="text-center border bg-muted h-auto py-2">Date</TableHead>
                                            </TableRow>
                                        </TableHeader>
                                        <TableBody>
                                            <TableRow>
                                                <TableCell className="text-center border p-2">{p.station_load.max_load}</TableCell>
                                                <TableCell className="text-center border p-2">{p.station_load.max_load_time}</TableCell>
                                                <TableCell className="text-center border p-2">{p.station_load.max_load_date}</TableCell>
                                            </TableRow>
                                        </TableBody>
                                    </Table>
                                </div>
                            )}
                        </TabsContent>
                    ))}
                </Tabs>
            </div>
        );
    }

    if (title === "Energy Consumption Daily Report") {
        const { sheets } = data || {};
        if (!sheets || sheets.length === 0) return <div className="text-center py-4">No data available</div>;

        // Move 33KV sheets to the bottom
        const otherSheets = sheets.filter(s => !s.name.toLowerCase().includes('33kv'));
        const kvSheets = sheets.filter(s => s.name.toLowerCase().includes('33kv'));
        const sortedSheets = [...otherSheets, ...kvSheets];

        return (
            <div className="space-y-8">
                {sortedSheets.map((sheet, index) => (
                    <div key={index} className="border rounded-md overflow-x-auto">
                        <h3 className="font-bold p-2 bg-muted border-b">{sheet.name}</h3>
                        <Table>
                            <TableHeader>
                                <TableRow>
                                    <TableHead className="text-center border bg-muted h-auto py-2">Sl. No.</TableHead>
                                    <TableHead className="text-center border bg-muted h-auto py-2">Meter Name</TableHead>
                                    <TableHead className="text-center border bg-muted h-auto py-2">Initial</TableHead>
                                    <TableHead className="text-center border bg-muted h-auto py-2">Final</TableHead>
                                    <TableHead className="text-center border bg-muted h-auto py-2">MF</TableHead>
                                    <TableHead className="text-center border bg-muted h-auto py-2">Consumption</TableHead>
                                    <TableHead className="text-center border bg-muted h-auto py-2">Unit</TableHead>
                                </TableRow>
                            </TableHeader>
                            <TableBody>
                                {sheet.meters.map((meter) => (
                                    <TableRow key={meter.sl_no}>
                                        <TableCell className="text-center border p-2">{meter.sl_no}</TableCell>
                                        <TableCell className="border p-2 whitespace-nowrap">{meter.name}</TableCell>
                                        <TableCell className="text-center border p-2">{meter.initial?.toFixed(2) ?? '-'}</TableCell>
                                        <TableCell className="text-center border p-2">{meter.final?.toFixed(2) ?? '-'}</TableCell>
                                        <TableCell className="text-center border p-2">{meter.mf}</TableCell>
                                        <TableCell className="text-center border p-2 font-medium">{meter.consumption?.toFixed(2) ?? '-'}</TableCell>
                                        <TableCell className="text-center border p-2">{meter.unit}</TableCell>
                                    </TableRow>
                                ))}
                            </TableBody>
                        </Table>
                    </div>
                ))}
            </div>
        );
    }

    if (title === "PTR Max–Min (Format-1)") {
        return (
            <div className="border rounded-md overflow-x-auto">
                <Table>
                    <TableHeader>
                        <TableRow>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">District Name</TableHead>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">Sub-Station Name</TableHead>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">PTR HV/LV in KV</TableHead>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">PTR Details/Rating in MVA</TableHead>
                            <TableHead colSpan={2} className="text-center border bg-muted h-auto py-2">General Loading</TableHead>
                            <TableHead colSpan={6} className="text-center border bg-muted h-auto py-2">Maximum Load Details</TableHead>
                            <TableHead colSpan={6} className="text-center border bg-muted h-auto py-2">Minimum Load Details</TableHead>
                        </TableRow>
                        <TableRow>
                            <TableHead className="text-center border bg-muted h-auto py-2">MW</TableHead>
                            <TableHead className="text-center border bg-muted h-auto py-2">MVAR</TableHead>
                            
                            <TableHead className="text-center border bg-muted h-auto py-2">Date</TableHead>
                            <TableHead className="text-center border bg-muted h-auto py-2">Time</TableHead>
                            <TableHead className="text-center border bg-muted h-auto py-2">MW</TableHead>
                            <TableHead className="text-center border bg-muted h-auto py-2">MVAR</TableHead>
                            <TableHead className="text-center border bg-muted h-auto py-2">MVA</TableHead>
                            <TableHead className="text-center border bg-muted h-auto py-2">Bus Voltage</TableHead>

                            <TableHead className="text-center border bg-muted h-auto py-2">Date</TableHead>
                            <TableHead className="text-center border bg-muted h-auto py-2">Time</TableHead>
                            <TableHead className="text-center border bg-muted h-auto py-2">MW</TableHead>
                            <TableHead className="text-center border bg-muted h-auto py-2">MVAR</TableHead>
                            <TableHead className="text-center border bg-muted h-auto py-2">MVA</TableHead>
                            <TableHead className="text-center border bg-muted h-auto py-2">Bus Voltage</TableHead>
                        </TableRow>
                    </TableHeader>
                    <TableBody>
                        {data && data.length > 0 ? (
                            data.map((row, index) => (
                                <TableRow key={index}>
                                    <TableCell className="text-center border p-2">{row.district}</TableCell>
                                    <TableCell className="text-center border p-2">{row.substation}</TableCell>
                                    <TableCell className="text-center border p-2">{row.ptr_kv}</TableCell>
                                    <TableCell className="text-center border p-2">{row.rating}</TableCell>
                                    
                                    <TableCell className="text-center border p-2">{row.general.mw?.toFixed(2)}</TableCell>
                                    <TableCell className="text-center border p-2">{row.general.mvar?.toFixed(2)}</TableCell>
                                    
                                    {/* Max Details */}
                                    <TableCell className="text-center border p-2 whitespace-nowrap">{row.max?.date ? new Date(row.max.date).toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '-'}</TableCell>
                                    <TableCell className="text-center border p-2">{row.max?.time || '-'}</TableCell>
                                    <TableCell className="text-center border p-2">{row.max?.mw?.toFixed(2)}</TableCell>
                                    <TableCell className="text-center border p-2">{row.max?.mvar?.toFixed(2)}</TableCell>
                                    <TableCell className="text-center border p-2">{row.max?.mva?.toFixed(2)}</TableCell>
                                    <TableCell className="text-center border p-2">-</TableCell>
                                    
                                    {/* Min Details */}
                                    <TableCell className="text-center border p-2 whitespace-nowrap">{row.min?.date ? new Date(row.min.date).toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '-'}</TableCell>
                                    <TableCell className="text-center border p-2">{row.min?.time || '-'}</TableCell>
                                    <TableCell className="text-center border p-2">{row.min?.mw?.toFixed(2)}</TableCell>
                                    <TableCell className="text-center border p-2">{row.min?.mvar?.toFixed(2)}</TableCell>
                                    <TableCell className="text-center border p-2">{row.min?.mva?.toFixed(2)}</TableCell>
                                    <TableCell className="text-center border p-2">-</TableCell>
                                </TableRow>
                            ))
                        ) : (
                            <TableRow>
                                <TableCell colSpan={18} className="text-center py-4">No data available</TableCell>
                            </TableRow>
                        )}
                    </TableBody>
                </Table>
            </div>
        );
    }

    if (title === "TL Max Loading (Format-4)") {
        // Calculate row spans for voltage to handle dynamic data robustly
        const voltageSpans = {};
        if (data && data.length > 0) {
            let currentVoltage = null;
            let startIndex = 0;
            data.forEach((row, index) => {
                if (row.voltage !== currentVoltage) {
                    if (currentVoltage !== null) {
                        voltageSpans[startIndex] = index - startIndex;
                    }
                    currentVoltage = row.voltage;
                    startIndex = index;
                }
                if (index === data.length - 1) {
                    voltageSpans[startIndex] = index - startIndex + 1;
                }
            });
        }

        return (
            <div className="border rounded-md overflow-x-auto">
                <Table>
                    <TableHeader>
                        <TableRow>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">Sl. No.</TableHead>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">District Name</TableHead>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">Voltage Level in KV</TableHead>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">Sub-Station Name</TableHead>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[250px]">Transmission Line Name</TableHead>
                            <TableHead colSpan={4} className="text-center border bg-muted h-auto py-2">MAXIMUM</TableHead>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">MD reached in 2026</TableHead>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">MD reached so far</TableHead>
                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2">Remarks</TableHead>
                        </TableRow>
                        <TableRow>
                            <TableHead className="text-center border bg-muted h-auto py-2">MW</TableHead>
                            <TableHead className="text-center border bg-muted h-auto py-2">MVAR</TableHead>
                            <TableHead className="text-center border bg-muted h-auto py-2">Date</TableHead>
                            <TableHead className="text-center border bg-muted h-auto py-2">Time</TableHead>
                        </TableRow>
                    </TableHeader>
                    <TableBody>
                        {data && data.length > 0 ? (
                            data.map((row, index) => (
                                <TableRow key={index}>
                                    <TableCell className="text-center border p-2">{row.sl_no}</TableCell>
                                    
                                    {/* District - Merged for all rows */}
                                    {index === 0 && (
                                        <TableCell rowSpan={data.length} className="text-center border p-2 bg-white align-middle">
                                            {row.district}
                                        </TableCell>
                                    )}

                                    {/* Voltage - Dynamic Merging */}
                                    {voltageSpans[index] && (
                                        <TableCell rowSpan={voltageSpans[index]} className="text-center border p-2 bg-white align-middle">
                                            {row.voltage}
                                        </TableCell>
                                    )}

                                    {/* Substation - Merged for all rows */}
                                    {index === 0 && (
                                        <TableCell rowSpan={data.length} className="text-center border p-2 bg-white align-middle">
                                            {row.substation}
                                        </TableCell>
                                    )}

                                    <TableCell className="text-center border p-2 whitespace-nowrap">{row.line_name}</TableCell>
                                    <TableCell className="text-center border p-2">{row.mw}</TableCell>
                                    <TableCell className="text-center border p-2">{row.mvar}</TableCell>
                                    <TableCell className="text-center border p-2 whitespace-nowrap">{row.date}</TableCell>
                                    <TableCell className="text-center border p-2">{row.time}</TableCell>
                                    <TableCell className="text-center border p-2">{row.md_2026}</TableCell>
                                    <TableCell className="text-center border p-2">{row.md_so_far}</TableCell>
                                    <TableCell className="text-center border p-2">{row.remarks}</TableCell>
                                </TableRow>
                            ))
                        ) : (
                            <TableRow>
                                <TableCell colSpan={12} className="text-center py-4">No data available</TableCell>
                            </TableRow>
                        )}
                    </TableBody>
                </Table>
            </div>
        );
    }

    if (title === "Interruptions") {
        const { sections } = data || {};
        const sectionList = Array.isArray(sections) ? sections : [];
        if (!sectionList.length) {
            return <div className="text-center py-4">No data available</div>;
        }

        return (
            <div className="w-full">
                <Tabs defaultValue={sectionList[0]?.id || "400kv"} className="w-full mt-2">
                    <TabsList className="grid w-full grid-cols-3 mb-4">
                        {sectionList.map((section) => (
                            <TabsTrigger key={section.id} value={section.id}>
                                {section.title}
                            </TabsTrigger>
                        ))}
                    </TabsList>

                    {sectionList.map((section) => (
                        <TabsContent key={section.id} value={section.id}>
                            <div className="border rounded-md overflow-x-auto">
                                <Table>
                                    <TableHeader>
                                        <TableRow>
                                            <TableHead colSpan={11} className="text-center border bg-muted h-auto py-2">
                                                {section.header}
                                            </TableHead>
                                        </TableRow>
                                        <TableRow>
                                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 w-[60px]">Sl. No</TableHead>
                                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[120px]">Date</TableHead>
                                            <TableHead colSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[140px]">Time</TableHead>
                                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[80px]">Duration</TableHead>
                                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[260px]">Cause of Interruption</TableHead>
                                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[260px]">Relay Indications / LC Work carried out</TableHead>
                                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[130px]">Break down declared or Not</TableHead>
                                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[150px]">Fault identified in patrolling</TableHead>
                                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[120px]">Fault Location</TableHead>
                                            <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[200px]">Remarks and action taken</TableHead>
                                        </TableRow>
                                        <TableRow>
                                            <TableHead className="text-center border bg-muted h-auto py-2 min-w-[70px]">From</TableHead>
                                            <TableHead className="text-center border bg-muted h-auto py-2 min-w-[70px]">To</TableHead>
                                        </TableRow>
                                    </TableHeader>
                                    <TableBody>
                                        {(() => {
                                            const groups = Array.isArray(section.groups) ? section.groups : [];
                                            if (!groups.length) {
                                                return (
                                                    <TableRow>
                                                        <TableCell colSpan={11} className="text-center py-4">
                                                            No interruptions found for this period.
                                                        </TableCell>
                                                    </TableRow>
                                                );
                                            }
                                            return groups.map((group) => (
                                                <React.Fragment key={group.name}>
                                                    <TableRow>
                                                        <TableCell colSpan={11} className="border font-semibold text-center">
                                                            {group.name}
                                                        </TableCell>
                                                    </TableRow>
                                                    {Array.isArray(group.rows) && group.rows.map((row, index) => (
                                                        <TableRow key={`${group.name}-${index}`}>
                                                            <TableCell className="text-center border p-2">
                                                                {row.sl_no ?? index + 1}
                                                            </TableCell>
                                                            <TableCell className="text-center border p-2 whitespace-nowrap">
                                                                {row.date}
                                                            </TableCell>
                                                            <TableCell className="text-center border p-2">
                                                                {row.time_from}
                                                            </TableCell>
                                                            <TableCell className="text-center border p-2">
                                                                {row.time_to}
                                                            </TableCell>
                                                            <TableCell className="text-center border p-2">
                                                                {row.duration}
                                                            </TableCell>
                                                            <TableCell className="border p-2 whitespace-pre-wrap">
                                                                {row.cause}
                                                            </TableCell>
                                                            <TableCell className="border p-2 whitespace-pre-wrap">
                                                                {row.relay}
                                                            </TableCell>
                                                            <TableCell className="text-center border p-2">
                                                                {row.breakdown}
                                                            </TableCell>
                                                            <TableCell className="text-center border p-2">
                                                                {row.fault_identified}
                                                            </TableCell>
                                                            <TableCell className="text-center border p-2">
                                                                {row.fault_location}
                                                            </TableCell>
                                                            <TableCell className="border p-2 whitespace-pre-wrap">
                                                                {row.remarks}
                                                            </TableCell>
                                                        </TableRow>
                                                    ))}
                                                </React.Fragment>
                                            ));
                                        })()}
                                    </TableBody>
                                </Table>
                            </div>
                        </TabsContent>
                    ))}
                </Tabs>
            </div>
        );
    }

    if (reportId === "mis-interruptions") {
        const rows = Array.isArray(data?.rows) ? data.rows : [];

        const monthLabels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        const monthLabel = month >= 1 && month <= 12 ? monthLabels[month - 1] : month;
        const headerLine1 = "STATEMENT-16";
        const headerLine2 = `PARTICULARS OF INTERRUPTIONS FOR THE MONTH OF ${monthLabel} - ${year} IN O & M-I DIVISION`;

        return (
            <div className="w-full space-y-2 mt-2">
                <div className="text-center font-semibold text-sm">{headerLine1}</div>
                <div className="text-center font-semibold text-sm">{headerLine2}</div>
                <div className="border rounded-md overflow-x-auto">
                    <Table>
                        <TableHeader>
                            <TableRow>
                                <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 w-[60px]">
                                    Sl. No
                                </TableHead>
                                <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[200px]">
                                    Name of the feeder
                                </TableHead>
                                <TableHead colSpan={4} className="text-center border bg-muted h-auto py-2 min-w-[240px]">
                                    Faulty Trippings
                                </TableHead>
                                <TableHead colSpan={2} className="text-center border bg-muted h-auto py-2">
                                    Incoming Supply
                                </TableHead>
                                <TableHead colSpan={2} className="text-center border bg-muted h-auto py-2">
                                    Load Relief
                                </TableHead>
                                <TableHead colSpan={2} className="text-center border bg-muted h-auto py-2">
                                    Equipment Failures
                                </TableHead>
                                <TableHead colSpan={2} className="text-center border bg-muted h-auto py-2">
                                    Break Downs
                                </TableHead>
                                <TableHead colSpan={2} className="text-center border bg-muted h-auto py-2">
                                    Pre Arranged
                                </TableHead>
                                <TableHead colSpan={2} className="text-center border bg-muted h-auto py-2">
                                    LC/NBFC
                                </TableHead>
                                <TableHead rowSpan={2} className="text-center border bg-muted h-auto py-2 min-w-[160px]">
                                    Remarks
                                </TableHead>
                            </TableRow>
                            <TableRow>
                                <TableHead className="text-center border bg-muted h-auto py-2">O/L</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">E/L</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Others</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Duration</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">No</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Duration</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">No</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Duration</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">No</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Duration</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">No</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Duration</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">No</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Duration</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">No</TableHead>
                                <TableHead className="text-center border bg-muted h-auto py-2">Duration</TableHead>
                            </TableRow>
                        </TableHeader>
                        <TableBody>
                            {rows && rows.length > 0 ? (
                                rows.map((row, index) => (
                                    <TableRow key={row.feeder_name || index}>
                                        <TableCell className="text-center border p-2">
                                            {row.sl_no ?? index + 1}
                                        </TableCell>
                                        <TableCell className="border p-2 whitespace-pre-wrap">
                                            {row.feeder_name}
                                        </TableCell>
                                        <TableCell className="text-center border p-2">
                                            {row.faulty_ol_count || ""}
                                        </TableCell>
                                        <TableCell className="text-center border p-2">
                                            {row.faulty_el_count || ""}
                                        </TableCell>
                                        <TableCell className="text-center border p-2">
                                            {row.faulty_other_count || ""}
                                        </TableCell>
                                        <TableCell className="text-center border p-2">
                                            {row.faulty_duration || ""}
                                        </TableCell>
                                        <TableCell className="text-center border p-2">
                                            {row.incoming_count || ""}
                                        </TableCell>
                                        <TableCell className="text-center border p-2">
                                            {row.incoming_duration || ""}
                                        </TableCell>
                                        <TableCell className="text-center border p-2">
                                            {row.load_relief_count || ""}
                                        </TableCell>
                                        <TableCell className="text-center border p-2">
                                            {row.load_relief_duration || ""}
                                        </TableCell>
                                        <TableCell className="text-center border p-2">
                                            {row.equipment_failures_count || ""}
                                        </TableCell>
                                        <TableCell className="text-center border p-2">
                                            {row.equipment_failures_duration || ""}
                                        </TableCell>
                                        <TableCell className="text-center border p-2">
                                            {row.breakdown_count || ""}
                                        </TableCell>
                                        <TableCell className="text-center border p-2">
                                            {row.breakdown_duration || ""}
                                        </TableCell>
                                        <TableCell className="text-center border p-2">
                                            {row.pre_arranged_count || ""}
                                        </TableCell>
                                        <TableCell className="text-center border p-2">
                                            {row.pre_arranged_duration || ""}
                                        </TableCell>
                                        <TableCell className="text-center border p-2">
                                            {row.lc_nbfc_count || ""}
                                        </TableCell>
                                        <TableCell className="text-center border p-2">
                                            {row.lc_nbfc_duration || ""}
                                        </TableCell>
                                        <TableCell className="border p-2 whitespace-pre-wrap">
                                            {row.remarks || ""}
                                        </TableCell>
                                    </TableRow>
                                ))
                            ) : (
                                <TableRow>
                                    <TableCell colSpan={19} className="text-center py-4">
                                        No data available for this period.
                                    </TableCell>
                                </TableRow>
                            )}
                        </TableBody>
                    </Table>
                </div>
            </div>
        );
    }

    if (title === "Line Losses" && reportId === "line-losses") {
        const rows = Array.isArray(data) ? data : data?.lines || [];

        return (
            <div className="border rounded-md max-w-full overflow-x-auto">
                <Table className="w-full table-fixed text-xs min-w-[2000px]">
                    <TableHeader>
                        <TableRow>
                            <TableHead rowSpan={3} className="text-center border bg-muted h-auto py-1 w-[3%]">Sl. No.</TableHead>
                            <TableHead rowSpan={3} className="text-center border bg-muted h-auto py-1 w-[10%] break-words">Name of the Feeder</TableHead>
                            <TableHead colSpan={8} className="text-center border bg-muted h-auto py-1">Shankarapally End</TableHead>
                            <TableHead colSpan={8} className="text-center border bg-muted h-auto py-1">Other End</TableHead>
                            <TableHead rowSpan={3} className="text-center border bg-muted h-auto py-1 w-[6%]">% Losses</TableHead>
                            <TableHead rowSpan={3} className="text-center border bg-muted h-auto py-1 w-[8%]">Remarks</TableHead>
                        </TableRow>
                        <TableRow>
                            <TableHead colSpan={4} className="text-center border bg-muted h-auto py-1">Import</TableHead>
                            <TableHead colSpan={4} className="text-center border bg-muted h-auto py-1">Export</TableHead>
                            <TableHead colSpan={4} className="text-center border bg-muted h-auto py-1">Import</TableHead>
                            <TableHead colSpan={4} className="text-center border bg-muted h-auto py-1">Export</TableHead>
                        </TableRow>
                        <TableRow>
                            {[1, 2, 3, 4].map((_, groupIdx) => (
                                <React.Fragment key={groupIdx}>
                                    <TableHead className="text-center border bg-muted h-auto py-1">Initial</TableHead>
                                    <TableHead className="text-center border bg-yellow-100 h-auto py-1">Final</TableHead>
                                    <TableHead className="text-center border bg-muted h-auto py-1">MF</TableHead>
                                    <TableHead className="text-center border bg-muted h-auto py-1">Cons.</TableHead>
                                </React.Fragment>
                            ))}
                        </TableRow>
                    </TableHeader>
                    <TableBody>
                        {rows && rows.length > 0 ? (
                            rows.map((row, index) => (
                                <TableRow key={index}>
                                    <TableCell className="text-center border px-2 py-2">
                                        {row.sl_no ?? index + 1}
                                    </TableCell>
                                    <TableCell className="border px-2 py-2 break-words">
                                        {row.feeder_name}
                                    </TableCell>

                                    <TableCell className="text-right border px-2 py-2 whitespace-nowrap">
                                        {row.shankarpally?.import?.initial}
                                    </TableCell>
                                    <TableCell className="text-right border px-2 py-2 bg-yellow-50 whitespace-nowrap">
                                        {row.shankarpally?.import?.final}
                                    </TableCell>
                                    <TableCell className="text-right border px-2 py-2 whitespace-nowrap">
                                        {row.shankarpally?.import?.mf}
                                    </TableCell>
                                    <TableCell className="text-right border px-2 py-2 font-medium whitespace-nowrap">
                                        {row.shankarpally?.import?.consumption != null
                                            ? row.shankarpally.import.consumption.toFixed(2)
                                            : ""}
                                    </TableCell>

                                    <TableCell className="text-right border px-2 py-2 whitespace-nowrap">
                                        {row.shankarpally?.export?.initial}
                                    </TableCell>
                                    <TableCell className="text-right border px-2 py-2 bg-yellow-50 whitespace-nowrap">
                                        {row.shankarpally?.export?.final}
                                    </TableCell>
                                    <TableCell className="text-right border px-2 py-2 whitespace-nowrap">
                                        {row.shankarpally?.export?.mf}
                                    </TableCell>
                                    <TableCell className="text-right border px-2 py-2 font-medium whitespace-nowrap">
                                        {row.shankarpally?.export?.consumption != null
                                            ? row.shankarpally.export.consumption.toFixed(2)
                                            : ""}
                                    </TableCell>

                                    <TableCell className="text-right border px-2 py-2 whitespace-nowrap">
                                        {row.other_end?.import?.initial}
                                    </TableCell>
                                    <TableCell className="text-right border px-2 py-2 bg-yellow-50 whitespace-nowrap">
                                        {row.other_end?.import?.final}
                                    </TableCell>
                                    <TableCell className="text-right border px-2 py-2 whitespace-nowrap">
                                        {row.other_end?.import?.mf}
                                    </TableCell>
                                    <TableCell className="text-right border px-2 py-2 font-medium whitespace-nowrap">
                                        {row.other_end?.import?.consumption != null
                                            ? row.other_end.import.consumption.toFixed(2)
                                            : ""}
                                    </TableCell>

                                    <TableCell className="text-right border px-2 py-2 whitespace-nowrap">
                                        {row.other_end?.export?.initial}
                                    </TableCell>
                                    <TableCell className="text-right border px-2 py-2 bg-yellow-50 whitespace-nowrap">
                                        {row.other_end?.export?.final}
                                    </TableCell>
                                    <TableCell className="text-right border px-2 py-2 whitespace-nowrap">
                                        {row.other_end?.export?.mf}
                                    </TableCell>
                                    <TableCell className="text-right border px-2 py-2 font-medium whitespace-nowrap">
                                        {row.other_end?.export?.consumption != null
                                            ? row.other_end.export.consumption.toFixed(2)
                                            : ""}
                                    </TableCell>

                                    <TableCell className="text-right border px-2 py-2 font-bold whitespace-nowrap">
                                        {row.stats?.pct_loss != null
                                            ? row.stats.pct_loss.toFixed(2)
                                            : ""}
                                    </TableCell>
                                    <TableCell className="border px-2 py-2 break-words">
                                        {row.remarks || ""}
                                    </TableCell>
                                </TableRow>
                            ))
                        ) : (
                            <TableRow>
                                <TableCell colSpan={20} className="text-center py-4">
                                    No data available
                                </TableCell>
                            </TableRow>
                        )}
                    </TableBody>
                </Table>
            </div>
        );
    }

    if (title === "Line Losses") {
        return (
            <div className="border rounded-md max-w-full overflow-x-auto">
                <Table className="w-full table-fixed text-xs min-w-[1200px]">
                    <TableHeader>
                        <TableRow>
                            <TableHead rowSpan={3} className="text-center border bg-muted h-auto py-1 w-[18%] break-words">Name of the Feeder</TableHead>
                            <TableHead colSpan={6} className="text-center border bg-muted h-auto py-1 w-[38%]">Shankarapally End</TableHead>
                            <TableHead colSpan={6} className="text-center border bg-muted h-auto py-1 w-[38%]">Other End</TableHead>
                            <TableHead rowSpan={3} className="text-center border bg-muted h-auto py-1 w-[6%]">%</TableHead>
                        </TableRow>
                        <TableRow>
                            <TableHead colSpan={3} className="text-center border bg-muted h-auto py-1">Import</TableHead>
                            <TableHead colSpan={3} className="text-center border bg-muted h-auto py-1">Export</TableHead>
                            <TableHead colSpan={3} className="text-center border bg-muted h-auto py-1">Import</TableHead>
                            <TableHead colSpan={3} className="text-center border bg-muted h-auto py-1">Export</TableHead>
                        </TableRow>
                        <TableRow>
                            {[1, 2, 3, 4].map((_, groupIdx) => (
                                <React.Fragment key={groupIdx}>
                                    <TableHead className="text-center border bg-muted h-auto py-1">Initial</TableHead>
                                    <TableHead className="text-center border bg-yellow-100 h-auto py-1">Final</TableHead>
                                    <TableHead className="text-center border bg-muted h-auto py-1">Cons.</TableHead>
                                </React.Fragment>
                            ))}
                        </TableRow>
                    </TableHeader>
                    <TableBody>
                        {data && data.length > 0 ? (
                            data.map((row, index) => (
                                <TableRow key={index}>
                                    <TableCell className="border px-2 py-2 break-words">{row.feeder_name}</TableCell>
                                    
                                    {/* Shankarpally Import */}
                                    <TableCell className="text-right border px-2 py-2 whitespace-nowrap">{row.shankarpally.import.initial}</TableCell>
                                    <TableCell className="text-right border px-2 py-2 bg-yellow-50 whitespace-nowrap">{row.shankarpally.import.final}</TableCell>
                                    <TableCell className="text-right border px-2 py-2 font-medium whitespace-nowrap">{row.shankarpally.import.consumption?.toFixed(2)}</TableCell>
                                    
                                    {/* Shankarpally Export */}
                                    <TableCell className="text-right border px-2 py-2 whitespace-nowrap">{row.shankarpally.export.initial}</TableCell>
                                    <TableCell className="text-right border px-2 py-2 bg-yellow-50 whitespace-nowrap">{row.shankarpally.export.final}</TableCell>
                                    <TableCell className="text-right border px-2 py-2 font-medium whitespace-nowrap">{row.shankarpally.export.consumption?.toFixed(2)}</TableCell>
                                    
                                    {/* Other End Import */}
                                    <TableCell className="text-right border px-2 py-2 whitespace-nowrap">{row.other_end.import.initial}</TableCell>
                                    <TableCell className="text-right border px-2 py-2 bg-yellow-50 whitespace-nowrap">{row.other_end.import.final}</TableCell>
                                    <TableCell className="text-right border px-2 py-2 font-medium whitespace-nowrap">{row.other_end.import.consumption?.toFixed(2)}</TableCell>
                                    
                                    {/* Other End Export */}
                                    <TableCell className="text-right border px-2 py-2 whitespace-nowrap">{row.other_end.export.initial}</TableCell>
                                    <TableCell className="text-right border px-2 py-2 bg-yellow-50 whitespace-nowrap">{row.other_end.export.final}</TableCell>
                                    <TableCell className="text-right border px-2 py-2 font-medium whitespace-nowrap">{row.other_end.export.consumption?.toFixed(2)}</TableCell>
                                    
                                    <TableCell className="text-right border px-2 py-2 font-bold whitespace-nowrap">{row.stats.pct_loss?.toFixed(2)}</TableCell>
                                </TableRow>
                            ))
                        ) : (
                            <TableRow>
                                <TableCell colSpan={14} className="text-center py-4">No data available</TableCell>
                            </TableRow>
                        )}
                    </TableBody>
                </Table>
            </div>
        );
    }

    return <div>Preview not available for this report type.</div>;
  };

  return (
    <Dialog open={isOpen} onOpenChange={onClose}>
      <DialogContent className="max-w-7xl max-h-[85vh] overflow-y-auto">
        <DialogHeader>
          <div className="flex items-center justify-between">
            <DialogTitle>
              {title} 
              {date ? ` - ${date.toLocaleDateString('en-GB', { day: '2-digit', month: 'long', year: 'numeric' })}` : 
                (subtitle ? ` - ${subtitle}` : (month && year ? ` - ${month}/${year}` : ''))}
            </DialogTitle>
            
            {onPrev && onNext && (
              <div className="flex items-center gap-2 mr-8">
                <Button 
                  variant="outline" 
                  size="icon" 
                  onClick={onPrev}
                  title="Previous Day"
                  disabled={loading}
                >
                  <ChevronLeft className="w-4 h-4" />
                </Button>
                <Button 
                  variant="outline" 
                  size="icon" 
                  onClick={onNext}
                  disabled={!hasNext || loading}
                  title="Next Day"
                >
                  <ChevronRight className="w-4 h-4" />
                </Button>
              </div>
            )}
          </div>
        </DialogHeader>
        {renderContent()}
      </DialogContent>
    </Dialog>
  );
}
