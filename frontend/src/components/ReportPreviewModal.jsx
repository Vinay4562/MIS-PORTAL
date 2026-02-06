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
import { Loader2 } from "lucide-react";

export function ReportPreviewModal({ isOpen, onClose, title, data, loading, year, month }) {
  
  const renderContent = () => {
    if (loading) {
      return (
        <div className="flex justify-center items-center py-10">
          <Loader2 className="w-8 h-8 animate-spin" />
        </div>
      );
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

    if (title === "Fortnight") {
        if (!data || !data.periods) return <div className="text-center py-4">No data available</div>;
        
        return (
            <div className="w-full">
                <Tabs defaultValue="1-15" className="w-full">
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


    if (title === "Line Losses") {
        return (
            <div className="border rounded-md overflow-x-auto">
                <Table>
                    <TableHeader>
                        <TableRow>
                            <TableHead rowSpan={3} className="text-center border bg-muted h-auto py-2">Sl. No.</TableHead>
                            <TableHead rowSpan={3} className="text-center border bg-muted h-auto py-2 min-w-[200px]">Name of the Feeder</TableHead>
                            <TableHead colSpan={8} className="text-center border bg-muted h-auto py-2">Shankarapally End</TableHead>
                            <TableHead colSpan={8} className="text-center border bg-muted h-auto py-2">Other End</TableHead>
                            <TableHead rowSpan={3} className="text-center border bg-muted h-auto py-2">% Losses</TableHead>
                            <TableHead rowSpan={3} className="text-center border bg-muted h-auto py-2">Remarks</TableHead>
                        </TableRow>
                        <TableRow>
                            <TableHead colSpan={4} className="text-center border bg-muted h-auto py-2">Import</TableHead>
                            <TableHead colSpan={4} className="text-center border bg-muted h-auto py-2">Export</TableHead>
                            <TableHead colSpan={4} className="text-center border bg-muted h-auto py-2">Import</TableHead>
                            <TableHead colSpan={4} className="text-center border bg-muted h-auto py-2">Export</TableHead>
                        </TableRow>
                        <TableRow>
                            {[1, 2, 3, 4].map((_, groupIdx) => (
                                <React.Fragment key={groupIdx}>
                                    <TableHead className="text-center border bg-muted h-auto py-2 min-w-[80px]">Initial</TableHead>
                                    <TableHead className="text-center border bg-muted h-auto py-2 min-w-[80px]">Final</TableHead>
                                    <TableHead className="text-center border bg-muted h-auto py-2 min-w-[50px]">MF</TableHead>
                                    <TableHead className="text-center border bg-muted h-auto py-2 min-w-[80px]">Cons.</TableHead>
                                </React.Fragment>
                            ))}
                        </TableRow>
                    </TableHeader>
                    <TableBody>
                        {data && data.length > 0 ? (
                            data.map((row, index) => (
                                <TableRow key={index}>
                                    <TableCell className="text-center border p-2">{row.sl_no}</TableCell>
                                    <TableCell className="border p-2 whitespace-nowrap">{row.feeder_name}</TableCell>
                                    
                                    {/* Shankarpally Import */}
                                    <TableCell className="text-center border p-2">{row.shankarpally.import.initial}</TableCell>
                                    <TableCell className="text-center border p-2">{row.shankarpally.import.final}</TableCell>
                                    <TableCell className="text-center border p-2">{row.shankarpally.import.mf}</TableCell>
                                    <TableCell className="text-center border p-2 font-medium">{row.shankarpally.import.consumption?.toFixed(2)}</TableCell>
                                    
                                    {/* Shankarpally Export */}
                                    <TableCell className="text-center border p-2">{row.shankarpally.export.initial}</TableCell>
                                    <TableCell className="text-center border p-2">{row.shankarpally.export.final}</TableCell>
                                    <TableCell className="text-center border p-2">{row.shankarpally.export.mf}</TableCell>
                                    <TableCell className="text-center border p-2 font-medium">{row.shankarpally.export.consumption?.toFixed(2)}</TableCell>
                                    
                                    {/* Other End Import */}
                                    <TableCell className="text-center border p-2">{row.other_end.import.initial}</TableCell>
                                    <TableCell className="text-center border p-2">{row.other_end.import.final}</TableCell>
                                    <TableCell className="text-center border p-2">{row.other_end.import.mf}</TableCell>
                                    <TableCell className="text-center border p-2 font-medium">{row.other_end.import.consumption?.toFixed(2)}</TableCell>
                                    
                                    {/* Other End Export */}
                                    <TableCell className="text-center border p-2">{row.other_end.export.initial}</TableCell>
                                    <TableCell className="text-center border p-2">{row.other_end.export.final}</TableCell>
                                    <TableCell className="text-center border p-2">{row.other_end.export.mf}</TableCell>
                                    <TableCell className="text-center border p-2 font-medium">{row.other_end.export.consumption?.toFixed(2)}</TableCell>
                                    
                                    <TableCell className="text-center border p-2 font-bold">{row.stats.pct_loss?.toFixed(2)}</TableCell>
                                    <TableCell className="text-center border p-2">-</TableCell>
                                </TableRow>
                            ))
                        ) : (
                            <TableRow>
                                <TableCell colSpan={20} className="text-center py-4">No data available</TableCell>
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
          <DialogTitle>{title} - {month}/{year}</DialogTitle>
        </DialogHeader>
        {renderContent()}
      </DialogContent>
    </Dialog>
  );
}
