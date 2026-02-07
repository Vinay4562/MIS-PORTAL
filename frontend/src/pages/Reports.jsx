import React, { useState } from 'react';
import axios from 'axios';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter, DialogDescription } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Checkbox } from "@/components/ui/checkbox";
import { FileText, Eye, Download, Calendar, Mail } from 'lucide-react';
import { toast } from 'sonner';
import { ReportPreviewModal } from '@/components/ReportPreviewModal';
import { Loader } from '@/components/ui/loader';

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

const REPORTS = [
  { id: 'boundary-meter', title: 'Boundary Meter Reading (33KV)', description: '33KV Boundary Meter Reading Report' },
  { id: 'fortnight', title: 'Fortnight', description: 'Fortnightly Report' },
  { id: 'daily-max-mva', title: 'Daily Max MVA (SAP)', description: 'Daily Max MVA Report from SAP' },
  { id: 'kpi', title: 'KPI', description: 'Key Performance Indicators Report' },
  { id: 'line-losses', title: 'Line Losses', description: 'Line Losses Report' },
  { id: 'ptr-max-min', title: 'PTR Max–Min (Format-1)', description: 'PTR Max-Min Data Format-1' },
  { id: 'tl-max-loading', title: 'TL Max Loading (Format-4)', description: 'Transmission Line Max Loading Format-4' },
];

export default function Reports() {
  const [year, setYear] = useState(new Date().getFullYear());
  const [month, setMonth] = useState(new Date().getMonth() + 1);
  const [showDateSelector, setShowDateSelector] = useState(true);
  const [loading, setLoading] = useState(null); // stores report.id if loading
  
  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];

  // Preview State
  const [previewOpen, setPreviewOpen] = useState(false);
  const [previewData, setPreviewData] = useState(null);
  const [previewLoading, setPreviewLoading] = useState(false);
  const [previewReport, setPreviewReport] = useState(null);

  // Email State
  const [emailDialogOpen, setEmailDialogOpen] = useState(false);
  const [recipientEmail, setRecipientEmail] = useState('');
  const [sendingEmail, setSendingEmail] = useState(false);
  const [selectedReportIds, setSelectedReportIds] = useState(REPORTS.map(r => r.id));

  const handleSendEmail = async () => {
    if (!recipientEmail) {
        toast.error("Please enter an email address");
        return;
    }
    
    // Basic email validation
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(recipientEmail)) {
        toast.error("Please enter a valid email address");
        return;
    }

    if (selectedReportIds.length === 0) {
        toast.error("Please select at least one report to send");
        return;
    }

    setSendingEmail(true);
    try {
        const token = localStorage.getItem('token');
        await axios.post(`${API}/reports/send-mail`, {
            email: recipientEmail,
            year,
            month,
            report_ids: selectedReportIds
        }, {
            headers: { Authorization: `Bearer ${token}` }
        });
        toast.success(`Reports sent successfully to ${recipientEmail}`);
        setEmailDialogOpen(false);
        setRecipientEmail('');
        // Reset selection to all for next time, or keep it? 
        // Let's reset to all as it's the default behavior
        setSelectedReportIds(REPORTS.map(r => r.id));
    } catch (error) {
        console.error("Failed to send email:", error);
        toast.error(error.response?.data?.detail || "Failed to send email");
    } finally {
        setSendingEmail(false);
    }
  };

  const handleSubmitDateSelection = () => {
    setShowDateSelector(false);
  };

  const handlePreview = async (report) => {
    if (['boundary-meter', 'daily-max-mva', 'kpi', 'line-losses', 'ptr-max-min', 'tl-max-loading', 'fortnight'].includes(report.id)) {
        setPreviewReport(report);
        setPreviewOpen(true);
        setPreviewLoading(true);
        setPreviewData(null);
        
        try {
            let endpoint = '';
            if (report.id === 'boundary-meter') {
                endpoint = `/reports/boundary-meter-33kv/data/${year}/${month}`;
            } else if (report.id === 'fortnight') {
                endpoint = `/reports/fortnight/preview/${year}/${month}`;
            } else if (report.id === 'daily-max-mva') {
                endpoint = `/reports/daily-max-mva/preview/${year}/${month}`;
            } else if (report.id === 'kpi') {
                endpoint = `/reports/kpi/preview/${year}/${month}`;
            } else if (report.id === 'line-losses') {
                endpoint = `/reports/line-losses/preview/${year}/${month}`;
            } else if (report.id === 'ptr-max-min') {
                endpoint = `/reports/ptr-max-min-format1/preview/${year}/${month}`;
            } else if (report.id === 'tl-max-loading') {
                endpoint = `/reports/tl-max-loading-format4/preview/${year}/${month}`;
            }

            const response = await axios.get(`${API}${endpoint}`);
            setPreviewData(response.data);
        } catch (error) {
            console.error('Preview error:', error);
            toast.error(error.response?.data?.detail || "Failed to load preview data");
            setPreviewData(null);
        } finally {
            setPreviewLoading(false);
        }
    } else {
        toast.info(`Preview for ${report.title} is coming soon`);
    }
  };

  const downloadReport = async (report, silent = false) => {
    // Define endpoints for supported reports
    const endpoints = {
        'boundary-meter': `/reports/boundary-meter-33kv/${year}/${month}`,
        'fortnight': `/reports/fortnight/${year}/${month}`,
        'daily-max-mva': `/reports/daily-max-mva/export/${year}/${month}`,
        'kpi': `/reports/kpi/export/${year}/${month}`,
        'line-losses': `/reports/line-losses/export/${year}/${month}`,
        'ptr-max-min': `/reports/ptr-max-min-format1/export/${year}/${month}`,
        'tl-max-loading': `/reports/tl-max-loading-format4/export/${year}/${month}`
    };

    if (!endpoints[report.id]) {
      if (!silent) toast.info(`Export for ${report.title} is coming soon`);
      return false;
    }

    try {
        if (!silent) setLoading(report.id);
        if (!silent) toast.info(`Generating ${report.title}...`);
        
        const response = await axios.get(`${API}${endpoints[report.id]}`, {
          responseType: 'blob'
        });
        
        // Extract filename from header or generate default
        const contentDisposition = response.headers['content-disposition'];
        let filename = `${report.title.replace(/\s+/g, '_')}_${month}_${year}.xlsx`;
        if (contentDisposition) {
            const filenameMatch = contentDisposition.match(/filename="?([^"]+)"?/);
            if (filenameMatch && filenameMatch.length === 2)
                filename = filenameMatch[1];
        }

        const url = window.URL.createObjectURL(new Blob([response.data]));
        const link = document.createElement('a');
        link.href = url;
        link.setAttribute('download', filename);
        document.body.appendChild(link);
        link.click();
        link.remove();
        
        if (!silent) toast.success("Report downloaded successfully");
        return true;
      } catch (error) {
        console.error('Export error:', error);
        if (!silent) toast.error(error.response?.data?.detail ? JSON.parse(await error.response.data.text()).detail : "Failed to export report");
        return false;
      } finally {
        if (!silent) setLoading(null);
      }
  };

  const handleExport = (report) => downloadReport(report, false);

  const handleDownloadAll = async () => {
    setLoading('all');
    toast.info("Starting bulk download...");
    let successCount = 0;
    
    // Filter reports that have endpoints implemented
    // We can know this by checking the endpoints object in downloadReport, 
    // but since it's local, we'll just try to download all known supported ones.
    // Or better, we replicate the supported check here or just iterate all.
    // Since downloadReport checks for support, we can just iterate.
    
    for (const report of REPORTS) {
        // Add a small delay to prevent browser blocking
        await new Promise(resolve => setTimeout(resolve, 500));
        const success = await downloadReport(report, true);
        if (success) successCount++;
    }
    
    if (successCount > 0) {
        toast.success(`Downloaded ${successCount} reports.`);
    } else {
        toast.warning("No reports were downloaded.");
    }
    setLoading(null);
  };

  if (showDateSelector) {
    return (
      <div className="flex items-center justify-center" style={{ minHeight: 'calc(100vh - 200px)' }}>
        <Card className="w-full max-w-md shadow-lg">
          <CardHeader>
            <CardTitle className="text-2xl font-heading flex items-center gap-2">
              <Calendar className="w-6 h-6" />
              Select Period
            </CardTitle>
          </CardHeader>
          <CardContent className="space-y-6">
            <div className="space-y-2">
              <label className="text-sm font-medium">Year</label>
              <Select value={year.toString()} onValueChange={(v) => setYear(parseInt(v))}>
                <SelectTrigger>
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  {Array.from({ length: 2030 - 2018 + 1 }, (_, i) => 2018 + i).map((y) => (
                    <SelectItem key={y} value={y.toString()}>
                      {y}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            <div className="space-y-2">
              <label className="text-sm font-medium">Month</label>
              <Select value={month.toString()} onValueChange={(v) => setMonth(parseInt(v))}>
                <SelectTrigger>
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  {monthNames.map((name, index) => (
                    <SelectItem key={index + 1} value={(index + 1).toString()}>
                      {name}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            <Button 
              onClick={handleSubmitDateSelection} 
              className="w-full" 
              size="lg"
            >
              Show Reports
            </Button>
          </CardContent>
        </Card>
      </div>
    );
  }

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h1 className="text-2xl font-bold text-slate-900 dark:text-slate-100">Reports</h1>
          <p className="text-slate-500 dark:text-slate-400">
            {monthNames[month - 1]} {year}
          </p>
        </div>

        <div className="flex items-center gap-2">
            <Button 
                variant="outline" 
                onClick={handleDownloadAll}
                className="gap-2"
                disabled={loading === 'all'}
            >
                {loading === 'all' ? <Loader size="sm" /> : <Download className="w-4 h-4" />}
                Download All
            </Button>
            <Button 
                variant="outline" 
                onClick={() => setShowDateSelector(true)}
                className="gap-2"
            >
                <Calendar className="w-4 h-4" />
            Period
        </Button>
        <Button 
            variant="outline" 
            onClick={() => setEmailDialogOpen(true)}
            className="gap-2"
        >
            <Mail className="w-4 h-4" />
            Send Mail
        </Button>
    </div>
  </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6">
        {REPORTS.map((report) => (
          <Card key={report.id} className="hover:shadow-md transition-shadow">
            <CardHeader className="pb-3">
              <div className="flex items-start justify-between">
                <div className="p-2 bg-blue-50 dark:bg-blue-900/20 rounded-lg">
                  <FileText className="w-6 h-6 text-blue-600 dark:text-blue-400" />
                </div>
              </div>
              <CardTitle className="mt-4 text-lg">{report.title}</CardTitle>
              <CardDescription>{report.description}</CardDescription>
            </CardHeader>
            <CardContent>
              <div className="flex gap-2 mt-2">
                <Button 
                  variant="outline" 
                  className="flex-1 gap-2"
                  onClick={() => handlePreview(report)}
                >
                  <Eye className="w-4 h-4" />
                  Preview
                </Button>
                <Button 
                  variant="default" 
                  className="flex-1 gap-2"
                  onClick={() => handleExport(report)}
                  disabled={loading === report.id}
                >
                  {loading === report.id ? <Loader size="sm" className="animate-spin text-white" /> : <Download className="w-4 h-4" />}
                  Export
                </Button>
              </div>
            </CardContent>
          </Card>
        ))}
      </div>
      
      {previewReport && (
        <ReportPreviewModal 
            isOpen={previewOpen}
            onClose={() => setPreviewOpen(false)}
            title={previewReport.title}
            data={previewData}
            loading={previewLoading}
            year={year}
            month={month}
        />
      )}

      <Dialog open={emailDialogOpen} onOpenChange={setEmailDialogOpen}>
        <DialogContent className="max-w-md">
            <DialogHeader>
                <DialogTitle>Send Reports via Email</DialogTitle>
                <DialogDescription>
                    Select reports and enter the recipient's email address. All reports for {monthNames[month-1]} {year} are available.
                </DialogDescription>
            </DialogHeader>
            <div className="grid gap-4 py-4">
                <div className="grid gap-2">
                    <Label htmlFor="email">
                        Email Address
                    </Label>
                    <Input
                        id="email"
                        type="email"
                        value={recipientEmail}
                        onChange={(e) => setRecipientEmail(e.target.value)}
                        placeholder="recipient@example.com"
                    />
                </div>
                
                <div className="space-y-3">
                    <div className="flex items-center justify-between">
                        <Label>Select Reports</Label>
                        <div className="flex gap-2">
                             <Button 
                                variant="ghost" 
                                size="sm" 
                                className="h-auto p-0 text-xs text-blue-600 hover:text-blue-800 hover:bg-transparent"
                                onClick={() => setSelectedReportIds(REPORTS.map(r => r.id))}
                             >
                                Select All
                             </Button>
                             <span className="text-xs text-slate-300">|</span>
                             <Button 
                                variant="ghost" 
                                size="sm" 
                                className="h-auto p-0 text-xs text-blue-600 hover:text-blue-800 hover:bg-transparent"
                                onClick={() => setSelectedReportIds([])}
                             >
                                None
                             </Button>
                        </div>
                    </div>
                    <div className="grid grid-cols-1 gap-2 max-h-[250px] overflow-y-auto border p-3 rounded-md bg-slate-50 dark:bg-slate-900/50">
                        {REPORTS.map((report) => (
                            <div key={report.id} className="flex items-center space-x-2 hover:bg-slate-100 dark:hover:bg-slate-800 p-1 rounded transition-colors">
                                <Checkbox 
                                    id={`report-${report.id}`} 
                                    checked={selectedReportIds.includes(report.id)}
                                    onCheckedChange={(checked) => {
                                        if (checked) {
                                            setSelectedReportIds(prev => [...prev, report.id]);
                                        } else {
                                            setSelectedReportIds(prev => prev.filter(id => id !== report.id));
                                        }
                                    }}
                                />
                                <label 
                                    htmlFor={`report-${report.id}`}
                                    className="text-sm font-medium leading-none peer-disabled:cursor-not-allowed peer-disabled:opacity-70 cursor-pointer flex-1 py-1"
                                >
                                    {report.title}
                                </label>
                            </div>
                        ))}
                    </div>
                </div>
            </div>
            <DialogFooter>
                <Button variant="outline" onClick={() => setEmailDialogOpen(false)} disabled={sendingEmail}>
                    Cancel
                </Button>
                <Button onClick={handleSendEmail} disabled={sendingEmail || selectedReportIds.length === 0}>
                    {sendingEmail ? (
                        <>
                            <Loader className="mr-2 h-4 w-4 animate-spin" />
                            Sending...
                        </>
                    ) : (
                        'Send'
                    )}
                </Button>
            </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}
