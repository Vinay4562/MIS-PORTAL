import React, { useState, useEffect } from 'react';
import axios from 'axios';
import { Checkbox } from "@/components/ui/checkbox";
import { Label } from "@/components/ui/label";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Download, FileText } from 'lucide-react';
import { toast } from 'sonner';

const API = `${process.env.REACT_APP_BACKEND_URL}/api`;

const FEEDER_ORDER = [
  "Bus Voltages & Station Load",
  "400KV MAHESHWARAM-2",
  "400KV MAHESHWARAM-1",
  "400KV NARSAPUR-1",
  "400KV NARSAPUR-2",
  "400KV KETHIREDDYPALLY-1",
  "400KV KETHIREDDYPALLY-2",
  "400KV NIZAMABAD-1",
  "400KV NIZAMABAD-2",
  "ICT-1 (315MVA)",
  "ICT-2 (315MVA)",
  "ICT-3 (315MVA)",
  "ICT-4 (500MVA)",
  "220KV PARIGI-1",
  "220KV PARIGI-2",
  "220KV THANDUR",
  "220KV GACHIBOWLI-1",
  "220KV GACHIBOWLI-2",
  "220KV KETHIREDDYPALLY",
  "220KV YEDDUMAILARAM-1",
  "220KV YEDDUMAILARAM-2",
  "220KV SADASIVAPET-1",
  "220KV SADASIVAPET-2"
];

export default function ReportGenerator() {
  const [feeders, setFeeders] = useState([]);
  const [selectedFeeders, setSelectedFeeders] = useState([]);
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    fetchFeeders();
  }, []);

  const fetchFeeders = async () => {
    try {
      const response = await axios.get(`${API}/feeders`);
      const sorted = response.data.sort((a, b) => {
        const indexA = FEEDER_ORDER.indexOf(a.name);
        const indexB = FEEDER_ORDER.indexOf(b.name);
        if (indexA !== -1 && indexB !== -1) return indexA - indexB;
        if (indexA !== -1) return -1;
        if (indexB !== -1) return 1;
        return a.name.localeCompare(b.name);
      });
      setFeeders(sorted);
    } catch (error) {
      console.error('Failed to fetch feeders:', error);
      toast.error('Failed to load feeders');
    }
  };

  const handleSelectAll = (checked) => {
    if (checked) {
      setSelectedFeeders(feeders.map(f => f.id));
    } else {
      setSelectedFeeders([]);
    }
  };

  const handleSelectFeeder = (id, checked) => {
    if (checked) {
      setSelectedFeeders([...selectedFeeders, id]);
    } else {
      setSelectedFeeders(selectedFeeders.filter(fid => fid !== id));
    }
  };

  const handleGenerateReport = () => {
    if (selectedFeeders.length === 0) {
      toast.error('Please select at least one feeder');
      return;
    }
    // TODO: Implement report generation logic
    toast.success(`Generating report for ${selectedFeeders.length} feeders...`);
  };

  return (
    <Card className="w-full max-w-2xl shadow-lg">
      <CardHeader>
        <CardTitle className="flex items-center gap-2">
          <FileText className="w-5 h-5 text-blue-600" />
          Report Generator
        </CardTitle>
      </CardHeader>
      <CardContent className="space-y-6">
        <div className="space-y-4">
          <div className="flex items-center justify-between pb-2 border-b border-slate-200 dark:border-slate-700">
            <Label className="text-base font-semibold text-slate-700 dark:text-slate-300">
              Select Feeders
            </Label>
            <div className="flex items-center space-x-2">
              <Checkbox 
                id="select-all" 
                checked={feeders.length > 0 && selectedFeeders.length === feeders.length}
                onCheckedChange={handleSelectAll}
              />
              <Label htmlFor="select-all" className="cursor-pointer font-medium">
                Select All
              </Label>
            </div>
          </div>
          
          <ScrollArea className="h-[300px] rounded-md border p-4 bg-slate-50 dark:bg-slate-900">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
              {feeders.map((feeder) => (
                <div key={feeder.id} className="flex items-center space-x-2 p-2 rounded hover:bg-white dark:hover:bg-slate-800 transition-colors">
                  <Checkbox 
                    id={`feeder-${feeder.id}`} 
                    checked={selectedFeeders.includes(feeder.id)}
                    onCheckedChange={(checked) => handleSelectFeeder(feeder.id, checked)}
                  />
                  <Label 
                    htmlFor={`feeder-${feeder.id}`}
                    className="cursor-pointer flex-1 text-sm text-slate-600 dark:text-slate-400"
                  >
                    {feeder.name}
                  </Label>
                </div>
              ))}
            </div>
          </ScrollArea>
          
          <div className="flex justify-between items-center pt-2">
            <span className="text-sm text-slate-500">
              {selectedFeeders.length} of {feeders.length} selected
            </span>
          </div>
        </div>

        <Button 
          className="w-full" 
          onClick={handleGenerateReport}
          disabled={selectedFeeders.length === 0 || loading}
        >
          <Download className="w-4 h-4 mr-2" />
          Generate Report
        </Button>
      </CardContent>
    </Card>
  );
}
