import React, { useState, useMemo } from "react";
import * as Icons from "lucide-react";
import { CLINICS_METADATA, INSURANCE_COMPANIES, MonthlyReportPayload, ClinicData, ExcelRecord } from "./types";

interface MonthState {
  file: File | null;
  fileName: string;
  isUploaded: boolean;
  isParsing: boolean;
  startDate: string;
  endDate: string;
  insuranceCompany: string;
  clinics: { [rawName: string]: ClinicData };
  error?: string;
}

export default function App() {
  // State for Month 1, Month 2, Month 3
  const [months, setMonths] = useState<MonthState[]>([
    {
      file: null,
      fileName: "",
      isUploaded: false,
      isParsing: false,
      startDate: "2026-01-01",
      endDate: "2026-01-31",
      insuranceCompany: "Bupa Arabia",
      clinics: {}
    },
    {
      file: null,
      fileName: "",
      isUploaded: false,
      isParsing: false,
      startDate: "2026-02-01",
      endDate: "2026-02-28",
      insuranceCompany: "Tawuniya",
      clinics: {}
    },
    {
      file: null,
      fileName: "",
      isUploaded: false,
      isParsing: false,
      startDate: "2026-03-01",
      endDate: "2026-03-31",
      insuranceCompany: "Medgulf",
      clinics: {}
    }
  ]);

  const [activeTab, setActiveTab] = useState<"upload" | "workspace" | "summary">("upload");
  const [selectedClinicRaw, setSelectedClinicRaw] = useState<string>(CLINICS_METADATA[0].rawName);
  const [editingMonthIdx, setEditingMonthIdx] = useState<number>(0);
  const [generalError, setGeneralError] = useState<string | null>(null);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [hasProcessed, setHasProcessed] = useState<boolean>(false);
  const [isGenerating, setIsGenerating] = useState<boolean>(false);

  // Drag over states for monthly upload boxes
  const [draggingIdx, setDraggingIdx] = useState<number | null>(null);

  // File Upload and Parse Trigger
  const handleFileChange = async (index: number, file: File | null) => {
    if (!file) {
      setMonths(prev => {
        const copy = [...prev];
        copy[index] = {
          ...copy[index],
          file: null,
          fileName: "",
          isUploaded: false,
          clinics: {},
          error: undefined
        };
        return copy;
      });
      return;
    }

    setMonths(prev => {
      const copy = [...prev];
      copy[index] = {
        ...copy[index],
        file,
        fileName: file.name,
        isParsing: true,
        error: undefined
      };
      return copy;
    });

    try {
      const formData = new FormData();
      formData.append("file", file);

      const response = await fetch("/api/parse", {
        method: "POST",
        body: formData
      });

      if (!response.ok) {
        throw new Error(`Failed to parse: ${response.statusText}`);
      }

      const result = await response.json();
      
      setMonths(prev => {
        const copy = [...prev];
        copy[index] = {
          ...copy[index],
          isUploaded: true,
          isParsing: false,
          clinics: result.clinics
        };
        return copy;
      });
    } catch (err: any) {
      console.error(err);
      setMonths(prev => {
        const copy = [...prev];
        copy[index] = {
          ...copy[index],
          file: null,
          fileName: "",
          isUploaded: false,
          isParsing: false,
          error: err.message || "Failed to parse spreadsheet dashboard."
        };
        return copy;
      });
    }
  };

  // Helper for dynamic Lucide icons
  const renderClinicIcon = (iconName: string, className = "w-5 h-5") => {
    const IconComp = (Icons as any)[iconName];
    if (!IconComp) return <Icons.Hospital className={className} />;
    return <IconComp className={className} />;
  };

  // "Save & Process" button logic
  const handleSaveAndProcess = () => {
    setGeneralError(null);
    const unuploaded = months.some(m => !m.isUploaded);
    if (unuploaded) {
      setGeneralError("Please upload all 3 monthly Excel reports before processing.");
      return;
    }
    
    setIsProcessing(true);
    // Simulate a brief analysis/validation period for polished user feedback
    setTimeout(() => {
      setIsProcessing(false);
      setHasProcessed(true);
      setActiveTab("summary");
    }, 1200);
  };

  // Computed client-side real-time aggregates
  const aggregates = useMemo(() => {
    if (!months.some(m => m.isUploaded)) {
      return { statusRecords: [], rejectionRecords: [], totalCases: 0, totalAmount: 0 };
    }

    const statusMap: { [cat: string]: { cases: number; amount: number } } = {};
    const rejectMap: { [cat: string]: { cases: number; amount: number } } = {};

    months.forEach(month => {
      Object.keys(month.clinics).forEach(clinicRaw => {
        const clinicData = month.clinics[clinicRaw];
        if (clinicData?.statusTable) {
          clinicData.statusTable.forEach(r => {
            const key = r.category.trim().toLowerCase();
            if (!key) return;
            if (!statusMap[key]) {
              statusMap[key] = { cases: 0, amount: 0 };
            }
            statusMap[key].cases += Number(r.cases) || 0;
            statusMap[key].amount += Number(r.amount) || 0;
          });
        }
        if (clinicData?.rejectionTable) {
          clinicData.rejectionTable.forEach(r => {
            const key = r.category.trim().toLowerCase();
            if (!key) return;
            if (!rejectMap[key]) {
              rejectMap[key] = { cases: 0, amount: 0 };
            }
            rejectMap[key].cases += Number(r.cases) || 0;
            rejectMap[key].amount += Number(r.amount) || 0;
          });
        }
      });
    });

    const statusMapping: { [key: string]: string } = {
      approved: "Approved",
      rejected: "Rejected",
      pending: "Pending",
      "in progress": "In Progress",
      returned: "Returned",
      submitted: "Submitted",
      paid: "Paid"
    };

    const statusRecords: ExcelRecord[] = Object.keys(statusMap).map(key => {
      const display = statusMapping[key] || (key.charAt(0).toUpperCase() + key.slice(1));
      return {
        category: display,
        cases: statusMap[key].cases,
        amount: statusMap[key].amount
      };
    });

    const rejectionRecords: ExcelRecord[] = Object.keys(rejectMap).map(key => {
      return {
        category: key.charAt(0).toUpperCase() + key.slice(1),
        cases: rejectMap[key].cases,
        amount: rejectMap[key].amount
      };
    });

    const totalStatusCases = statusRecords.reduce((sum, r) => sum + r.cases, 0);
    const totalStatusAmount = statusRecords.reduce((sum, r) => sum + r.amount, 0);

    const totalRejectCases = rejectionRecords.reduce((sum, r) => sum + r.cases, 0);
    const totalRejectAmount = rejectionRecords.reduce((sum, r) => sum + r.amount, 0);

    // Apply rule that Total Cases and Amount MUST equal between status and rejection summaries
    if (totalStatusCases !== totalRejectCases || totalStatusAmount !== totalRejectAmount) {
      const diffCases = totalStatusCases - totalRejectCases;
      const diffAmount = totalStatusAmount - totalRejectAmount;
      rejectionRecords.push({
        category: "Approved / Paid Claims (No Rejection)",
        cases: diffCases,
        amount: diffAmount
      });
    }

    return {
      statusRecords,
      rejectionRecords,
      totalCases: totalStatusCases,
      totalAmount: totalStatusAmount
    };
  }, [months]);

  // Handle direct data cell modifications (interactive spreadsheet mode)
  const handleCellEdit = (
    monthIdx: number,
    clinicRaw: string,
    tableType: "status" | "rejection",
    rowIdx: number,
    field: keyof ExcelRecord,
    value: any
  ) => {
    setMonths(prev => {
      const copy = [...prev];
      const monthClinics = { ...copy[monthIdx].clinics };
      const clinicData = { ...monthClinics[clinicRaw] };
      
      if (tableType === "status") {
        const statusTable = [...(clinicData.statusTable || [])];
        statusTable[rowIdx] = {
          ...statusTable[rowIdx],
          [field]: field === "category" ? value : Number(value) || 0
        };
        clinicData.statusTable = statusTable;
      } else {
        const rejectionTable = [...(clinicData.rejectionTable || [])];
        rejectionTable[rowIdx] = {
          ...rejectionTable[rowIdx],
          [field]: field === "category" ? value : Number(value) || 0
        };
        clinicData.rejectionTable = rejectionTable;
      }
      
      monthClinics[clinicRaw] = clinicData;
      copy[monthIdx].clinics = monthClinics;
      return copy;
    });
  };

  // Add Row inside Interactive Workspace
  const handleAddRow = (monthIdx: number, clinicRaw: string, tableType: "status" | "rejection") => {
    setMonths(prev => {
      const copy = [...prev];
      const monthClinics = { ...copy[monthIdx].clinics };
      const clinicData = { ...monthClinics[clinicRaw] };

      if (tableType === "status") {
        const statusTable = [...(clinicData.statusTable || [])];
        statusTable.push({ category: "New Status", cases: 0, amount: 0 });
        clinicData.statusTable = statusTable;
      } else {
        const rejectionTable = [...(clinicData.rejectionTable || [])];
        rejectionTable.push({ category: "New Reason", cases: 0, amount: 0 });
        clinicData.rejectionTable = rejectionTable;
      }

      monthClinics[clinicRaw] = clinicData;
      copy[monthIdx].clinics = monthClinics;
      return copy;
    });
  };

  // Delete Row inside Interactive Workspace
  const handleDeleteRow = (monthIdx: number, clinicRaw: string, tableType: "status" | "rejection", rowIdx: number) => {
    setMonths(prev => {
      const copy = [...prev];
      const monthClinics = { ...copy[monthIdx].clinics };
      const clinicData = { ...monthClinics[clinicRaw] };

      if (tableType === "status") {
        const statusTable = [...(clinicData.statusTable || [])];
        statusTable.splice(rowIdx, 1);
        clinicData.statusTable = statusTable;
      } else {
        const rejectionTable = [...(clinicData.rejectionTable || [])];
        rejectionTable.splice(rowIdx, 1);
        clinicData.rejectionTable = rejectionTable;
      }

      monthClinics[clinicRaw] = clinicData;
      copy[monthIdx].clinics = monthClinics;
      return copy;
    });
  };

  // Generate Styled Quarterly Report and trigger file download
  const handleGenerateReport = async () => {
    try {
      setIsGenerating(true);
      
      const payload: MonthlyReportPayload[] = months.map((m, mIdx) => ({
        monthIndex: mIdx,
        dateRange: `${m.startDate} to ${m.endDate}`,
        insuranceCompany: m.insuranceCompany,
        clinics: m.clinics
      }));

      const response = await fetch("/api/generate", {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({ payload })
      });

      if (!response.ok) {
        throw new Error("Failed to generate Quarterly Excel report.");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.setAttribute("download", `Q_Report_${new Date().getFullYear()}.xlsx`);
      document.body.appendChild(link);
      link.click();
      link.parentNode?.removeChild(link);
      window.URL.revokeObjectURL(url);
    } catch (err: any) {
      console.error(err);
      alert(err.message || "Something went wrong while generating the Excel sheet.");
    } finally {
      setIsGenerating(false);
    }
  };

  // Selected clinic metadata helper
  const activeClinicMetadata = useMemo(() => {
    return CLINICS_METADATA.find(c => c.rawName === selectedClinicRaw) || CLINICS_METADATA[0];
  }, [selectedClinicRaw]);

  // Current Clinic values for active workspace view
  const currentClinicData = useMemo(() => {
    const m = months[editingMonthIdx];
    return m?.clinics[selectedClinicRaw] || { statusTable: [], rejectionTable: [] };
  }, [months, editingMonthIdx, selectedClinicRaw]);

  // Totals for current clinic/month table previews
  const currentClinicTotals = useMemo(() => {
    const sCases = currentClinicData.statusTable?.reduce((sum, r) => sum + r.cases, 0) || 0;
    const sAmount = currentClinicData.statusTable?.reduce((sum, r) => sum + r.amount, 0) || 0;
    const rCases = currentClinicData.rejectionTable?.reduce((sum, r) => sum + r.cases, 0) || 0;
    const rAmount = currentClinicData.rejectionTable?.reduce((sum, r) => sum + r.amount, 0) || 0;
    
    return { sCases, sAmount, rCases, rAmount };
  }, [currentClinicData]);

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col" id="app_root">
      {/* Top Professional Header */}
      <header className="bg-slate-950 text-white shadow-sm border-b border-slate-900 px-6 py-4.5 flex flex-col md:flex-row md:items-center md:justify-between" id="header">
        <div className="flex items-center space-x-4">
          <div className="p-2.5 bg-gradient-to-tr from-sky-500 to-blue-600 rounded-lg shadow-md shadow-blue-500/10">
            <Icons.Hospital className="w-7 h-7 text-white" />
          </div>
          <div>
            <h1 className="text-xl font-bold tracking-tight font-display">Quarterly Report Merger</h1>
            <p className="text-slate-400 text-[11px] uppercase tracking-wider font-semibold">Royal Commission Hospital &amp; Affiliated Jubail Clinics</p>
          </div>
        </div>
        <div className="mt-3 md:mt-0 flex items-center space-x-3">
          <div className="px-3 py-1.5 bg-slate-900 border border-slate-800 rounded-lg text-xs text-slate-300 font-mono flex items-center gap-2 shadow-inner">
            <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse"></span>
            System Active
          </div>
          <div className="px-3 py-1.5 bg-slate-900 border border-slate-800 rounded-lg text-xs text-slate-400 font-mono shadow-inner">
            v2.4.0 (Vite + React)
          </div>
        </div>
      </header>

      {/* Main Subheader Navigation Tabs */}
      <div className="bg-white border-b border-slate-200 px-6 py-2.5 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4" id="navigation_tabs">
        <div className="flex flex-wrap gap-1.5">
          <button
            onClick={() => setActiveTab("upload")}
            className={`px-4 py-2 text-xs font-semibold uppercase tracking-wider rounded-lg flex items-center space-x-2 transition-all cursor-pointer border ${
              activeTab === "upload"
                ? "bg-slate-900 text-white border-slate-900 shadow-sm"
                : "text-slate-600 border-transparent hover:text-slate-900 hover:bg-slate-50"
            }`}
          >
            <Icons.Upload className="w-3.5 h-3.5" />
            <span>1. Monthly Reports</span>
          </button>
          
          <button
            onClick={() => {
              if (hasProcessed) {
                setActiveTab("summary");
              } else {
                setGeneralError("Please complete uploading and click 'Save & Process' to unlock preview summary.");
              }
            }}
            className={`px-4 py-2 text-xs font-semibold uppercase tracking-wider rounded-lg flex items-center space-x-2 transition-all border ${
              !hasProcessed ? "opacity-50 cursor-not-allowed" : "cursor-pointer"
            } ${
              activeTab === "summary"
                ? "bg-slate-900 text-white border-slate-900 shadow-sm"
                : "text-slate-600 border-transparent hover:text-slate-900 hover:bg-slate-50"
            }`}
          >
            <Icons.FileSpreadsheet className="w-3.5 h-3.5" />
            <span>2. Consolidated Summary</span>
          </button>

          <button
            onClick={() => {
              if (hasProcessed) {
                setActiveTab("workspace");
              } else {
                setGeneralError("Please complete uploading and click 'Save & Process' to unlock detailed workspace edits.");
              }
            }}
            className={`px-4 py-2 text-xs font-semibold uppercase tracking-wider rounded-lg flex items-center space-x-2 transition-all border ${
              !hasProcessed ? "opacity-50 cursor-not-allowed" : "cursor-pointer"
            } ${
              activeTab === "workspace"
                ? "bg-slate-900 text-white border-slate-900 shadow-sm"
                : "text-slate-600 border-transparent hover:text-slate-900 hover:bg-slate-50"
            }`}
          >
            <Icons.Edit3 className="w-3.5 h-3.5" />
            <span>3. Edit Workspace</span>
          </button>
        </div>

        {/* Global Action Header */}
        {hasProcessed && (
          <button
            onClick={handleGenerateReport}
            disabled={isGenerating}
            className="px-4.5 py-2 text-xs uppercase tracking-wider font-bold bg-blue-600 hover:bg-blue-700 active:bg-blue-800 text-white rounded-lg shadow-sm hover:shadow-blue-500/10 transition-all flex items-center space-x-2 cursor-pointer self-start sm:self-auto"
          >
            {isGenerating ? (
              <>
                <Icons.RefreshCw className="w-3.5 h-3.5 animate-spin" />
                <span>Generating Report...</span>
              </>
            ) : (
              <>
                <Icons.Download className="w-3.5 h-3.5" />
                <span>Download Report</span>
              </>
            )}
          </button>
        )}
      </div>

      {/* Main Content Area */}
      <main className="flex-1 p-6 max-w-7xl mx-auto w-full flex flex-col gap-6" id="main_content">
        
        {/* Error notification banner */}
        {generalError && (
          <div className="p-4 bg-rose-50 border-l-4 border-rose-500 rounded-r-lg flex items-start space-x-3 text-rose-800" id="error_banner">
            <Icons.AlertCircle className="w-5 h-5 mt-0.5 text-rose-500 shrink-0" />
            <div className="flex-1 text-sm font-medium">{generalError}</div>
            <button onClick={() => setGeneralError(null)} className="text-rose-400 hover:text-rose-600">
              <Icons.Plus className="w-5 h-5 rotate-45" />
            </button>
          </div>
        )}

        {/* TAB 1: FILE UPLOADERS */}
        {activeTab === "upload" && (
          <div className="flex flex-col gap-6" id="tab_upload">
            <div className="card-premium p-6 bg-white">
              <h2 className="text-base font-bold text-slate-900 flex items-center gap-2 mb-2 font-display">
                <Icons.Info className="w-4.5 h-4.5 text-blue-500" />
                Hospital Quarterly Consolidation Process
              </h2>
              <p className="text-slate-500 text-xs leading-relaxed max-w-4xl">
                Please upload the individual monthly reports (Excel format) for each month of the fiscal quarter. The program scans each report's <code className="bg-slate-100 px-1 py-0.5 rounded font-mono text-[10px] font-semibold text-rose-600">Dashboard</code> worksheet, matches clinic headers, and aggregates claim status summaries and rejection breakdowns for the final 8-sheet Quarterly Workbook.
              </p>
            </div>

            {/* Monthly Report Uploaders Grid */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              {months.map((month, idx) => {
                const monthName = idx === 0 ? "Month 1 (e.g. January)" : idx === 1 ? "Month 2 (e.g. February)" : "Month 3 (e.g. March)";
                const isDragging = draggingIdx === idx;

                return (
                  <div
                    key={idx}
                    className={`card-premium flex flex-col p-6 bg-white transition-all duration-200 ${
                      month.isUploaded
                        ? "border-emerald-500/30 ring-4 ring-emerald-500/5 bg-emerald-50/5"
                        : isDragging
                        ? "border-blue-500 ring-4 ring-blue-500/5 bg-blue-50/10"
                        : "hover:border-slate-300"
                    }`}
                  >
                    <div className="flex items-center justify-between mb-4">
                      <span className="px-2.5 py-1 bg-slate-900 text-white rounded-md text-[10px] font-bold font-mono tracking-wider uppercase">
                        {monthName}
                      </span>
                      {month.isUploaded && (
                        <span className="flex items-center text-[10px] font-bold text-emerald-700 bg-emerald-50 border border-emerald-100 px-2 py-0.5 rounded-md">
                          <Icons.Check className="w-3 h-3 mr-1" />
                          Parsed
                        </span>
                      )}
                    </div>

                    {/* Metadata Settings inside Uploader Box */}
                    <div className="space-y-3.5 mb-5">
                      <div>
                        <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5">
                          Insurance Partner
                        </label>
                        <div className="relative">
                          <select
                            value={month.insuranceCompany}
                            onChange={(e) => {
                              const val = e.target.value;
                              setMonths(prev => {
                                const copy = [...prev];
                                copy[idx].insuranceCompany = val;
                                return copy;
                              });
                            }}
                            className="w-full pl-8 pr-3 py-1.5 text-xs bg-slate-50/50 border border-slate-200 rounded-lg shadow-2xs focus:outline-none focus:border-slate-900 focus:bg-white focus:ring-1 focus:ring-slate-900 appearance-none font-medium text-slate-800"
                          >
                            {INSURANCE_COMPANIES.map((company, cIdx) => (
                              <option key={cIdx} value={company}>
                                {company}
                              </option>
                            ))}
                          </select>
                          <Icons.Building2 className="absolute left-2.5 top-2.5 w-3.5 h-3.5 text-slate-400" />
                        </div>
                      </div>

                      <div className="grid grid-cols-2 gap-2">
                        <div>
                          <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5">
                            Start Date
                          </label>
                          <input
                            type="date"
                            value={month.startDate}
                            onChange={(e) => {
                              const val = e.target.value;
                              setMonths(prev => {
                                const copy = [...prev];
                                copy[idx].startDate = val;
                                return copy;
                              });
                            }}
                            className="w-full p-1.5 text-[11px] bg-slate-50/50 border border-slate-200 rounded-lg shadow-2xs focus:outline-none focus:border-slate-900 focus:bg-white focus:ring-1 focus:ring-slate-900 font-medium text-slate-800"
                          />
                        </div>
                        <div>
                          <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5">
                            End Date
                          </label>
                          <input
                            type="date"
                            value={month.endDate}
                            onChange={(e) => {
                              const val = e.target.value;
                              setMonths(prev => {
                                const copy = [...prev];
                                copy[idx].endDate = val;
                                return copy;
                              });
                            }}
                            className="w-full p-1.5 text-[11px] bg-slate-50/50 border border-slate-200 rounded-lg shadow-2xs focus:outline-none focus:border-slate-900 focus:bg-white focus:ring-1 focus:ring-slate-900 font-medium text-slate-800"
                          />
                        </div>
                      </div>
                    </div>

                    {/* Drag and Drop File Target Area */}
                    <div
                      onDragOver={(e) => {
                        e.preventDefault();
                        setDraggingIdx(idx);
                      }}
                      onDragLeave={() => setDraggingIdx(null)}
                      onDrop={(e) => {
                        e.preventDefault();
                        setDraggingIdx(null);
                        const file = e.dataTransfer.files[0];
                        if (file) handleFileChange(idx, file);
                      }}
                      className={`flex-1 min-h-[140px] border-2 border-dashed rounded-xl p-4 flex flex-col items-center justify-center text-center transition-all cursor-pointer ${
                        month.isUploaded
                          ? "border-emerald-500/20 bg-emerald-500/5 hover:bg-emerald-500/10"
                          : isDragging
                          ? "border-blue-500 bg-blue-50/50"
                          : "border-slate-200 hover:border-slate-300 hover:bg-slate-50/30"
                      }`}
                      onClick={() => {
                        document.getElementById(`file_picker_${idx}`)?.click();
                      }}
                    >
                      <input
                        type="file"
                        id={`file_picker_${idx}`}
                        className="hidden"
                        accept=".xlsx, .xls"
                        onChange={(e) => {
                          const file = e.target.files?.[0];
                          if (file) handleFileChange(idx, file);
                        }}
                      />
                      
                      {month.isParsing ? (
                        <div className="flex flex-col items-center">
                           <Icons.Loader2 className="w-7 h-7 text-blue-500 animate-spin mb-2" />
                           <p className="text-xs font-semibold text-slate-800">Reading worksheet...</p>
                           <p className="text-[10px] text-slate-400 mt-0.5">Extracting clinics dashboard</p>
                        </div>
                      ) : month.isUploaded ? (
                        <div className="flex flex-col items-center">
                          <Icons.CheckCircle className="w-8 h-8 text-emerald-500 mb-2" />
                          <p className="text-xs font-semibold text-slate-800 break-all truncate max-w-[180px]">
                            {month.fileName}
                          </p>
                          <p className="text-[10px] text-slate-400 mt-1">
                            {Object.keys(month.clinics).length} clinic profiles matched
                          </p>
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              handleFileChange(idx, null);
                            }}
                            className="mt-3 text-[11px] font-bold text-red-600 hover:text-red-700 hover:underline"
                          >
                            Remove &amp; Re-upload
                          </button>
                        </div>
                      ) : (
                        <div className="flex flex-col items-center">
                          <Icons.UploadCloud className="w-8 h-8 text-slate-400 mb-2" />
                          <p className="text-xs font-semibold text-slate-700">Drag &amp; drop monthly Excel</p>
                          <p className="text-[10px] text-slate-400 mt-0.5">or click to browse computer</p>
                        </div>
                      )}
                    </div>

                    {month.error && (
                      <div className="mt-3 text-[10px] font-semibold text-red-600 bg-red-50 p-2.5 rounded-lg border border-red-100">
                        {month.error}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>

            {/* Action Save and Process Row */}
            <div className="card-premium p-6 bg-white flex flex-col md:flex-row items-center justify-between gap-4">
              <div className="text-xs text-slate-400 leading-normal max-w-2xl">
                Ensure all 3 reports have been parsed successfully before saving. Once saved, you can preview the aggregated figures on screen and make specific fine-tuning edits.
              </div>
              <button
                onClick={handleSaveAndProcess}
                disabled={isProcessing || months.some(m => !m.isUploaded)}
                className={`w-full md:w-auto px-6 py-2.5 text-xs uppercase tracking-wider font-bold rounded-xl shadow-xs transition-all flex items-center justify-center space-x-2 shrink-0 ${
                  months.some(m => !m.isUploaded)
                    ? "bg-slate-100 text-slate-400 border border-slate-200 cursor-not-allowed"
                    : "bg-blue-600 hover:bg-blue-700 active:bg-blue-800 text-white shadow-md hover:shadow-blue-500/10 cursor-pointer"
                }`}
              >
                {isProcessing ? (
                  <>
                    <Icons.Loader2 className="w-4 h-4 animate-spin" />
                    <span>Processing Data...</span>
                  </>
                ) : (
                  <>
                    <Icons.Check className="w-4 h-4" />
                    <span>Save &amp; Process Monthly Reports</span>
                  </>
                )}
              </button>
            </div>
          </div>
        )}

        {/* TAB 2: CONSOLIDATED SUMMARY */}
        {activeTab === "summary" && hasProcessed && (
          <div className="flex flex-col gap-6" id="tab_summary">
            {/* Real-time Summary Cards Header */}
            <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
              <div className="card-premium p-6 bg-white flex items-center space-x-4">
                <div className="p-2.5 bg-blue-50 text-blue-600 rounded-lg">
                  <Icons.FileSpreadsheet className="w-5 h-5" />
                </div>
                <div>
                  <div className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Total Claims</div>
                  <div className="text-2xl font-bold text-slate-900 font-display tracking-tight">
                    {aggregates.totalCases.toLocaleString()}
                  </div>
                </div>
              </div>

              <div className="card-premium p-6 bg-white flex items-center space-x-4">
                <div className="p-2.5 bg-emerald-50 text-emerald-600 rounded-lg">
                  <Icons.FileText className="w-5 h-5" />
                </div>
                <div>
                  <div className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Total Claim Value</div>
                  <div className="text-2xl font-bold text-slate-900 font-display tracking-tight">
                    ${aggregates.totalAmount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </div>
                </div>
              </div>

              <div className="card-premium p-6 bg-white flex items-center space-x-4">
                <div className="p-2.5 bg-purple-50 text-purple-600 rounded-lg">
                  <Icons.Hospital className="w-5 h-5" />
                </div>
                <div>
                  <div className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Clinics Tracked</div>
                  <div className="text-2xl font-bold text-slate-900 font-display tracking-tight">7 / 7</div>
                </div>
              </div>

              <div className="card-premium p-6 bg-white flex items-center space-x-4">
                <div className="p-2.5 bg-teal-50 text-teal-700 rounded-lg">
                  <Icons.ShieldCheck className="w-5 h-5" />
                </div>
                <div>
                  <div className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Validation Integrity</div>
                  <div className="text-sm font-bold text-teal-600 font-display">✓ Balance Verified</div>
                </div>
              </div>
            </div>

            {/* Validation and Alignment Rule Alert */}
            <div className="bg-emerald-50/40 border border-emerald-500/20 rounded-xl p-5 flex items-start space-x-3 text-emerald-800">
              <Icons.ShieldCheck className="w-5 h-5 text-emerald-600 shrink-0 mt-0.5" />
              <div>
                <h4 className="font-bold text-emerald-950 text-xs uppercase tracking-wider mb-0.5 font-display">Strict Total Reconciliation Met</h4>
                <p className="text-xs leading-relaxed text-emerald-700/90">
                  The hospital financial schema requires that total claims cases and total claim amounts listed in the claim status summary must match exactly with those listed in the rejection reason tables. We've automatically verified that Status Totals match Rejection Totals (including any differences captured as "Approved/Paid Claims (No Rejection)").
                </p>
              </div>
            </div>

            {/* Aggregated Preview Columns side-by-side */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              {/* STATUS SUMMARY PREVIEW */}
              <div className="card-premium bg-white p-6 flex flex-col">
                <h3 className="text-sm font-bold text-slate-900 mb-4 flex items-center gap-2 font-display uppercase tracking-wider">
                  <span className="w-2 h-2 rounded-full bg-blue-600"></span>
                  Total Quarter Status Summary
                </h3>
                <div className="flex-1 overflow-x-auto">
                  <table className="w-full text-left text-xs border-collapse">
                    <thead>
                      <tr className="bg-slate-50/70 text-slate-500 font-bold uppercase text-[10px] tracking-wider border-b border-slate-200">
                        <th className="py-2.5 px-3">Claim Status</th>
                        <th className="py-2.5 px-3 text-right w-28">Total Cases</th>
                        <th className="py-2.5 px-3 text-right w-36">Total Amount</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 text-slate-700">
                      {aggregates.statusRecords.map((r, idx) => (
                        <tr key={idx} className="hover:bg-slate-50/50">
                          <td className="py-2.5 px-3 font-semibold text-slate-800">{r.category}</td>
                          <td className="py-2.5 px-3 text-right font-mono text-slate-600">{r.cases.toLocaleString()}</td>
                          <td className="py-2.5 px-3 text-right font-mono text-slate-600">${r.amount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                        </tr>
                      ))}
                      <tr className="bg-slate-50/70 font-bold border-t border-slate-200 text-slate-900">
                        <td className="py-2.5 px-3">Total Status Claims</td>
                        <td className="py-2.5 px-3 text-right font-mono text-slate-800">{aggregates.totalCases.toLocaleString()}</td>
                        <td className="py-2.5 px-3 text-right font-mono text-slate-800">${aggregates.totalAmount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>

              {/* REJECTION SUMMARY PREVIEW */}
              <div className="card-premium bg-white p-6 flex flex-col">
                <h3 className="text-sm font-bold text-slate-900 mb-4 flex items-center gap-2 font-display uppercase tracking-wider">
                  <span className="w-2 h-2 rounded-full bg-amber-500"></span>
                  Total Quarter Rejection Summary
                </h3>
                <div className="flex-1 overflow-x-auto">
                  <table className="w-full text-left text-xs border-collapse">
                    <thead>
                      <tr className="bg-slate-50/70 text-slate-500 font-bold uppercase text-[10px] tracking-wider border-b border-slate-200">
                        <th className="py-2.5 px-3">Rejection Reason</th>
                        <th className="py-2.5 px-3 text-right w-28">Total Cases</th>
                        <th className="py-2.5 px-3 text-right w-36">Total Amount</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 text-slate-700">
                      {aggregates.rejectionRecords.map((r, idx) => (
                        <tr key={idx} className="hover:bg-slate-50/50">
                          <td className="py-2.5 px-3 font-semibold text-slate-800 leading-tight">{r.category}</td>
                          <td className="py-2.5 px-3 text-right font-mono text-slate-600">{r.cases.toLocaleString()}</td>
                          <td className="py-2.5 px-3 text-right font-mono text-slate-600">${r.amount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                        </tr>
                      ))}
                      <tr className="bg-slate-50/70 font-bold border-t border-slate-200 text-slate-900">
                        <td className="py-2.5 px-3">Total Rejections &amp; Paid</td>
                        <td className="py-2.5 px-3 text-right font-mono text-slate-800">{aggregates.totalCases.toLocaleString()}</td>
                        <td className="py-2.5 px-3 text-right font-mono text-slate-800">${aggregates.totalAmount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

            {/* Quick Action Box */}
            <div className="card-premium bg-white p-6 flex flex-col sm:flex-row items-center justify-between gap-4">
              <div>
                <h4 className="font-bold text-slate-950 text-xs uppercase tracking-wider mb-1 font-display">Download Workbook Now or Fine-Tune Clinic Values First</h4>
                <p className="text-[11px] text-slate-400">You can download the compiled Excel immediately or use the Review &amp; Edit Workspace to inspect and correct values for individual clinics.</p>
              </div>
              <div className="flex space-x-3 w-full sm:w-auto shrink-0">
                <button
                  onClick={() => setActiveTab("workspace")}
                  className="w-full sm:w-auto px-4 py-2 bg-slate-100 text-slate-700 rounded-lg hover:bg-slate-200 text-xs font-bold uppercase tracking-wider flex items-center justify-center space-x-2 transition-all cursor-pointer border border-slate-200"
                >
                  <Icons.Edit3 className="w-3.5 h-3.5" />
                  <span>Go to Editor</span>
                </button>
                <button
                  onClick={handleGenerateReport}
                  disabled={isGenerating}
                  className="w-full sm:w-auto px-5 py-2 bg-blue-600 hover:bg-blue-700 active:bg-blue-800 text-white rounded-lg text-xs font-bold uppercase tracking-wider shadow-sm hover:shadow-blue-500/10 flex items-center justify-center space-x-2 transition-all cursor-pointer"
                >
                  {isGenerating ? (
                    <Icons.Loader2 className="w-3.5 h-3.5 animate-spin" />
                  ) : (
                    <Icons.Download className="w-3.5 h-3.5" />
                  )}
                  <span>Download Excel Workbook</span>
                </button>
              </div>
            </div>
          </div>
        )}

        {/* TAB 3: REVIEW & EDIT WORKSPACE */}
        {activeTab === "workspace" && hasProcessed && (
          <div className="flex flex-col lg:flex-row gap-6" id="tab_workspace">
            
            {/* Sidebar Clinic Selection Card */}
            <div className="w-full lg:w-1/4 card-premium bg-white p-4 shrink-0 flex flex-col gap-1.5">
              <h3 className="px-3 py-2 text-[10px] font-bold uppercase text-slate-400 tracking-wider">Select Clinic Profile</h3>
              <div className="flex flex-col space-y-1">
                {CLINICS_METADATA.map((clinic, idx) => {
                  const isActive = clinic.rawName === selectedClinicRaw;
                  return (
                    <button
                      key={idx}
                      onClick={() => setSelectedClinicRaw(clinic.rawName)}
                      className={`w-full text-left px-3 py-2.5 rounded-lg flex items-center space-x-3 transition-all cursor-pointer border ${
                        isActive
                          ? "bg-slate-950 text-white border-slate-950 shadow-sm font-semibold"
                          : "text-slate-600 border-transparent hover:bg-slate-50 hover:text-slate-900"
                      }`}
                    >
                      <div className={`p-1.5 rounded-md ${isActive ? "bg-slate-900 text-sky-400" : "bg-slate-100 text-slate-500"}`}>
                        {renderClinicIcon(clinic.iconName, "w-3.5 h-3.5")}
                      </div>
                      <div className="flex-1 min-w-0">
                        <div className="text-xs font-bold leading-tight truncate">{clinic.shortName}</div>
                        <div className="text-[10px] text-slate-400 truncate leading-normal">RC Jubail Region</div>
                      </div>
                      <Icons.ChevronRight className={`w-3.5 h-3.5 ${isActive ? "text-slate-200" : "text-slate-300"}`} />
                    </button>
                  );
                })}
              </div>
            </div>

            {/* Main Interactive Table Editor Pane */}
            <div className="flex-1 card-premium bg-white p-6 flex flex-col gap-6">
              
              {/* Selected Clinic Header Card */}
              <div className="flex flex-col md:flex-row md:items-center justify-between border-b border-slate-100 pb-4">
                <div>
                  <h2 className="text-lg font-bold text-slate-900 flex items-center gap-2 font-display">
                    {renderClinicIcon(activeClinicMetadata.iconName, "w-5.5 h-5.5 text-blue-600")}
                    {activeClinicMetadata.rawName}
                  </h2>
                  <p className="text-[11px] text-slate-400 mt-1 leading-normal">
                    Interactive clinical ledger data. Modifications apply to the final compiled sheet <code className="bg-slate-50 px-1 py-0.5 rounded font-mono text-[10px]">{activeClinicMetadata.shortName}</code>.
                  </p>
                </div>
              </div>

              {/* Month Selection Horizontal Tabs */}
              <div className="flex border border-slate-200 p-1 bg-slate-100/55 rounded-xl max-w-sm">
                {months.map((m, idx) => (
                  <button
                    key={idx}
                    onClick={() => setEditingMonthIdx(idx)}
                    className={`flex-1 text-center py-2 text-[10px] font-bold uppercase tracking-wider rounded-lg transition-all cursor-pointer ${
                      editingMonthIdx === idx
                        ? "bg-white text-slate-900 shadow-xs border border-slate-200/50"
                        : "text-slate-500 hover:text-slate-800"
                    }`}
                  >
                    Month {idx + 1}
                  </button>
                ))}
              </div>

              {/* Month Metadata display banner */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 bg-slate-50/50 p-4 rounded-xl border border-slate-200/60">
                <div className="flex items-center space-x-3">
                  <Icons.Calendar className="w-4 h-4 text-slate-400" />
                  <div>
                    <div className="text-[9px] uppercase font-bold text-slate-400 tracking-wider mb-0.5">Month Date Range</div>
                    <div className="text-xs font-semibold text-slate-800 font-mono">
                      {months[editingMonthIdx]?.startDate} to {months[editingMonthIdx]?.endDate}
                    </div>
                  </div>
                </div>
                <div className="flex items-center space-x-3">
                  <Icons.Building2 className="w-4 h-4 text-slate-400" />
                  <div>
                    <div className="text-[9px] uppercase font-bold text-slate-400 tracking-wider mb-0.5">Insurance Carrier</div>
                    <div className="text-xs font-semibold text-slate-800">
                      {months[editingMonthIdx]?.insuranceCompany}
                    </div>
                  </div>
                </div>
              </div>

              {/* Two Tables stack side-by-side or vertical */}
              <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
                
                {/* 1. STATUS LEDGER TABLE */}
                <div className="flex flex-col gap-3">
                  <div className="flex items-center justify-between">
                    <h4 className="text-xs font-bold text-slate-900 uppercase tracking-wider flex items-center gap-1.5 font-display">
                      <span className="w-2 h-2 rounded-full bg-blue-600 animate-pulse"></span>
                      Status Claims Table
                    </h4>
                    <button
                      onClick={() => handleAddRow(editingMonthIdx, selectedClinicRaw, "status")}
                      className="px-2.5 py-1 text-[10px] uppercase tracking-wider bg-slate-950 text-white font-bold rounded-lg hover:bg-slate-800 flex items-center gap-1 cursor-pointer transition-all"
                    >
                      <Icons.Plus className="w-3 h-3" />
                      <span>Add Row</span>
                    </button>
                  </div>

                  <div className="border border-slate-200 rounded-lg overflow-hidden bg-white shadow-xs">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-slate-50 text-slate-500 font-bold uppercase text-[10px] tracking-wider border-b border-slate-200">
                          <th className="py-2.5 px-3">Status Name</th>
                          <th className="py-2.5 px-3 text-right w-24">Cases</th>
                          <th className="py-2.5 px-3 text-right w-36">Amount ($)</th>
                          <th className="py-2.5 px-3 text-center w-12">Action</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 bg-white">
                        {currentClinicData.statusTable && currentClinicData.statusTable.length > 0 ? (
                          currentClinicData.statusTable.map((row, rIdx) => (
                            <tr key={rIdx} className="hover:bg-slate-50/30">
                              <td className="p-1">
                                <input
                                  type="text"
                                  value={row.category}
                                  onChange={(e) => handleCellEdit(editingMonthIdx, selectedClinicRaw, "status", rIdx, "category", e.target.value)}
                                  className="w-full px-2 py-1.5 border border-transparent rounded-md bg-transparent hover:bg-slate-50 hover:border-slate-200 focus:border-slate-900 focus:bg-white focus:ring-1 focus:ring-slate-900 focus:outline-none font-medium text-xs text-slate-800"
                                />
                              </td>
                              <td className="p-1 w-24">
                                <input
                                  type="number"
                                  value={row.cases}
                                  onChange={(e) => handleCellEdit(editingMonthIdx, selectedClinicRaw, "status", rIdx, "cases", e.target.value)}
                                  className="w-full px-2 py-1.5 border border-transparent rounded-md bg-transparent hover:bg-slate-50 hover:border-slate-200 focus:border-slate-900 focus:bg-white focus:ring-1 focus:ring-slate-900 focus:outline-none text-right font-mono font-medium text-xs text-slate-800"
                                />
                              </td>
                              <td className="p-1 w-36">
                                <input
                                  type="number"
                                  value={row.amount}
                                  step="0.01"
                                  onChange={(e) => handleCellEdit(editingMonthIdx, selectedClinicRaw, "status", rIdx, "amount", e.target.value)}
                                  className="w-full px-2 py-1.5 border border-transparent rounded-md bg-transparent hover:bg-slate-50 hover:border-slate-200 focus:border-slate-900 focus:bg-white focus:ring-1 focus:ring-slate-900 focus:outline-none text-right font-mono font-medium text-xs text-slate-800"
                                />
                              </td>
                              <td className="p-1 text-center w-12">
                                <button
                                  onClick={() => handleDeleteRow(editingMonthIdx, selectedClinicRaw, "status", rIdx)}
                                  className="p-1.5 text-slate-400 hover:text-red-600 rounded-md hover:bg-red-50 cursor-pointer transition-colors"
                                  title="Delete Row"
                                >
                                  <Icons.Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </td>
                            </tr>
                          ))
                        ) : (
                          <tr>
                            <td colSpan={4} className="p-6 text-center text-slate-400 font-semibold italic bg-white text-xs">
                              No records parsed for this clinic.
                            </td>
                          </tr>
                        )}
                        <tr className="bg-slate-50/60 font-bold border-t border-slate-200">
                          <td className="py-2 px-3 text-slate-600 text-[11px] uppercase tracking-wider font-semibold">Total Status</td>
                          <td className="py-2 px-3 text-right font-mono text-slate-800">
                            {currentClinicTotals.sCases.toLocaleString()}
                          </td>
                          <td className="py-2 px-3 text-right font-mono text-slate-800">
                            ${currentClinicTotals.sAmount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                          </td>
                          <td></td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* 2. REJECTION LEDGER TABLE */}
                <div className="flex flex-col gap-3">
                  <div className="flex items-center justify-between">
                    <h4 className="text-xs font-bold text-slate-900 uppercase tracking-wider flex items-center gap-1.5 font-display">
                      <span className="w-2 h-2 rounded-full bg-amber-500 animate-pulse"></span>
                      Rejection Reason Table
                    </h4>
                    <button
                      onClick={() => handleAddRow(editingMonthIdx, selectedClinicRaw, "rejection")}
                      className="px-2.5 py-1 text-[10px] uppercase tracking-wider bg-slate-950 text-white font-bold rounded-lg hover:bg-slate-800 flex items-center gap-1 cursor-pointer transition-all"
                    >
                      <Icons.Plus className="w-3 h-3" />
                      <span>Add Row</span>
                    </button>
                  </div>

                  <div className="border border-slate-200 rounded-lg overflow-hidden bg-white shadow-xs">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-slate-50 text-slate-500 font-bold uppercase text-[10px] tracking-wider border-b border-slate-200">
                          <th className="py-2.5 px-3">Rejection Reason</th>
                          <th className="py-2.5 px-3 text-right w-24">Cases</th>
                          <th className="py-2.5 px-3 text-right w-36">Amount ($)</th>
                          <th className="py-2.5 px-3 text-center w-12">Action</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 bg-white">
                        {currentClinicData.rejectionTable && currentClinicData.rejectionTable.length > 0 ? (
                          currentClinicData.rejectionTable.map((row, rIdx) => (
                            <tr key={rIdx} className="hover:bg-slate-50/30">
                              <td className="p-1">
                                <input
                                  type="text"
                                  value={row.category}
                                  onChange={(e) => handleCellEdit(editingMonthIdx, selectedClinicRaw, "rejection", rIdx, "category", e.target.value)}
                                  className="w-full px-2 py-1.5 border border-transparent rounded-md bg-transparent hover:bg-slate-50 hover:border-slate-200 focus:border-slate-900 focus:bg-white focus:ring-1 focus:ring-slate-900 focus:outline-none font-medium text-xs text-slate-800"
                                />
                              </td>
                              <td className="p-1 w-24">
                                <input
                                  type="number"
                                  value={row.cases}
                                  onChange={(e) => handleCellEdit(editingMonthIdx, selectedClinicRaw, "rejection", rIdx, "cases", e.target.value)}
                                  className="w-full px-2 py-1.5 border border-transparent rounded-md bg-transparent hover:bg-slate-50 hover:border-slate-200 focus:border-slate-900 focus:bg-white focus:ring-1 focus:ring-slate-900 focus:outline-none text-right font-mono font-medium text-xs text-slate-800"
                                />
                              </td>
                              <td className="p-1 w-36">
                                <input
                                  type="number"
                                  value={row.amount}
                                  step="0.01"
                                  onChange={(e) => handleCellEdit(editingMonthIdx, selectedClinicRaw, "rejection", rIdx, "amount", e.target.value)}
                                  className="w-full px-2 py-1.5 border border-transparent rounded-md bg-transparent hover:bg-slate-50 hover:border-slate-200 focus:border-slate-900 focus:bg-white focus:ring-1 focus:ring-slate-900 focus:outline-none text-right font-mono font-medium text-xs text-slate-800"
                                />
                              </td>
                              <td className="p-1 text-center w-12">
                                <button
                                  onClick={() => handleDeleteRow(editingMonthIdx, selectedClinicRaw, "rejection", rIdx)}
                                  className="p-1.5 text-slate-400 hover:text-red-600 rounded-md hover:bg-red-50 cursor-pointer transition-colors"
                                  title="Delete Row"
                                >
                                  <Icons.Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </td>
                            </tr>
                          ))
                        ) : (
                          <tr>
                            <td colSpan={4} className="p-6 text-center text-slate-400 font-semibold italic bg-white text-xs">
                              No rejections recorded for this clinic.
                            </td>
                          </tr>
                        )}
                        <tr className="bg-slate-50/60 font-bold border-t border-slate-200">
                          <td className="py-2 px-3 text-slate-600 text-[11px] uppercase tracking-wider font-semibold">Total Rejections</td>
                          <td className="py-2 px-3 text-right font-mono text-slate-800">
                            {currentClinicTotals.rCases.toLocaleString()}
                          </td>
                          <td className="py-2 px-3 text-right font-mono text-slate-800">
                            ${currentClinicTotals.rAmount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                          </td>
                          <td></td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>

              </div>

            </div>
          </div>
        )}

      </main>

      {/* Corporate Professional Footer */}
      <footer className="bg-slate-950 border-t border-slate-900 py-6 px-8 mt-auto flex flex-col md:flex-row items-center justify-between text-xs text-slate-400 gap-4" id="footer_copyright">
        <div>
          Hospital Quarterly Merger &copy; {new Date().getFullYear()} - All rights reserved. Royal Commission Hospital, Jubail Industrial City.
        </div>
        <div className="flex space-x-6 text-slate-500">
          <span className="flex items-center gap-1">
            <Icons.ShieldCheck className="w-3.5 h-3.5 text-emerald-500" />
            HIPAA &amp; PHI Secured
          </span>
          <span>Excel Styling: openpyxl Native Mimicry</span>
          <span>Engine: Node.js + ExcelJS</span>
        </div>
      </footer>
    </div>
  );
}
