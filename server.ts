import express from "express";
import path from "path";
import multer from "multer";
import ExcelJS from "exceljs";
import { createServer as createViteServer } from "vite";
import { CLINICS_METADATA, ExcelRecord, ClinicData, MonthlyReportPayload } from "./src/types";

const app = express();
const PORT = 3000;

// Configure body-parser limit for large payloads (since edited data might be sent back)
app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ limit: "50mb", extended: true }));

// Configure multer for memory storage
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024 } // 25MB limit
});

// Helper to clean and parse cells to number
function parseNumber(val: any): number {
  if (val === null || val === undefined) return 0;
  if (typeof val === "number") return val;
  if (typeof val === "object" && val !== null) {
    if ("result" in val) {
      return parseNumber(val.result);
    }
  }
  const str = val.toString().replace(/[\$,\s%]/g, "");
  const parsed = parseFloat(str);
  return isNaN(parsed) ? 0 : parsed;
}

// Helper to extract clean strings from cells
function parseString(val: any): string {
  if (val === null || val === undefined) return "";
  if (typeof val === "string") return val.trim();
  if (typeof val === "object" && val !== null) {
    if ("richText" in val && Array.isArray(val.richText)) {
      return val.richText.map((t: any) => t.text || "").join("").trim();
    }
    if ("result" in val) {
      return parseString(val.result);
    }
    if ("text" in val) {
      return parseString(val.text);
    }
  }
  return val.toString().trim();
}

// Helper to match clinic name with resilience
function isClinicMatch(cellValue: string, rawClinicName: string): boolean {
  const cellClean = cellValue.toLowerCase().trim().replace(/\s+/g, " ");
  const targetClean = rawClinicName.toLowerCase().trim().replace(/\s+/g, " ");
  
  if (cellClean === targetClean) return true;
  
  // Specific resilient checks for clinic segments
  if (targetClean.includes("hijaz") && cellClean.includes("hijaz")) return true;
  if (targetClean.includes("dafi") && cellClean.includes("dafi")) return true;
  if (targetClean.includes("howilat") && cellClean.includes("howilat")) return true;
  if (targetClean.includes("farouq") && cellClean.includes("farouq")) return true;
  if (targetClean.includes("jalmoud") && cellClean.includes("jalmoud")) return true;
  if (targetClean.includes("hospital in jubail") && cellClean.includes("hospital in jubail")) return true;
  if (targetClean.includes("ras al khair") && cellClean.includes("ras al khair")) return true;

  return false;
}

// Parser logic
async function parseMonthlyReport(buffer: Buffer): Promise<{ [rawName: string]: ClinicData }> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  
  // Search for the "Dashboard" sheet case-insensitively
  let sheet = workbook.getWorksheet("Dashboard") || workbook.getWorksheet("dashboard");
  if (!sheet) {
    sheet = workbook.worksheets.find(s => s.name.toLowerCase() === "dashboard") || workbook.worksheets[0];
  }
  
  const clinicsResult: { [rawName: string]: ClinicData } = {};
  for (const c of CLINICS_METADATA) {
    clinicsResult[c.rawName] = {
      statusTable: [],
      rejectionTable: []
    };
  }

  if (!sheet) {
    return clinicsResult;
  }

  // Find cell positions for all clinics
  interface FoundClinic {
    rawName: string;
    row: number;
    col: number;
  }
  const foundClinics: FoundClinic[] = [];

  sheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      const val = cell.value;
      if (val) {
        const textVal = parseString(val);
        for (const clinic of CLINICS_METADATA) {
          if (isClinicMatch(textVal, clinic.rawName)) {
            foundClinics.push({
              rawName: clinic.rawName,
              row: rowNumber,
              col: colNumber
            });
          }
        }
      }
    });
  });

  // Sort by row number to determine boundaries
  foundClinics.sort((a, b) => a.row - b.row);

  for (let i = 0; i < foundClinics.length; i++) {
    const current = foundClinics[i];
    const nextRowBoundary = i + 1 < foundClinics.length ? foundClinics[i + 1].row : sheet.rowCount + 1;
    // Scan cells below the clinic anchor
    let statusHeader: { r: number; c: number } | null = null;
    let rejectionHeader: { r: number; c: number } | null = null;

    // Search row range for tables
    const maxSearchRow = Math.min(current.row + 35, nextRowBoundary);
    for (let r = current.row + 1; r < maxSearchRow; r++) {
      for (let c = 1; c <= 20; c++) {
        const val = sheet.getCell(r, c).value;
        if (val) {
          const str = parseString(val).toLowerCase();
          
          // Identify Status Header
          if (str === "status") {
            const next1 = parseString(sheet.getCell(r, c + 1).value).toLowerCase();
            const next2 = parseString(sheet.getCell(r, c + 2).value).toLowerCase();
            if (next1.includes("case") || next2.includes("amount") || next1.includes("amt") || next2.includes("sum")) {
              statusHeader = { r, c };
            }
          }
          
          // Identify Rejection Header
          if (str.includes("rejection reason") || str.includes("rejection reasons") || str === "rejection" || str.includes("reason for rejection")) {
            const next1 = parseString(sheet.getCell(r, c + 1).value).toLowerCase();
            const next2 = parseString(sheet.getCell(r, c + 2).value).toLowerCase();
            if (next1.includes("case") || next2.includes("amount") || next1.includes("amt") || next2.includes("sum")) {
              rejectionHeader = { r, c };
            }
          }
        }
      }
    }

    // Extract Status Table
    if (statusHeader) {
      let rIdx = statusHeader.r + 1;
      while (rIdx < nextRowBoundary) {
        const cat = parseString(sheet.getCell(rIdx, statusHeader.c).value);
        if (!cat || cat.toLowerCase().startsWith("total") || cat.toLowerCase() === "sum" || cat.toLowerCase().includes("grand total")) {
          break;
        }
        const cases = parseNumber(sheet.getCell(rIdx, statusHeader.c + 1).value);
        const amount = parseNumber(sheet.getCell(rIdx, statusHeader.c + 2).value);
        clinicsResult[current.rawName].statusTable.push({ category: cat, cases, amount });
        rIdx++;
        if (rIdx > statusHeader.r + 30) break; // safeguard
      }
    }

    // Extract Rejection Table
    if (rejectionHeader) {
      let rIdx = rejectionHeader.r + 1;
      while (rIdx < nextRowBoundary) {
        const cat = parseString(sheet.getCell(rIdx, rejectionHeader.c).value);
        if (!cat || cat.toLowerCase().startsWith("total") || cat.toLowerCase() === "sum" || cat.toLowerCase().includes("grand total")) {
          break;
        }
        const cases = parseNumber(sheet.getCell(rIdx, rejectionHeader.c + 1).value);
        const amount = parseNumber(sheet.getCell(rIdx, rejectionHeader.c + 2).value);
        clinicsResult[current.rawName].rejectionTable.push({ category: cat, cases, amount });
        rIdx++;
        if (rIdx > rejectionHeader.r + 30) break; // safeguard
      }
    }
  }

  return clinicsResult;
}

// REST API endpoint to parse uploaded file
app.post("/api/parse", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file was uploaded." });
    }
    const data = await parseMonthlyReport(req.file.buffer);
    res.json({ filename: req.file.originalname, clinics: data });
  } catch (error: any) {
    console.error("Error parsing Excel:", error);
    res.status(500).json({ error: "Failed to parse file: " + error.message });
  }
});

// Helper to style spreadsheet cells
function styleHeader(cell: ExcelJS.Cell) {
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF1F497D" } // Deep Navy Blue
  };
  cell.font = {
    name: "Inter",
    color: { argb: "FFFFFFFF" },
    bold: true,
    size: 11
  };
  cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  cell.border = {
    top: { style: "thin", color: { argb: "FFD3D3D3" } },
    left: { style: "thin", color: { argb: "FFD3D3D3" } },
    bottom: { style: "medium", color: { argb: "FF1F497D" } },
    right: { style: "thin", color: { argb: "FFD3D3D3" } }
  };
}

function styleData(cell: ExcelJS.Cell, alignment: "left" | "center" | "right" = "left", isBold = false) {
  cell.font = { name: "Inter", size: 10, bold: isBold };
  cell.alignment = { vertical: "middle", horizontal: alignment };
  cell.border = {
    top: { style: "thin", color: { argb: "FFE0E0E0" } },
    left: { style: "thin", color: { argb: "FFE0E0E0" } },
    bottom: { style: "thin", color: { argb: "FFE0E0E0" } },
    right: { style: "thin", color: { argb: "FFE0E0E0" } }
  };
}

function styleTotal(cell: ExcelJS.Cell, alignment: "left" | "center" | "right" = "left") {
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFF2F2F2" } // Light Gray
  };
  cell.font = { name: "Inter", size: 10, bold: true };
  cell.alignment = { vertical: "middle", horizontal: alignment };
  cell.border = {
    top: { style: "thin", color: { argb: "FF1F497D" } },
    left: { style: "thin", color: { argb: "FFE0E0E0" } },
    bottom: { style: "double", color: { argb: "FF1F497D" } }, // Double bottom accounting line
    right: { style: "thin", color: { argb: "FFE0E0E0" } }
  };
}

// REST API endpoint to generate merged report
app.post("/api/generate", async (req, res) => {
  try {
    const { payload }: { payload: MonthlyReportPayload[] } = req.body;

    if (!payload || payload.length !== 3) {
      return res.status(400).json({ error: "Three months of report data are required." });
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Quarterly Report Merger";

    // 1. Calculate Aggregations for "Quarter Summary" sheet
    const rawStatusRecords: ExcelRecord[] = [];
    const rawRejectionRecords: ExcelRecord[] = [];

    payload.forEach(month => {
      Object.keys(month.clinics).forEach(clinicRaw => {
        const clinicData = month.clinics[clinicRaw];
        if (clinicData.statusTable) {
          rawStatusRecords.push(...clinicData.statusTable);
        }
        if (clinicData.rejectionTable) {
          rawRejectionRecords.push(...clinicData.rejectionTable);
        }
      });
    });

    // Grouping status summary
    const statusMap: { [cat: string]: { cases: number; amount: number } } = {};
    rawStatusRecords.forEach(r => {
      const catNorm = r.category.trim();
      const catKey = catNorm.toLowerCase();
      if (!catKey) return;
      if (!statusMap[catKey]) {
        statusMap[catKey] = { cases: 0, amount: 0 };
      }
      statusMap[catKey].cases += r.cases;
      statusMap[catKey].amount += r.amount;
    });

    // Make beautiful capitalized display status names
    const statusMapping: { [key: string]: string } = {
      approved: "Approved",
      rejected: "Rejected",
      pending: "Pending",
      "in progress": "In Progress",
      returned: "Returned",
      submitted: "Submitted",
      paid: "Paid"
    };

    const aggregatedStatus: ExcelRecord[] = Object.keys(statusMap).map(key => {
      const display = statusMapping[key] || (key.charAt(0).toUpperCase() + key.slice(1));
      return {
        category: display,
        cases: statusMap[key].cases,
        amount: statusMap[key].amount
      };
    });

    // Grouping rejection summary
    const rejectMap: { [cat: string]: { cases: number; amount: number } } = {};
    rawRejectionRecords.forEach(r => {
      const catNorm = r.category.trim();
      const catKey = catNorm.toLowerCase();
      if (!catKey) return;
      if (!rejectMap[catKey]) {
        rejectMap[catKey] = { cases: 0, amount: 0 };
      }
      rejectMap[catKey].cases += r.cases;
      rejectMap[catKey].amount += r.amount;
    });

    const aggregatedRejection: ExcelRecord[] = Object.keys(rejectMap).map(key => {
      return {
        category: key.charAt(0).toUpperCase() + key.slice(1),
        cases: rejectMap[key].cases,
        amount: rejectMap[key].amount
      };
    });

    // Total calculations to ensure rule validation: Status Total = Rejection Total
    const totalStatusCases = aggregatedStatus.reduce((sum, r) => sum + r.cases, 0);
    const totalStatusAmount = aggregatedStatus.reduce((sum, r) => sum + r.amount, 0);

    const totalRejectCases = aggregatedRejection.reduce((sum, r) => sum + r.cases, 0);
    const totalRejectAmount = aggregatedRejection.reduce((sum, r) => sum + r.amount, 0);

    // Apply strict matching rule: adjust rejection to match status total by adding approved/paid/other cases
    if (totalStatusCases !== totalRejectCases || totalStatusAmount !== totalRejectAmount) {
      const diffCases = totalStatusCases - totalRejectCases;
      const diffAmount = totalStatusAmount - totalRejectAmount;
      
      aggregatedRejection.push({
        category: "Approved / Paid Claims (No Rejection Reason)",
        cases: diffCases,
        amount: diffAmount
      });
    }

    // --- SHEET 1: Quarter Summary ---
    const summarySheet = workbook.addWorksheet("Quarter Summary", { views: [{ showGridLines: true }] });
    
    // Add title
    summarySheet.addRow([]);
    const titleRow = summarySheet.addRow(["HOSPITAL QUARTERLY CONSOLIDATED REPORT - SUMMARY"]);
    titleRow.getCell(1).font = { name: "Inter", size: 16, bold: true, color: { argb: "FF1F497D" } };
    summarySheet.addRow([]);

    // Table 1: Status Summary Header
    summarySheet.addRow(["Total Quarter Status Summary"]).getCell(1).font = { name: "Inter", size: 12, bold: true };
    const stHeader = summarySheet.addRow(["Status", "Total Cases", "Total Amount"]);
    stHeader.eachCell(c => styleHeader(c));

    aggregatedStatus.forEach(r => {
      const row = summarySheet.addRow([r.category, r.cases, r.amount]);
      styleData(row.getCell(1), "left");
      styleData(row.getCell(2), "right");
      row.getCell(2).numFmt = "#,##0";
      styleData(row.getCell(3), "right");
      row.getCell(3).numFmt = "$#,##0.00";
    });

    // Status Total
    const stTotalRow = summarySheet.addRow(["Total Status Claims", totalStatusCases, totalStatusAmount]);
    styleTotal(stTotalRow.getCell(1), "left");
    styleTotal(stTotalRow.getCell(2), "right");
    stTotalRow.getCell(2).numFmt = "#,##0";
    styleTotal(stTotalRow.getCell(3), "right");
    stTotalRow.getCell(3).numFmt = "$#,##0.00";

    summarySheet.addRow([]);
    summarySheet.addRow([]);

    // Table 2: Rejection Summary Header
    summarySheet.addRow(["Total Quarter Rejection Summary"]).getCell(1).font = { name: "Inter", size: 12, bold: true };
    const rjHeader = summarySheet.addRow(["Rejection Reason", "Total Cases", "Total Amount"]);
    rjHeader.eachCell(c => styleHeader(c));

    aggregatedRejection.forEach(r => {
      const row = summarySheet.addRow([r.category, r.cases, r.amount]);
      styleData(row.getCell(1), "left");
      styleData(row.getCell(2), "right");
      row.getCell(2).numFmt = "#,##0";
      styleData(row.getCell(3), "right");
      row.getCell(3).numFmt = "$#,##0.00";
    });

    // Rejection Total (Should match Status Total exactly)
    const finalRejCases = aggregatedRejection.reduce((sum, r) => sum + r.cases, 0);
    const finalRejAmount = aggregatedRejection.reduce((sum, r) => sum + r.amount, 0);
    const rjTotalRow = summarySheet.addRow(["Total Rejections & Paid", finalRejCases, finalRejAmount]);
    styleTotal(rjTotalRow.getCell(1), "left");
    styleTotal(rjTotalRow.getCell(2), "right");
    rjTotalRow.getCell(2).numFmt = "#,##0";
    styleTotal(rjTotalRow.getCell(3), "right");
    rjTotalRow.getCell(3).numFmt = "$#,##0.00";

    // Auto-fit summary sheet columns
    summarySheet.columns.forEach(col => {
      let maxLen = 15;
      col.eachCell({ includeEmpty: false }, c => {
        if (c.value) {
          const l = c.value.toString().length;
          if (l > maxLen) maxLen = l;
        }
      });
      col.width = Math.min(maxLen + 4, 60);
    });

    // --- SHEETS 2-8: Clinic Specific Sheets ---
    CLINICS_METADATA.forEach(clinic => {
      const clinicSheet = workbook.addWorksheet(clinic.shortName, { views: [{ showGridLines: true }] });
      
      clinicSheet.addRow([]);
      const clHeader = clinicSheet.addRow([`${clinic.rawName.toUpperCase()} - QUARTER DETAILS`]);
      clHeader.getCell(1).font = { name: "Inter", size: 14, bold: true, color: { argb: "FF1F497D" } };
      clinicSheet.addRow([]);

      // Stack Month 1, Month 2, Month 3 tables vertically
      payload.forEach((month, mIdx) => {
        const monthNum = mIdx + 1;
        const clinicData = month.clinics[clinic.rawName] || { statusTable: [], rejectionTable: [] };
        
        // Month Section Title
        const mTitle = clinicSheet.addRow([`MONTH ${monthNum}: ${month.dateRange || "N/A"} (Insurance: ${month.insuranceCompany || "N/A"})`]);
        mTitle.getCell(1).font = { name: "Inter", size: 11, bold: true, italic: true, color: { argb: "FF1F497D" } };
        clinicSheet.addRow([]);

        // --- STATUS TABLE ---
        const sSub = clinicSheet.addRow(["Status Table"]);
        sSub.getCell(1).font = { name: "Inter", size: 10, bold: true };
        const sHead = clinicSheet.addRow(["Status", "Cases", "Amount"]);
        sHead.eachCell(c => styleHeader(c));

        let sCasesTotal = 0;
        let sAmountTotal = 0;

        if (clinicData.statusTable && clinicData.statusTable.length > 0) {
          clinicData.statusTable.forEach(r => {
            const row = clinicSheet.addRow([r.category, r.cases, r.amount]);
            styleData(row.getCell(1), "left");
            styleData(row.getCell(2), "right");
            row.getCell(2).numFmt = "#,##0";
            styleData(row.getCell(3), "right");
            row.getCell(3).numFmt = "$#,##0.00";

            sCasesTotal += r.cases;
            sAmountTotal += r.amount;
          });
        } else {
          // Empty indicator row styled gracefully
          const row = clinicSheet.addRow(["No data recorded", 0, 0]);
          styleData(row.getCell(1), "left", true);
          styleData(row.getCell(2), "right");
          styleData(row.getCell(3), "right");
        }

        // Status Table Total
        const sTot = clinicSheet.addRow(["Total Status", sCasesTotal, sAmountTotal]);
        styleTotal(sTot.getCell(1), "left");
        styleTotal(sTot.getCell(2), "right");
        sTot.getCell(2).numFmt = "#,##0";
        styleTotal(sTot.getCell(3), "right");
        sTot.getCell(3).numFmt = "$#,##0.00";

        clinicSheet.addRow([]);

        // --- REJECTION TABLE ---
        const rSub = clinicSheet.addRow(["Rejection Table"]);
        rSub.getCell(1).font = { name: "Inter", size: 10, bold: true };
        const rHead = clinicSheet.addRow(["Rejection Reason", "Cases", "Amount"]);
        rHead.eachCell(c => styleHeader(c));

        let rCasesTotal = 0;
        let rAmountTotal = 0;

        if (clinicData.rejectionTable && clinicData.rejectionTable.length > 0) {
          clinicData.rejectionTable.forEach(r => {
            const row = clinicSheet.addRow([r.category, r.cases, r.amount]);
            styleData(row.getCell(1), "left");
            styleData(row.getCell(2), "right");
            row.getCell(2).numFmt = "#,##0";
            styleData(row.getCell(3), "right");
            row.getCell(3).numFmt = "$#,##0.00";

            rCasesTotal += r.cases;
            rAmountTotal += r.amount;
          });
        } else {
          // Empty indicator row styled gracefully
          const row = clinicSheet.addRow(["No rejections recorded", 0, 0]);
          styleData(row.getCell(1), "left", true);
          styleData(row.getCell(2), "right");
          styleData(row.getCell(3), "right");
        }

        // Rejection Table Total
        const rTot = clinicSheet.addRow(["Total Rejections", rCasesTotal, rAmountTotal]);
        styleTotal(rTot.getCell(1), "left");
        styleTotal(rTot.getCell(2), "right");
        rTot.getCell(2).numFmt = "#,##0";
        styleTotal(rTot.getCell(3), "right");
        rTot.getCell(3).numFmt = "$#,##0.00";

        clinicSheet.addRow([]);
        clinicSheet.addRow([]); // Double spacing between months
      });

      // Auto column widths for clinic sheet
      clinicSheet.columns.forEach(col => {
        let maxLen = 15;
        col.eachCell({ includeEmpty: false }, c => {
          if (c.value) {
            const l = c.value.toString().length;
            if (l > maxLen) maxLen = l;
          }
        });
        col.width = Math.min(maxLen + 4, 60);
      });
    });

    // Write Excel to buffer and send it back
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=Quarterly_Report.xlsx");
    
    await workbook.xlsx.write(res);
    res.end();
  } catch (error: any) {
    console.error("Error generating Excel:", error);
    res.status(500).json({ error: "Failed to generate report: " + error.message });
  }
});

// Configure Vite integration for dev server or static file hosting for production
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa"
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
