import React, { useState } from "react";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import "bootstrap/dist/css/bootstrap.min.css";
const StatsAutomate= () => {
    const [julFile, setJulFile] = useState(null);
    const [finalFile, setFinalFile] = useState(null);
    const [month, setMonth] = useState("");
    const [isProcessing, setIsProcessing] = useState(false);
    const handleJulFileChange = (e) => {
       setJulFile(e.target.files[0]);
     };
     const handleFinalFileChange = (e) => {
       setFinalFile(e.target.files[0]);
     };
     const handleMonthChange = (e) => {
       setMonth(e.target.value);
     };
     const processFiles = async () => {
      if (!julFile || !finalFile || !month) {
        alert("Please upload both files and select a month.");
        return;
      }
      setIsProcessing(true);
      try {
        // Read workbooks
        const julWorkbook = new ExcelJS.Workbook();
        await julWorkbook.xlsx.load(await julFile.arrayBuffer());
        const finalWorkbook = new ExcelJS.Workbook();
        await finalWorkbook.xlsx.load(await finalFile.arrayBuffer());
        const sheetRowMapping = {
          "Ad copy QC": [2, 14],
          "Retail Ad copy QC": [15, 27],
          "LA Uploads": [28, 36],
          "Coding and Uploads": [37, 45],
          "Fulfillment Digital": [46, 54],
          Screenshot: [55, 63],
          "Enterprise QC": [64, 76],
          "Enterprise Uploads": [77, 85],
          DR: [86, 94],
          "Amp DR": [95, 103],
          "Amp OE": [104, 112],
          ROE: [113, 121],
          QAR: [122, 130],
          MG: [131, 140],
          ADM: [141, 149],
          "MR-AA Reporting": [150, 158],
          Uti: [159, 159],
          Attendance: [160, 162],
          "PKT Scores": [164, 164],
        };
        
        // Process each sheet
        for (const [sheetName, [startRow, endRow]] of Object.entries(sheetRowMapping)) {
          const julSheet = julWorkbook.getWorksheet(sheetName);
          if (!julSheet) continue;
          // Find Emp ID column (1-based index)
          const headerRow = julSheet.getRow(1);
          let empIdCol = null;
          headerRow.eachCell((cell, colNumber) => {
            if (cell.text?.trim()?.toLowerCase() === "emp id") {
              empIdCol = colNumber;
            }
          });
          if (!empIdCol) continue;
          // Create EmpID to row mapping
          const julData = new Map();
          julSheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) { // Skip header
              const empId = row.getCell(empIdCol).text?.trim()?.toUpperCase();
              if (empId) julData.set(empId, rowNumber);
            }
          });
          // Process final workbook sheets
          finalWorkbook.eachSheet((finalSheet) => {
            const finalEmpIdCell = finalSheet.getRow(1).getCell(1);
            if (!finalEmpIdCell) return;
            const finalEmpId = finalEmpIdCell.text?.trim()?.toUpperCase();
            if (!finalEmpId || !julData.has(finalEmpId)) return;
            const julRowNumber = julData.get(finalEmpId);
            // Calculate target column (month)
            const targetCol = parseInt(month); // Directly use month value for 1-based column
            for (let rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
              const julRow = julSheet.getRow(julRowNumber);
              const finalRow = finalSheet.getRow(rowIdx);
              // Calculate source column (1-based)
              const sourceCol = (rowIdx - startRow) + 3; // Adjusted for 1-based index
              // Get cells
              const sourceCell = julRow.getCell(sourceCol);
              const targetCell = finalRow.getCell(targetCol);
              // Copy value
              targetCell.value = sourceCell.value;

              if(sourceCell.fill &&  sourceCell.fill.fgColor){
                targetCell.fill = {
                    type : "pattern",
                    pattern : "solid",
                    fgColor : {
                        arfb: sourceCell.fill.fgColor
                    },
                }
              }
              // Copy styles
              targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
              // Copy number format
              if (sourceCell.numFmt) {
                targetCell.numFmt = sourceCell.numFmt;
              }
            }
          });
        }
        // Save updated file
        const buffer = await finalWorkbook.xlsx.writeBuffer();
        saveAs(new Blob([buffer]), "Updated_Stats.xlsx");
      } catch (error) {
        console.error("Error processing files:", error);
        alert("An error occurred while processing the files.");
      } finally {
        setIsProcessing(false);
      }
    };
 return (
  <div className="container mt-5 p-4 border rounded shadow bg-light">
<h2 className="text-center fw-bold text-primary mb-4">Stats Automate</h2>
<div className="row g-3">
<div className="col-md">
<div className="input-group p-3">
<input type="file" className="form-control" accept=".xlsx" onChange={handleJulFileChange} />
<label className="input-group-text">Upload Performance File</label>
</div>
<div className="input-group p-3">
<input type="file" className="form-control" accept=".xlsx" onChange={handleFinalFileChange} />
<label className="input-group-text">Upload Stats File</label>
</div>
</div>
<div className="col-md p-3">
<div className="form-floating">
<select className="form-select" value={month} onChange={handleMonthChange}>
<option value="">Select Month</option>
<option value="2">January</option>
<option value="3">February</option>
<option value="4">March</option>
<option value="5">April</option>
<option value="6">May</option>
<option value="7">June</option>
<option value="8">July</option>
<option value="9">August</option>
<option value="10">September</option>
<option value="11">October</option>
<option value="12">November</option>
<option value="13">December</option>
</select>
<label>Choose a Month</label>
</div>
</div>
</div>
<div className="mt-4 text-center">
<button className="btn btn-lg btn-primary" onClick={processFiles} disabled={isProcessing}>
         {isProcessing ? "Processing..." : "Process and Download"}
</button>
</div>
     {isProcessing && (
<div className="text-center my-3">
<div className="spinner-border text-primary" role="status">
<span className="visually-hidden">Processing...</span>
</div>
<p>Processing...</p>
</div>
     )}
</div>
 );
}
export default StatsAutomate;
