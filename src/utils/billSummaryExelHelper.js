const Excel = require("exceljs");
const PDFDocument = require("pdfkit");
const fs = require("fs");
const { formatDateDDMMYYYY } = require("./utils");

const generateExcel = async (arr, fromDate, toDate) => {
  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet("Sheet1");

  let debitRow = 4;
  let creditRow = 4;
  let creditRowTotal = 0;
  let debitRowTotal = 0;
  worksheet.mergeCells(`A1:H1`);
  worksheet.getCell("A1").value = "Company Name";
  worksheet.mergeCells(`A2:H2`);
  worksheet.getCell("A2").value = `Bill Summary From ${formatDateDDMMYYYY(
    fromDate
  )} To ${formatDateDDMMYYYY(toDate)}`;

  worksheet.addRow([
    "Party Name",
    "Gross MTM",
    "Brok",
    "Credit Amt",
    "Party Name",
    "Gross MTM",
    "Brok",
    "Debit Amt",
  ]);

  arr.forEach((entry) => {
    if (entry.netCreditAmount) {
      worksheet.getCell(`A${creditRow}`).value = entry.partyName;
      worksheet.getCell(`B${creditRow}`).value = entry.gross_MTM.toLocaleString(
        "hi-IN",
        { style: "decimal" }
      );
      worksheet.getCell(`C${creditRow}`).value = entry.brokAmt;
      worksheet.getCell(`D${creditRow}`).value =
        entry.netCreditAmount.toLocaleString("hi-IN", { style: "decimal" });
      creditRow++;
      creditRowTotal = creditRowTotal + entry.netCreditAmount;
    } else if (entry.netDebitAmount) {
      worksheet.getCell(`E${debitRow}`).value = entry.partyName;
      worksheet.getCell(`F${debitRow}`).value = entry.gross_MTM.toLocaleString(
        "hi-IN",
        { style: "decimal" }
      );
      worksheet.getCell(`G${debitRow}`).value = entry.brokAmt;
      worksheet.getCell(`H${debitRow}`).value =
        entry.netDebitAmount.toLocaleString("hi-IN", { style: "decimal" });
      debitRow++;
      debitRowTotal = debitRowTotal + entry.netDebitAmount;
    }
  });

  let formatLimit = creditRow > debitRow ? creditRow : debitRow;
  console.log("formatLimit================>", formatLimit);

  worksheet.getCell(`D${formatLimit}`).value = creditRowTotal.toLocaleString(
    "hi-IN",
    { style: "decimal" }
  );
  worksheet.getCell(`H${formatLimit}`).value = debitRowTotal.toLocaleString(
    "hi-IN",
    { style: "decimal" }
  );

  const columnWidths = {
    A: 25,
    B: 18,
    C: 15,
    D: 18,
    E: 25,
    F: 18,
    G: 15,
    H: 18,
  };

  for (const [column, width] of Object.entries(columnWidths)) {
    worksheet.getColumn(column).alignment = { horizontal: "center" };
    worksheet.getColumn(column).width = width;
  }

  const blueStyle = {
    font: { color: { argb: "0000FF" } },
    alignment: { horizontal: "center" },
  };
  const redStyle = {
    font: { color: { argb: "FF0000" } },
    alignment: { horizontal: "center" },
  };

  for (let i = 3; i < formatLimit; i++) {
    const row = worksheet.getRow(i);
    row.eachCell({ includeEmpty: false }, function (cell, colNumber) {
      if (colNumber >= 1 && colNumber <= 4) {
        cell.style = blueStyle;
      } else if (colNumber >= 5 && colNumber <= 8) {
        cell.style = redStyle;
      }
    });
  }
  for (const [column, width] of Object.entries(columnWidths)) {
    worksheet.getCell(`${column}${formatLimit}`).border = {
      top: { style: "thick" },
      bottom: { style: "thick" },
    };
  }

  await workbook.xlsx.writeFile("output.xlsx");
  const buffer = await workbook.xlsx.writeBuffer();
  return buffer;
};

const generatePDF = async (arr) => {
  const doc = new PDFDocument();
  doc.pipe(fs.createWriteStream("output.pdf"));

  // Set up fonts
  doc.font("Helvetica-Bold");

  // Set up column headers
  doc.fillColor("blue").fontSize(9);
  doc.text("Party Name", 20, 50, { width: 85 });
  doc.text("Gross MTM", 140, 50, { width: 70 });
  doc.text("Brok", 210, 50, { width: 50 });
  doc.text("Credit Amt", 245, 50, { width: 70 });
  doc.fillColor("red").fontSize(9);
  doc.text("Party Name", 320, 50, { width: 85 });
  doc.text("Gross MTM", 440, 50, { width: 70 });
  doc.text("Brok", 500, 50, { width: 50 });
  doc.text("Credit Amt", 535, 50, { width: 70 });

  // Set up rows
  let debitRow = 60;
  let creditRow = 60;

  arr.forEach((entry) => {
    if (entry.netCreditAmount) {
      doc.fillColor("blue").fontSize(8);
      doc.text(entry.partyName, 20, (creditRow += 20), { width: 85 });
      doc.text(entry.gross_MTM.toString(), 140, creditRow, { width: 70 });
      doc.text(entry.brokAmt.toString(), 210, creditRow, { width: 50 });
      doc.text(entry.netCreditAmount.toString(), 245, creditRow, {
        width: 70,
      });
    } else if (entry.netDebitAmount) {
      doc.fillColor("red").fontSize(8);
      doc.text(entry.partyName, 320, (debitRow += 20), { width: 90 });
      doc.text(entry.gross_MTM.toString(), 440, debitRow, { width: 70 });
      doc.text(entry.brokAmt.toString(), 500, debitRow, { width: 50 });
      doc.text(entry.netDebitAmount.toString(), 535, debitRow, { width: 70 });
    }
  });

  // Finalize PDF
  doc.end();
};

module.exports = { generateExcel, generatePDF };
