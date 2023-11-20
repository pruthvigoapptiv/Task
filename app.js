const fs = require('fs');
const pdfparse = require('pdf-parse');
const XLSX = require('xlsx');

const pdffile = fs.readFileSync('one.pdf');

pdfparse(pdffile).then(function (data) {
  // Display the number of pages
  console.log('Number of Pages:', data.numpages);

  // Display the text content
  console.log('Text Content:', data.text);

  const dateRegex = /Dated\s+:(\d{2}-\d{2}-\d{4})/;
  var docnopattern = /(\d+\/\d+\/[A-Z]\d+-\d+)/;

  let match;
  var i = 0;
  const extractedDates = [];

  while ((match = dateRegex.exec(data.text)) !== null) {
    const currentDate = match[1];
    console.log("Dated: " + currentDate);
    extractedDates.push({ Dated: currentDate });
    i = i + 1;
    if (i >= data.numpages) {
      break;
    }
  }

  // Export extracted dates to Excel
  const ws = XLSX.utils.json_to_sheet(extractedDates);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet 1');
  XLSX.writeFile(wb, 'extracted_dates.xlsx');

  console.log('Excel file "extracted_dates.xlsx" created successfully.');
});
