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
  var docnopattern = /([A-Z]{9})+(\d{6})/;

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
let match1;
var i =0;
while ((match1 = docnopattern.exec(data.text)) !== null) {
  const currentDate = match1[0];
  console.log("Remarks: " + currentDate);
  extractedDates.push({ Remarks: currentDate });
  i = i + 1;
  if (i >= data.numpages) {
    break;
  }
}
const regex = /CREDIT NOTE/;
var j =0;

while(data.text.match(regex))
{
  var match2 = data.text.match(regex);
  var count = data.text.indexOf(match2[0]);
  while(data.text[count]!='.')
{
  count++;
}
count+=3;
var ans="";
while(data.text[count]!=']')
{
  count++;
ans+=data.text[count];
}
console.log(ans);
extractedDates.push({ To : ans });
var count = data.text.indexOf(match2[0]);
j++;
if(j>data.numpages)
{
  break;
}
}
var regex2 = /Amount in words/;
var match2 = data.text.match(regex2);
var j=0;
while(data.text.match(regex2)){
  var match2 = data.text.match(regex2);
var count = data.text.indexOf(match2[0])

j++;
count+=18;
var ans=""
while(data.text[count]!='.')
{
  ans+=data.text[count]
  // console.log(data.text[count]);
  count++;
}
console.log(ans);
extractedDates.push({ GrandTotal : ans });
if(j>data.numpages)
{
  break;
}
}

const extractedDates2 = [];
  // Export extracted dates to Excel
  const ws = XLSX.utils.json_to_sheet(extractedDates);
  const ws1 = XLSX.utils.json_to_sheet(extractedDates2);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet 1');
  XLSX.utils.book_append_sheet(wb, ws1, 'Sheet 2');

  XLSX.writeFile(wb, 'extracted_dates.xlsx');

  console.log('Excel file "extracted_dates.xlsx" created successfully.');
});
