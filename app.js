const fs = require('fs');
const pdfparse = require('pdf-parse');
const XLSX = require('xlsx');

const pdffile = fs.readFileSync('one.pdf');

pdfparse(pdffile).then(function (data) {
  console.log('Number of Pages:', data.numpages);
  console.log('Text Content:', data.text);
  const extractedDates = [];

  var dateAns;
  var totalAns;
for(let i=0;i<data.text.length-20;i++)
{
if(data.text[i]=='D' && data.text[i+1]=='a' && data.text[i+2]=='t' && data.text[i+4]=='d')
{
  dateAns = data.text.substr(i,18);
}

if(
  data.text[i]=='A' && 
  data.text[i+1]=='m' &&
  data.text[i+2]=='o' &&
  data.text[i+7]=='i' &&
  data.text[i+8]=='n' 

  
  )
{
  var temp = "";
  while(data.text[i+18]!='.')
  {
    temp+=data.text[i+18]
    i++;
  }
  totalAns=temp;
  extractedDates.push({ Dated: dateAns, GrandTotal: totalAns });


}
}
  // Export extracted dates to Excel
  const ws = XLSX.utils.json_to_sheet(extractedDates);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet 1');

  XLSX.writeFile(wb, 'extracted_dates.xlsx');

  console.log('Excel file "extracted_dates.xlsx" created successfully.');
});
