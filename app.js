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
  var docNo;
  var ToTitle;
  var remarks;
for(let i=0;i<data.text.length-20;i++)
{
  if(data.text.substr(i,7)=="Message")
  {
    i=i+9
    remarks=data.text.substr(i,15);
  }
  if(data.text[i]==']')
  {
  var k =i;
   for( k=i;k>=10;k--)
   {
    if(data.text[k]=='m')
    {
      break;
    }
   }
   k=k+2
   var temp=""
   for(k=k;k<k+60;k++)
   {
    temp+=data.text[k]
    if(data.text[k]==']')
    {
      
      break;
    }
   }
   ToTitle=temp;
  }
if(data.text[i]=='D' && data.text[i+1]=='a' && data.text[i+2]=='t' && data.text[i+4]=='d')
{
  dateAns = data.text.substr(i,18);
}

if(data.text.substr(i,15)=="Amount in words")
{
  var temp = "";
  i=i+18;
  while(data.text[i]!='.')
  {
    temp+=data.text[i]
    i++;
  }
  totalAns=temp;


}
if(data.text.substr(i,7)=="Doc No.")
{
  i=i+7
  docNo = data.text.substr(i,14)
  extractedDates.push({To: ToTitle, Dated: dateAns, GrandTotal: totalAns, DocNo : docNo, Remarks: remarks });

}

}
  // Export extracted dates to Excel
  const ws = XLSX.utils.json_to_sheet(extractedDates);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet 1');

  XLSX.writeFile(wb, 'extracted_dates.xlsx');

  console.log('Excel file "extracted_dates.xlsx" created successfully.');
});
