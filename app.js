const fs = require('fs');
const pdfparse = require('pdf-parse');
const XLSX = require('xlsx');

const pdffile = fs.readFileSync('one.pdf');

pdfparse(pdffile).then(function (data) {
  console.log('Number of Pages:', data.numpages);
  const extractedDates = [];

  var dateAns;
  var totalAns;
  var docNo;
  var ToTitle;
  var remarks;
  // iteration over the whole string of the pdf and search for the required field mentioned above
for(let i=0;i<data.text.length-20;i++)
{
  // conditon for To:Title using the logic that is it followed by email and ends at char ']' 
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
   console.log(ToTitle);
  }
    //For Dated just check whether the string Dated is present or not 
  if(data.text.substr(i,5)=="Dated")
  {
    dateAns = data.text.substr(i,18);
    console.log(dateAns);
  }
  // For Grand Total it is after the string Amount in words 
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
// Similar logic is for Doc No.
    if(data.text.substr(i,7)=="Doc No.")
    {
      i=i+7
      docNo = data.text.substr(i,14)
      // Syntax for exporting everything to excel
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
