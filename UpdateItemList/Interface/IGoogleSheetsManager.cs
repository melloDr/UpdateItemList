using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Sheets.v4.Data;

namespace UpdateItemList.Interface
{
    public interface IGoogleSheetsManager
    {
        Spreadsheet CreateNew(string documentName);
        Spreadsheet WriteData(string sprspreadsheetId, string sheetName, DataTable dataTable);
        
    }
}
