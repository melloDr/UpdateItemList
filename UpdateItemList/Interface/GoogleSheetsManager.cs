using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using NPOI.POIFS.FileSystem;

namespace UpdateItemList.Interface
{
    public class GoogleSheetsManager : IGoogleSheetsManager
    {
        private readonly UserCredential _credential;
        public GoogleSheetsManager(UserCredential credential)
        {
            _credential = credential;
        }
        public Spreadsheet CreateNew(string documentName)
        {
            if (string.IsNullOrEmpty(documentName))
                throw new ArgumentNullException(nameof(documentName));

            using (var sheetsService = new SheetsService(new BaseClientService.Initializer() { HttpClientInitializer = _credential }))
            {
                var documentCreationRequest = sheetsService.Spreadsheets.Create(new Spreadsheet()
                {
                    Sheets = new List<Sheet>()
                    {
                        new Sheet()
                        {
                            Properties = new SheetProperties()
                            {
                                Title = documentName
                            }
                        }
                    },

                    Properties = new SpreadsheetProperties()
                    {
                        Title = documentName
                    }
                });

                return documentCreationRequest.Execute();
            }
        }

        public Spreadsheet WriteData(string sprspreadsheetId, string sheetName, DataTable dataTable)
        {
            using (var sheetsService = new SheetsService(new BaseClientService.Initializer() { HttpClientInitializer = _credential }))
            {
                var documentCreationRequest = sheetsService.Spreadsheets.Get(sprspreadsheetId);
                return documentCreationRequest.Execute();
            }
        }
    }
}
