using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;

namespace UpdateItemList.Utils
{
    public class GoogleSheetsHelper
    {
        private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private static readonly string ApplicationName = "ItemList_v01";
        private static readonly string CredentialsPath = "D:\\User data\\Longnpt Data\\WorkSpace\\C#Winform\\UpdateItemList\\UpdateItemList\\client_secret\\credentials.json"; // Đường dẫn đến file JSON đã tải về

        SheetsService service { get; set; }
        internal async Task<SheetsService> GetSheetsService()
        {
            GoogleCredential credential;


            //using (var stream = new FileStream(CredentialsPath, FileMode.Open, FileAccess.Read))
            //{
            //    credential = GoogleCredential.FromStream(stream)
            //        .CreateScoped(Scopes);
            //}
            UserCredential clientSecrets ;
            using (var stream = new FileStream(CredentialsPath, FileMode.Open, FileAccess.Read))
            {
                clientSecrets = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                            GoogleClientSecrets.FromStream(stream).Secrets,
                            new[] { DriveService.ScopeConstants.DriveFile },
                            "user",
                            CancellationToken.None);

            } ;


            return new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = clientSecrets,
                ApplicationName = ApplicationName
            });
        }
        internal async Task WriteDataTableToSheet(string spreadsheetId, string sheetName, DataTable dataTable)
        {
            try
            {
                service = await GetSheetsService();
                // Chuẩn bị dữ liệu
                var valueRange = new ValueRange();
                var values = new List<IList<object>>();

                // Thêm tiêu đề
                var header = new List<object>();
                foreach (DataColumn column in dataTable.Columns)
                {
                    header.Add(column.ColumnName);
                }
                values.Add(header);

                // Thêm dữ liệu từ DataTable
                foreach (DataRow row in dataTable.Rows)
                {
                    var data = new List<object>();
                    foreach (var item in row.ItemArray)
                    {
                        data.Add(item.ToString());
                    }
                    values.Add(data);
                }

                // Định dạng phạm vi (ví dụ: Sheet1!A1)
                string range = $"{sheetName}!A1";

                valueRange.Values = values;

                // Gửi yêu cầu cập nhật
                var request = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                var check = request.Execute();
                string a = "";
                string aa = "";
                string aaa = a + aa;
                MessageBox.Show(aaa);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã có lỗi xảy ra: " + ex.ToString());
            }
            
        }

    }
}
