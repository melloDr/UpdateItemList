using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using SixLabors.ImageSharp.ColorSpaces.Conversion;

namespace UpdateItemList.Utils
{
    public class GoogleSheetsHelper
    {
        private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private static readonly string ApplicationName = "itemList_v01";
        private static readonly string CredentialsPath = "D:\\User data\\Longnpt Data\\WorkSpace\\C#Winform\\UpdateItemList\\UpdateItemList\\client_secret\\credentials.json";

        SheetsService service { get; set; }

        internal async Task<SheetsService> GetSheetsService()
        {
            try
            {
                UserCredential clientSecrets;
                using (var stream = new FileStream(CredentialsPath, FileMode.Open, FileAccess.Read))
                {
                    clientSecrets = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                                GoogleClientSecrets.FromStream(stream).Secrets,
                                Scopes, // Use Sheets scope instead of Drive scope
                                "user",
                                CancellationToken.None);
                }

                return new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = clientSecrets,
                    ApplicationName = ApplicationName
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error initializing Sheets service: {ex.Message}");
                throw;
            }
        }
        public static UserCredential Login(string googleClientId, string googleClientSecret, string[] scopes)
        {
            ClientSecrets secrets = new ClientSecrets()
            {
                ClientId = googleClientId,
                ClientSecret = googleClientSecret
            };

            return GoogleWebAuthorizationBroker.AuthorizeAsync(secrets, scopes, "user", CancellationToken.None).Result;
        }

        // Method to verify if spreadsheet exists
        internal async Task<bool> SpreadsheetExists(string spreadsheetId)
        {
            try
            {
                if (service == null)
                    service = await GetSheetsService();

                var request = service.Spreadsheets.Get(spreadsheetId);
                var response = await request.ExecuteAsync();
                return response != null;
            }
            catch (Google.GoogleApiException ex) when (ex.HttpStatusCode == System.Net.HttpStatusCode.NotFound)
            {
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error checking spreadsheet: {ex.Message}");
                return false;
            }
        }
        // Method to verify if sheet exists within spreadsheet
        internal async Task<bool> SheetExists(string spreadsheetId, string sheetName)
        {
            try
            {
                if (service == null)
                    service = await GetSheetsService();

                var request = service.Spreadsheets.Get(spreadsheetId);
                var response = await request.ExecuteAsync();

                return response.Sheets.Any(sheet =>
                    string.Equals(sheet.Properties.Title, sheetName, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error checking sheet: {ex.Message}");
                return false;
            }
        }

        // Method to create a new sheet if it doesn't exist
        internal async Task<bool> CreateSheetIfNotExists(string spreadsheetId, string sheetName)
        {
            try
            {
                if (await SheetExists(spreadsheetId, sheetName))
                    return true;

                if (service == null)
                    service = await GetSheetsService();

                var addSheetRequest = new AddSheetRequest()
                {
                    Properties = new SheetProperties()
                    {
                        Title = sheetName
                    }
                };

                var batchUpdateRequest = new BatchUpdateSpreadsheetRequest()
                {
                    Requests = new List<Request>()
                    {
                        new Request()
                        {
                            AddSheet = addSheetRequest
                        }
                    }
                };

                var request = service.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadsheetId);
                await request.ExecuteAsync();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating sheet: {ex.Message}");
                return false;
            }
        }

        internal async Task<string> WriteDataTableToSheet(string spreadsheetId, string sheetName, DataTable dataTable)
        {
            try
            {
                // Validate inputs
                if (string.IsNullOrEmpty(spreadsheetId))
                {
                    return "Spreadsheet ID cannot be empty";
                }

                if (string.IsNullOrEmpty(sheetName))
                {
                    return "Sheet name cannot be empty";
                }

                if (dataTable == null || dataTable.Rows.Count == 0)
                {
                    return "DataTable is empty or null";
                }

                // Initialize service
                service = await GetSheetsService();

                // Check if spreadsheet exists
                if (!await SpreadsheetExists(spreadsheetId))
                {
                    return $"Spreadsheet with ID '{spreadsheetId}' not found. Please check the spreadsheet ID and ensure you have access to it.";
                }

                // Create sheet if it doesn't exist
                if (!await CreateSheetIfNotExists(spreadsheetId, sheetName))
                {
                    return $"Failed to create or access sheet '{sheetName}'";
                }

                // Clear existing data first
                var clearRequest = service.Spreadsheets.Values.Clear(new ClearValuesRequest(), spreadsheetId, $"{sheetName}!A:Z");
                await clearRequest.ExecuteAsync();

                // Prepare data
                var valueRange = new ValueRange();
                var values = new List<IList<object>>();

                // Add headers
                var header = new List<object>();
                foreach (DataColumn column in dataTable.Columns)
                {
                    header.Add(column.ColumnName ?? "");
                }
                values.Add(header);

                // Add data rows
                foreach (DataRow row in dataTable.Rows)
                {
                    var data = new List<object>();
                    foreach (var item in row.ItemArray)
                    {
                        data.Add(item?.ToString() ?? "");
                    }
                    values.Add(data);
                }

                valueRange.Values = values;

                // Define range starting from A1
                string range = $"{sheetName}!A1";

                // Execute update request
                var updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

                var response = await updateRequest.ExecuteAsync();

                MessageBox.Show($"Successfully updated {response.UpdatedRows} rows and {response.UpdatedColumns} columns in sheet '{sheetName}'");
                return null; // Success
            }
            catch (Google.GoogleApiException gex)
            {
                string errorMessage = $"Google API Error: {gex.Message}";
                if (gex.HttpStatusCode == System.Net.HttpStatusCode.NotFound)
                {
                    errorMessage += "\n\nPossible causes:\n" +
                                  "1. Spreadsheet ID is incorrect\n" +
                                  "2. Spreadsheet doesn't exist\n" +
                                  "3. You don't have access to the spreadsheet\n" +
                                  "4. Sheet name is incorrect";
                }
                else if (gex.HttpStatusCode == System.Net.HttpStatusCode.Forbidden)
                {
                    errorMessage += "\n\nPermission denied. Make sure:\n" +
                                  "1. You have edit access to the spreadsheet\n" +
                                  "2. Your OAuth credentials have the correct scopes";
                }

                MessageBox.Show(errorMessage);
                return gex.Message;
            }
            catch (Exception ex)
            {
                string errorMessage = $"Unexpected error: {ex.Message}";
                MessageBox.Show(errorMessage);
                return ex.Message;
            }
        }

        // Helper method to extract spreadsheet ID from URL
        public static string ExtractSpreadsheetId(string urlOrId)
        {
            if (string.IsNullOrEmpty(urlOrId))
                return urlOrId;

            // If it's already just an ID (no slashes), return as is
            if (!urlOrId.Contains("/"))
                return urlOrId;

            // Extract ID from Google Sheets URL
            var parts = urlOrId.Split('/');
            for (int i = 0; i < parts.Length; i++)
            {
                if (parts[i] == "d" && i + 1 < parts.Length)
                {
                    return parts[i + 1];
                }
            }

            return urlOrId; // Return original if extraction fails
        }
    }
}