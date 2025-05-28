using System.Data;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UpdateItemList.Interface;
using UpdateItemList.Utils;

namespace UpdateItemList
{
    public partial class Form1 : Form
    {
        HandleDatatable handleDatatable = new HandleDatatable();

        DataTable dataTable = new DataTable();
        public Form1()
        {
            InitializeComponent();
        }

        #region handleExcelData

        private DataTable ReadExcelData(string filePath, int i, ref DataTable dataTable)
        {

            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook workbook = new XSSFWorkbook(fileStream);
                ISheet sheet = workbook.GetSheetAt(0); // Lấy sheet đầu tiên

                int solutionRowIndex = -1;
                int startCol = 1; // Cột B (Item Type)
                int endCol = 13; // Cột N



                // Tìm dòng chứa "Solution" trong cột B (Item Type)
                for (int rowIdx = 5; rowIdx <= sheet.LastRowNum; rowIdx++)
                {
                    IRow row = sheet.GetRow(rowIdx);
                    if (row?.GetCell(1)?.ToString().Trim().Equals("Solution", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        solutionRowIndex = rowIdx;
                        break;
                    }
                }

                if (solutionRowIndex == -1)
                    throw new Exception("Không tìm thấy ô 'Solution'");

                if (i == 0)
                {
                    // Tạo tiêu đề DataTable từ dòng 5(cột B đến N)
                    IRow headerRow = sheet.GetRow(4);
                    for (int col = startCol; col <= endCol; col++)
                    {
                        string header = headerRow.GetCell(col)?.ToString() ?? $"Column{col}";
                        dataTable.Columns.Add(header.Trim());
                    }
                }

                // Thêm dữ liệu từ dòng Solution đến cuối
                IFormulaEvaluator evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
                for (int rowIdx = solutionRowIndex; rowIdx <= sheet.LastRowNum; rowIdx++)
                {
                    IRow currentRow = sheet.GetRow(rowIdx);
                    if (currentRow == null) continue;

                    DataRow dataRow = dataTable.NewRow();
                    bool isEmptyRow = true;

                    for (int col = startCol; col <= endCol; col++)
                    {
                        ICell cell = currentRow.GetCell(col);
                        string cellValue = handleDatatable.GetCellValue(cell, evaluator); // Sử dụng hàm mới

                        //Xử lý đặc biệt cho cột Date Created
                        if (dataTable.Columns[col - startCol].ColumnName == "Date Created" && double.TryParse(cellValue, out double oaDate))
                        {
                            dataRow[col - startCol] = DateTime.FromOADate(oaDate).ToString("dd/MM/yyyy");
                            isEmptyRow = false;
                        }
                        else
                        {
                            dataRow[col - startCol] = cellValue.Trim();
                            if (!string.IsNullOrEmpty(cellValue)) isEmptyRow = false;
                        }
                    }

                    if (!isEmptyRow) dataTable.Rows.Add(dataRow);
                }
            }
            return dataTable;
        }

        #endregion

        private void btnFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

            folderBrowserDialog.Description = "Chọn một thư mục:";
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                txtFolder.Text = folderBrowserDialog.SelectedPath;
            }
            string selectedFolderPath = txtFolder.Text;
            if (string.IsNullOrEmpty(selectedFolderPath)) return;

            string[] excelFiles = Directory.GetFiles(selectedFolderPath, "*.xls*")
                                             .Where(file => file.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                                                            file.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                                                            file.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase) ||
                                                            file.EndsWith(".xlsb", StringComparison.OrdinalIgnoreCase))
                                             .ToArray();

            for (int i = 0; i < excelFiles.Length; i++)
            {
                dataTable = ReadExcelData(excelFiles[i], i, ref dataTable);
            }
            int insertPosition = dataTable.Columns.Count - 1;
            DataColumn newColumn = new DataColumn("DateDelete", typeof(string));
            dataTable.Columns.Add(newColumn);
            newColumn.SetOrdinal(insertPosition);
            handleDatatable.CreateOutputFile(dataTable, txtFolder.Text + "\\output");
            handleDatatable.CopyDataTableToClipboard(dataTable);
            MessageBox.Show("Handled! Copied to clipboard!");
        }

        private async void btnWriteData_Click(object sender, EventArgs e)
        {
            string a = string.Empty;
            try
            {
                string client_id = "1040164464124-mhmrgtl70icqhubokdklj291n3n2p3bg.apps.googleusercontent.com";
                string client_secret = "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";

                string[] scopes = new[] { Google.Apis.Sheets.v4.SheetsService.Scope.Spreadsheets };
                UserCredential credential = GoogleTryConnect.Login(client_id, client_secret, scopes);
                GoogleSheetsManager manager = new GoogleSheetsManager(credential);
                //string spreadsheetId = "17ghQ20O-2Id6HUmLsCznftXxdPb1we1VyHv5Jfeniso"; //1y4z8xJtt0_bvfG3f_g-dlvdfApAHTxsGMCM8eCDZ-L4
                //string spreadsheetId = "1y4z8xJtt0_bvfG3f_g-dlvdfApAHTxsGMCM8eCDZ-L4"; 
                string spreadsheetId = "1Pi31w37XDOWb1R6C9kAEPWxv91SP5zaK";
                string sheetName = "0"; // Tên sheet cần ghi

                // ID của Google Sheet (lấy từ URL: https://docs.google.com/spreadsheets/d/{spreadsheetId}/edit)

                // Gọi phương thức ghi dữ liệu
                //var newSheet = manager.CreateNew("Test document titled");
                var sheet = manager.WriteData(spreadsheetId, sheetName, dataTable);

                txtErr.Text += Environment.NewLine + ("SpreadsheetId: " + sheet.SpreadsheetId);
                txtErr.Text += Environment.NewLine + ("SpreadsheetUrl: " + sheet.SpreadsheetUrl);
                // MessageBox.Show("Dữ liệu đã được ghi thành công!");
            }
            catch (Exception ex)
            {
                txtErr.Text = a;
                MessageBox.Show($"Lỗi: {ex.Message}");
            }

        }

        private void btnCreateSheet_Click(object sender, EventArgs e)
        {
            string client_id = "1040164464124-mhmrgtl70icqhubokdklj291n3n2p3bg.apps.googleusercontent.com";
            string client_secret = "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";

            string[] scopes = new[] { Google.Apis.Sheets.v4.SheetsService.Scope.Spreadsheets };

            UserCredential credential = GoogleTryConnect.Login(client_id, client_secret, scopes);
            GoogleSheetsManager manager = new GoogleSheetsManager(credential);

            var newSheet = manager.CreateNew("Test document titled");
            txtErr.Text += Environment.NewLine + ("SpreadsheetId: " + newSheet.SpreadsheetId);
            txtErr.Text += Environment.NewLine + ("SpreadsheetUrl: " + newSheet.SpreadsheetUrl);

        }
    }
}
