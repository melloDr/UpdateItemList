using System.Data;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UpdateItemList.Utils;

namespace UpdateItemList
{
    public partial class Form1 : Form
    {
        HandleDatatable handleDatatable = new HandleDatatable();
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

            DataTable dataTable = new DataTable();
            for (int i = 0; i < excelFiles.Length; i++)
            {
                dataTable = ReadExcelData(excelFiles[i], i, ref dataTable);
            }
            handleDatatable.CreateOutputFile(dataTable, txtFolder.Text + "\\output");

        }
    }
}
