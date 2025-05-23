using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace UpdateItemList.Utils
{
    internal class HandleDatatable
    {
        internal string GetCellValue(ICell cell, IFormulaEvaluator evaluator)
        {
            if (cell == null) return "";

            switch (cell.CellType)
            {
                case CellType.Formula:
                    CellValue evaluatedValue = evaluator.Evaluate(cell);
                    return FormatCellValue(cell, evaluatedValue);
                default:
                    return FormatCellValue(cell, null);
            }
        }
        private string FormatCellValue(ICell cell, CellValue? evaluatedValue)
        {
            if (evaluatedValue != null)
            {
                switch (evaluatedValue.CellType)
                {
                    case CellType.Numeric:
                        string cellValue = cell?.ToString() ?? "0";
                        double.TryParse(cellValue, out double oaDate);

                        return DateUtil.IsCellDateFormatted(cell)
                            ? DateTime.FromOADate(oaDate).ToString("dd/MM/yyyy")
                            : cell.NumericCellValue.ToString();
                    case CellType.String:
                        return cell.StringCellValue;
                    case CellType.Boolean:
                        return cell.BooleanCellValue.ToString();
                    default:
                        return "";
                }
            }

            switch (cell.CellType)
            {
                case CellType.Numeric:
                    string cellValue = cell?.ToString() ?? "0";
                    double.TryParse(cellValue, out double oaDate);

                    return DateUtil.IsCellDateFormatted(cell)
                        ? DateTime.FromOADate(oaDate).ToString("dd/MM/yyyy")
                        : cell.NumericCellValue.ToString();
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                default:
                    return "";
            }
        }
        internal void CreateOutputFile(DataTable dataTable, string rootPath)
        {
            string fullPath = Path.Combine(rootPath, "MyDataExport.xlsx");
            CreateDirectoryIfNotExists(rootPath);
            ExportDataTableToExcel(dataTable, fullPath);

        }
        private static void CreateDirectoryIfNotExists(string path)
        {
            if (!Directory.Exists(path))
            {
                try
                {
                    Directory.CreateDirectory(path);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Lỗi khi tạo thư mục '{path}': {ex.Message}");
                    throw;
                }
            }
        }
        private static void ExportDataTableToExcel(DataTable dataTable, string filePath)
        {
            IWorkbook workbook;
            string fileExtension = Path.GetExtension(filePath).ToLower();

            if (fileExtension == ".xlsx")
            {
                workbook = new XSSFWorkbook(); // Dành cho .xlsx
            }
            else if (fileExtension == ".xls")
            {
                workbook = new HSSFWorkbook(); // Dành cho .xls
            }
            else
            {
                throw new ArgumentException("Định dạng file không được hỗ trợ. Vui lòng sử dụng .xls hoặc .xlsx.");
            }

            ISheet sheet = workbook.CreateSheet("Sheet1"); // Tạo một sheet mới với tên "Sheet1"

            try
            {
                // Tạo hàng đầu tiên (header)
                IRow headerRow = sheet.CreateRow(0);
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    ICell cell = headerRow.CreateCell(i);
                    cell.SetCellValue(dataTable.Columns[i].ColumnName);
                }

                // Đổ dữ liệu từ DataTable vào các hàng tiếp theo
                for (int r = 0; r < dataTable.Rows.Count; r++)
                {
                    IRow dataRow = sheet.CreateRow(r + 1); // Bắt đầu từ hàng thứ 2 (chỉ số 1)
                    for (int c = 0; c < dataTable.Columns.Count; c++)
                    {
                        ICell cell = dataRow.CreateCell(c);
                        object cellValue = dataTable.Rows[r][c];

                        // Xử lý kiểu dữ liệu để ghi vào Excel đúng cách
                        if (cellValue != DBNull.Value && cellValue != null)
                        {
                            if (cellValue is int || cellValue is long || cellValue is double || cellValue is float || cellValue is decimal)
                            {
                                cell.SetCellValue(Convert.ToDouble(cellValue));
                            }
                            else if (cellValue is DateTime)
                            {
                                cell.SetCellValue((DateTime)cellValue);
                                // Tùy chọn: thiết lập định dạng ngày tháng
                                ICellStyle dateStyle = workbook.CreateCellStyle();
                                dateStyle.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy-mm-dd");
                                cell.CellStyle = dateStyle;
                            }
                            else if (cellValue is bool)
                            {
                                cell.SetCellValue((bool)cellValue);
                            }
                            else
                            {
                                cell.SetCellValue(cellValue.ToString());
                            }
                        }
                        else
                        {
                            cell.SetCellValue(""); // Hoặc để trống nếu giá trị là DBNull
                        }
                    }
                }

                // Tùy chọn: Tự động điều chỉnh độ rộng cột
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    sheet.AutoSizeColumn(i);
                }

                // Lưu workbook vào file
                using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi xuất Excel: {ex.Message}");
                throw; // Re-throw để xử lý lỗi ở nơi gọi
            }
            finally
            {
                workbook.Close(); // Đảm bảo workbook được đóng
            }
        }


    }
}
