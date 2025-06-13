using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using Common.Helper;

namespace GoogleSheetUploader.Helper {
   public class ExcelHelper : IDisposable {
      private List<Title> _titles { get; } = new List<Title>();
      private bool _disposed = false;
      private readonly string _fileName;
      private readonly bool _hasTitle;

      public ExcelPackage ExcelDoc { get; set; } = null;
      public bool PrintHeaders { get; set; } = false;
      public string ContentType { get; set; } = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

      public ExcelWorkbook WBook { get { return ExcelDoc.Workbook; } }

      public ExcelHelper(string fileName, bool hasTitle = true) {
         _fileName = fileName ?? throw new ArgumentNullException(nameof(fileName));
         _hasTitle = hasTitle;
         ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
      }

      public ExcelHelper(List<Title> titles, string path, bool printDataTitle = true) {
         ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
         var fileInfoXls = new FileInfo(path);
         ExcelDoc = new ExcelPackage(fileInfoXls);
         _titles.AddRange(titles);
         PrintHeaders = printDataTitle;
      }

      public void AddTitle(Title t) {
         _titles.Add(t);
      }

      public void SetPrintTitle(bool printDataTitle) {
         PrintHeaders = printDataTitle;
      }

      public ExcelWorksheet GetSheet(int sheetIndex) {
         if (ExcelDoc.Workbook.Worksheets.Count == 0)
            return ExcelDoc.Workbook.Worksheets.Add("sheet1");
         else if (ExcelDoc.Workbook.Worksheets.Count == 1)
            return ExcelDoc.Workbook.Worksheets[0];
         else
            return ExcelDoc.Workbook.Worksheets[sheetIndex];
      }

      public void LoadDataTable(DataTable dt, int sheetIndex = 0, string printCell = "A1", bool drawBorder = true) {
         var sheet = GetSheet(sheetIndex);

         foreach (var t in _titles) {
            sheet.Cells[t.Cell].Value = t.Text;
         }

         if (dt?.Rows?.Count == 0)
            return;

         sheet.Cells[printCell].LoadFromDataTable(dt, PrintHeaders);

         var rowCount = PrintHeaders ? dt.Rows.Count + 1 : dt.Rows.Count;
         if (drawBorder)
            WriteBorder(sheet, rowCount, dt.Columns.Count, printCell);
      }

      public void WriteBorder(ExcelWorksheet sheet, int RowCount, int ColCount, string address) {
         int fromX, fromY, toX, toY;

         var cell = new ExcelCellAddress(address);
         fromX = cell.Row;
         fromY = cell.Column;
         toX = fromX + RowCount - 1;
         toY = fromY + ColCount - 1;

         if (fromX > 0 && fromY > 0 && toX > 0 && toY > 0) {
            var modelTable = sheet.Cells[fromX, fromY, toX, toY];
            modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
         }
      }

      public IList<IList<Object>> ReadExcel(bool haveHeader = true, int sheetIndex = 0) {
         var result = new List<IList<Object>>();
         var sheet = GetSheet(sheetIndex);
         var startRow = haveHeader ? 2 : 1;
         var endRow = sheet.Dimension.End.Row;
         var endCol = sheet.Dimension.End.Column;

         for (int row = startRow; row <= endRow; row++) {
            var rowData = new List<Object>();
            for (int col = 1; col <= endCol; col++) {
               var cell = sheet.Cells[row, col];
               rowData.Add(cell.Value);
            }
            result.Add(rowData);
         }

         return result;
      }

      public class Title {
         public string Cell { get; set; }
         public string Text { get; set; }
      }

      public void Dispose() {
         Dispose(true);
         GC.SuppressFinalize(this);
      }

      protected virtual void Dispose(bool disposing) {
         if (!_disposed) {
            if (disposing) {
               ExcelDoc?.Dispose();
            }
            _disposed = true;
         }
      }

      ~ExcelHelper() {
         Dispose(false);
      }

      public static List<DataRow> ReadXlsFile(string filePath) {
         if (string.IsNullOrEmpty(filePath)) throw new ArgumentNullException(nameof(filePath));

         try {
            var dataRows = new List<DataRow>();
            using (var package = new ExcelPackage(new FileInfo(filePath))) {
               var worksheet = package.Workbook.Worksheets[0];
               if (worksheet == null) {
                  throw new InvalidOperationException("找不到工作表");
               }

               var rowCount = worksheet.Dimension?.Rows ?? 0;
               var colCount = worksheet.Dimension?.Columns ?? 0;

               if (rowCount == 0 || colCount == 0) {
                  throw new InvalidOperationException("工作表是空的");
               }

               // 讀取標題列
               var headers = new Dictionary<string, int>();
               for (int col = 1; col <= colCount; col++) {
                  var headerValue = worksheet.Cells[1, col].Value?.ToString();
                  if (!string.IsNullOrEmpty(headerValue)) {
                     headers[headerValue] = col;
                  }
               }

               // 讀取資料列
               for (int row = 2; row <= rowCount; row++) {
                  var dataRow = new DataRow {
                     Cell = new Dictionary<string, string>()
                  };

                  foreach (var header in headers) {
                     var cellValue = worksheet.Cells[row, header.Value].Value?.ToString() ?? string.Empty;
                     dataRow.Cell[header.Key] = cellValue;
                  }

                  dataRows.Add(dataRow);
               }
            }

            return dataRows;
         } catch (Exception ex) {
            LogHelper.Error("讀取 Excel 檔案時發生錯誤", ex);
            throw;
         }
      }

      public IList<IList<object>> ReadExcel() {
         try {
            LogHelper.Info($"開始讀取 Excel 檔案：{_fileName}");

            using var fs = new FileStream(_fileName, FileMode.Open, FileAccess.Read);
            var workbook = new XSSFWorkbook(fs);
            var sheet = workbook.GetSheetAt(0);
            var result = new List<IList<object>>();

            var startRow = _hasTitle ? 1 : 0;
            for (int i = startRow; i <= sheet.LastRowNum; i++) {
               var row = sheet.GetRow(i);
               if (row == null) continue;

               var rowData = new List<object>();
               for (int j = 0; j < row.LastCellNum; j++) {
                  var cell = row.GetCell(j);
                  rowData.Add(GetCellValue(cell));
               }
               result.Add(rowData);
            }

            LogHelper.Info($"成功讀取 {result.Count} 筆資料");
            return result;
         } catch (Exception ex) {
            LogHelper.Error("讀取 Excel 檔案時發生錯誤", ex);
            throw;
         }
      }

      private static object GetCellValue(ICell? cell) {
         if (cell == null) return string.Empty;

         switch (cell.CellType) {
            case CellType.Numeric:
               if (DateUtil.IsCellDateFormatted(cell)) {
                  return cell.DateCellValue;
               }
               return cell.NumericCellValue;
            case CellType.String:
               return cell.StringCellValue ?? string.Empty;
            case CellType.Boolean:
               return cell.BooleanCellValue;
            case CellType.Formula:
               switch (cell.CachedFormulaResultType) {
                  case CellType.Numeric:
                     return cell.NumericCellValue;
                  case CellType.String:
                     return cell.StringCellValue ?? string.Empty;
                  default:
                     return cell.ToString() ?? string.Empty;
               }
            default:
               return cell.ToString() ?? string.Empty;
         }
      }
   }

   public class DataRow {
      public required Dictionary<string, string> Cell { get; init; }
   }
} 