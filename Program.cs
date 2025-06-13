using System;
using System.Net.Http;
using System.Threading.Tasks;
using Common.Helper;
using System.IO;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.EnvironmentVariables;
using System.Data;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using GoogleSheetUploader.Helper;

namespace GoogleSheetUploader
{
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                LogHelper.Info("程式開始執行");

                // 讀取設定檔
                LogHelper.Info("開始讀取設定檔");
                var environment = Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT") ?? "Development";
                LogHelper.Info($"目前環境: {environment}");

                var configuration = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .AddJsonFile($"appsettings.{environment}.json", optional: true, reloadOnChange: true)
                    .AddEnvironmentVariables()
                    .Build();
                LogHelper.Info("設定檔讀取完成");

                // 從設定檔讀取參數
                var spreadsheetId = configuration["GoogleSheets:SpreadsheetId"] ?? throw new InvalidOperationException("SpreadsheetId 未設定");
                var sheetName = configuration["GoogleSheets:SheetName"] ?? throw new InvalidOperationException("SheetName 未設定");
                var clientSecretPath = configuration["GoogleSheets:ClientSecretPath"] ?? throw new InvalidOperationException("ClientSecretPath 未設定");
                var userEmail = configuration["GoogleSheets:UserEmail"] ?? throw new InvalidOperationException("UserEmail 未設定");

                var baseUrl = configuration["Download:BaseUrl"] ?? throw new InvalidOperationException("BaseUrl 未設定");
                var urlFormat = configuration["Download:UrlFormat"] ?? throw new InvalidOperationException("UrlFormat 未設定");
                var yearsToDownload = int.Parse(configuration["Download:YearsToDownload"] ?? throw new InvalidOperationException("YearsToDownload 未設定"));
                var customerName = configuration["Download:Filter:CustomerName"] ?? throw new InvalidOperationException("CustomerName 未設定");
                var excludeServiceType = configuration["Download:Filter:ExcludeServiceType"] ?? throw new InvalidOperationException("ExcludeServiceType 未設定");

                LogHelper.Info($"開始處理 {yearsToDownload} 年的資料");

                var allData = new List<IList<Object>>();
                bool isFirstFile = true;

                // 下載並處理每一年的資料
                for (int year = 0; year < yearsToDownload; year++)
                {
                    DateTime checkDay = DateTime.Now.AddDays(-year * 365 - 1);
                    string beginDay = $"{checkDay:yyyy}-01-01";
                    string endDay = year == 0 ? 
                        checkDay.ToString("yyyy-MM-dd") : // 當年使用當前日期
                        $"{checkDay:yyyy}-12-31";         // 往年使用年底
                    
                    string url = baseUrl + string.Format(urlFormat, beginDay, endDay);
                    string fileName = $"{checkDay:yyyy}-call.xls";
                    string filePath = Path.Combine(Directory.GetCurrentDirectory(), fileName);

                    LogHelper.Info($"開始處理 {fileName}");

                    // 下載檔案
                    string tempFilePath = await DownloadXlsFile(url, filePath);
                    //LogHelper.Info($"檔案已下載到: {tempFilePath}");

                    // 使用 NPOI 讀取 XLS 檔案
                    //LogHelper.Info("開始讀取 XLS 檔案資料");

                    using (var stream = new FileStream(tempFilePath, FileMode.Open, FileAccess.Read))
                    {
                        IWorkbook workbook = new HSSFWorkbook(stream);
                        ISheet worksheet = workbook.GetSheetAt(0);
                        LogHelper.Info($"工作表名稱: {worksheet.SheetName}, 總列數: {worksheet.LastRowNum + 1}");

                        // 取得欄位索引
                        var headers = new Dictionary<string, int>();
                        var headerRow = worksheet.GetRow(0);
                        if (headerRow == null) throw new InvalidOperationException("找不到標題列");

                        //LogHelper.Info("標題列內容:");
                        for (int col = 0; col < headerRow.LastCellNum; col++)
                        {
                            var cell = headerRow.GetCell(col);
                            if (cell != null)
                            {
                                var cellValue = cell.StringCellValue;
                                headers[cellValue] = col;
                                //LogHelper.Info($"欄位 {col}: {cellValue}");
                            }
                        }

                        // 過濾資料
                        var filteredRows = new List<DataRow>();
                        for (int row = 1; row <= worksheet.LastRowNum; row++)
                        {
                            var dataRow = worksheet.GetRow(row);
                            if (dataRow == null) continue;

                            var rowData = new DataRow();
                            rowData.Cell = new Dictionary<string, string>();

                            foreach (var header in headers)
                            {
                                var cell = dataRow.GetCell(header.Value);
                                rowData.Cell[header.Key] = cell?.ToString() ?? string.Empty;
                            }

                            // 檢查過濾條件
                            bool matchesFilter = rowData.Cell["客戶名稱"] == customerName &&
                                !string.IsNullOrEmpty(rowData.Cell["客戶單位"]) &&
                                !string.IsNullOrEmpty(rowData.Cell["合約編號"]) &&
                                rowData.Cell["服務類型"] != excludeServiceType;

                            

                            if (matchesFilter)
                            {
                                filteredRows.Add(rowData);
                                //LogHelper.Info($"符合條件的資料行 {row}: 客戶名稱={rowData.Cell["客戶名稱"]}, 合約編號={rowData.Cell["合約編號"]}, 服務類型={rowData.Cell["服務類型"]}");
                            }
                        }

                        LogHelper.Info($"讀取到 {filteredRows.Count} 筆符合條件的資料");

                        // 如果是第一個檔案，添加標題行
                        if (isFirstFile)
                        {
                            var titleRow = new List<object>
                            {
                                "合約編號", "客戶單位", "到場時間", "服務類型", "叫修說明", "處理說明",
                                "是否更換備品", "更換備品", "工程師完成狀態", "到場工程師", "處理時間(時)",
                                "廠牌1", "型號1", "序號1", "廠牌2", "型號2", "序號2",
                                "廠牌3", "型號3", "序號3"
                            };
                            allData.Add(titleRow);
                            isFirstFile = false;
                        }

                        // 處理資料行
                        string tempOutNumber = "";
                        int processedRows = 0;
                        foreach (var row in filteredRows)
                        {
                            if (row.Cell["出勤編號"] == tempOutNumber)
                                continue;

                            tempOutNumber = row.Cell["出勤編號"];
                            var dataRow = new List<object>
                            {
                                row.Cell["合約編號"], row.Cell["客戶單位"], row.Cell["到場時間"], row.Cell["服務類型"],
                                row.Cell["叫修說明"], row.Cell["處理說明"], row.Cell["是否更換備品"], row.Cell["更換備品"],
                                row.Cell["工程師完成狀態"], row.Cell["到場工程師"], row.Cell["處理時間(時)"],
                                row.Cell["廠牌1"], row.Cell["型號1"], row.Cell["序號1"], row.Cell["廠牌2"],
                                row.Cell["型號2"], row.Cell["序號2"], row.Cell["廠牌3"], row.Cell["型號3"], row.Cell["序號3"]
                            };
                            allData.Add(dataRow);
                            processedRows++;
                        }

                        LogHelper.Info($"處理完成 {processedRows} 筆資料");
                    }

                    // 清理暫存檔案
                    //File.Delete(tempFilePath);
                    //LogHelper.Info($"已刪除暫存檔案: {tempFilePath}");
                }

                // 上傳到 Google Sheets
                LogHelper.Info("開始上傳資料到 Google Sheets");
                var googleSheetHelper = new GoogleSheetHelper(clientSecretPath, userEmail, spreadsheetId, sheetName);
                
                // 清空 Google Sheet
                LogHelper.Info("正在清空 Google Sheet...");
                await googleSheetHelper.ClearSheetAsync(sheetName);
                
                // 上傳資料
                LogHelper.Info("正在上傳資料...");
                await googleSheetHelper.WriteSheetAsync(sheetName, allData);

                LogHelper.Info("所有處理完成！");
            }
            catch (Exception ex)
            {
                LogHelper.Error("發生錯誤", ex);
            }
        }

        static async Task<string> DownloadXlsFile(string url, string filePath)
        {
            LogHelper.Info($"開始下載檔案: {url}");
            
            // 檢查檔案是否存在
            if (File.Exists(filePath))
            {
                LogHelper.Info($"檔案已存在，將進行覆蓋: {filePath}");
                try
                {
                    File.Delete(filePath);
                    LogHelper.Info("已刪除舊檔案");
                }
                catch (Exception ex)
                {
                    LogHelper.Error($"刪除舊檔案時發生錯誤: {ex.Message}");
                    throw;
                }
            }

            using (var client = new HttpClient())
            {
                try
                {
                    var response = await client.GetAsync(url);
                    response.EnsureSuccessStatusCode();

                    using (var fileStream = File.Create(filePath))
                    {
                        await response.Content.CopyToAsync(fileStream);
                    }
                    LogHelper.Info($"檔案下載完成: {filePath}");
                    return filePath;
                }
                catch (Exception ex)
                {
                    LogHelper.Error($"下載檔案時發生錯誤: {ex.Message}");
                    throw;
                }
            }
        }
    }

    public class DataRow
    {
        public Dictionary<string, string> Cell { get; set; } = new Dictionary<string, string>();
    }
}
