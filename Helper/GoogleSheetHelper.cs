using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Common.Helper;

namespace GoogleSheetUploader.Helper {
   public class GoogleSheetHelper {
      readonly string[] ScopesReadOnly = { SheetsService.Scope.SpreadsheetsReadonly };
      readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
      readonly string _applicationName = "Google Sheet Uploader";
      private readonly string _secretJson;
      private readonly string _user;
      private readonly string _spreadsheetId;
      private readonly string _sheetName;
      private readonly TokenHelper _tokenHelper;
      private readonly SheetsService _sheetsService;

      public GoogleSheetHelper(string secretJson = "client_secret.json",
                               string user = "andisand@gmail.com",
                               string spreadsheetId = "1Z6qN3fIQ95fEjVPWc42sq-MPXFrD4UtKzmrs0j90wgo",
                               string sheetName = "sheet1") {
         if (string.IsNullOrEmpty(secretJson)) throw new ArgumentNullException(nameof(secretJson));
         if (string.IsNullOrEmpty(user)) throw new ArgumentNullException(nameof(user));
         if (string.IsNullOrEmpty(spreadsheetId)) throw new ArgumentNullException(nameof(spreadsheetId));
         if (string.IsNullOrEmpty(sheetName)) throw new ArgumentNullException(nameof(sheetName));

         _secretJson = secretJson;
         _user = user;
         _spreadsheetId = spreadsheetId;
         _sheetName = sheetName;
         _tokenHelper = new TokenHelper(secretJson, user, Scopes);

         var credential = _tokenHelper.GetCredentialAsync().GetAwaiter().GetResult();
         _sheetsService = new SheetsService(new BaseClientService.Initializer() {
            HttpClientInitializer = credential,
            ApplicationName = _applicationName,
         });
      }

      public async Task<IList<IList<Object>>> ReadSheetAsync(string range = null) {
         IList<IList<Object>> values = null;
         try {
            string sheetRange = range == null ? _sheetName : $@"{_sheetName}!{range}";

            // 讀取資料
            var request = _sheetsService.Spreadsheets.Values.Get(_spreadsheetId, sheetRange);
            var response = await request.ExecuteAsync();
            values = response.Values;

         } catch (Exception ex) {
            Console.WriteLine($@"ReadSheet Error! msg={ex.Message}");
            throw;
         }
         return values;
      }

      public async Task WriteSheetAsync(string sheetName, IList<IList<object>> values) {
         if (string.IsNullOrEmpty(sheetName)) throw new ArgumentNullException(nameof(sheetName));
         if (values == null) throw new ArgumentNullException(nameof(values));

         try {
            var range = $"{sheetName}!A1";
            var valueRange = new ValueRange {
               Values = values
            };

            var request = _sheetsService.Spreadsheets.Values.Update(valueRange, _spreadsheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            var response = await request.ExecuteAsync();
            LogHelper.Info($"已更新 {response.UpdatedCells} 個儲存格");
         } catch (Exception ex) {
            LogHelper.Error("寫入 Google Sheet 時發生錯誤", ex);
            throw;
         }
      }

      public async Task AppendSheetAsync(IList<IList<Object>> dataList, string range = null) {
         try {
            string sheetRange = range == null ? _sheetName : $@"{_sheetName}!{range}";

            // 準備資料
            var valueRange = new ValueRange {
               Values = dataList
            };

            // 新增資料
            var request = _sheetsService.Spreadsheets.Values.Append(valueRange, _spreadsheetId, sheetRange);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var response = await request.ExecuteAsync();
            Console.WriteLine($@"append {response.Updates.UpdatedRows} rows success.");

         } catch (Exception ex) {
            Console.WriteLine($@"AppendSheet Error! msg={ex.Message}");
            throw;
         }
      }

      public async Task ClearSheetAsync(string sheetName = "Sheet1") {
         if (string.IsNullOrEmpty(sheetName)) throw new ArgumentNullException(nameof(sheetName));

         try {
            
            var range = $"{sheetName}!A1:Z1000";
            var clearRequest = new ClearValuesRequest();
            var request = _sheetsService.Spreadsheets.Values.Clear(clearRequest, _spreadsheetId, range);
            var response = await request.ExecuteAsync();
            LogHelper.Info($"已清除工作表 {sheetName}");
         } catch (Exception ex) {
            LogHelper.Error("清除 Google Sheet 時發生錯誤", ex);
            throw;
         }
      }
   }
} 