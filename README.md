# Google Sheet 資料上傳工具

這是一個用於下載 XLS 檔案並將資料上傳至 Google Sheets 的 .NET 工具。

## 功能特點

- 從指定 URL 下載 XLS 檔案
- 支援多年度資料下載
- 資料過濾功能
- 自動處理 Google OAuth2 認證
- 支援資料上傳至 Google Sheets

## 必要條件

- .NET 6.0 或更新版本
- Google Cloud Platform 專案
- OAuth 2.0 憑證檔案

## 設定說明

1. 在 Google Cloud Console 建立專案並啟用 Google Sheets API
2. 下載 OAuth 2.0 憑證檔案（client_secret.json）
3. 將憑證檔案放在專案根目錄
4. 修改 `appsettings.json` 設定檔：

```json
{
  "GoogleSheets": {
    "SpreadsheetId": "您的試算表ID",
    "SheetName": "工作表名稱",
    "ClientSecretPath": "client_secret.json",
    "UserEmail": "您的Google帳號"
  },
  "Download": {
    "BaseUrl": "下載網址基礎部分",
    "UrlFormat": "下載網址格式",
    "YearsToDownload": 1,
    "Filter": {
      "CustomerName": "客戶名稱",
      "ContractNumber": "合約編號",
      "ExcludeServiceType": "排除的服務類型"
    }
  }
}
```

## 使用方式

1. 確保已正確設定 `appsettings.json`
2. 執行程式：

```bash
dotnet run
```

## 注意事項

- 首次執行時會開啟瀏覽器要求授權
- 授權後會自動儲存 token 以供後續使用
- 請確保網路連線正常
- 請確保有足夠的磁碟空間存放暫存檔案

## 授權

MIT License 