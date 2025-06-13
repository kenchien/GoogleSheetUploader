# 桃教網路報修報表上傳工具

這是一個自動化工具，用於下載網路報修資料並上傳至 Google Sheets。

## 功能特點

- 自動下載指定年份範圍的報修資料
- 支援資料過濾和處理
- 自動上傳處理後的資料到 Google Sheets
- 支援多環境配置（Development、Production）
- 自動發送執行結果通知郵件
- 詳細的日誌記錄

## 系統需求

- .NET Core 3.1 或更高版本
- Google Cloud Platform 專案和服務帳號憑證
- Office 365 郵件帳號（用於發送通知）

## 安裝步驟

1. 複製專案到本地：
```bash
git clone [repository-url]
```

2. 安裝必要的 NuGet 套件：
```bash
dotnet restore
```

3. 設定環境變數：
   - 設定 `ASPNETCORE_ENVIRONMENT`（預設為 "Production"）

4. 配置設定檔：
   - 在專案根目錄建立 `appsettings.Production.json` 或 `appsettings.Development.json`
   - 填入必要的設定值（見下方配置說明）

## 配置說明

在 appsettings.json 中需要設定以下參數：

### Google Sheets 設定
```json
{
  "GoogleSheets": {
    "SpreadsheetId": "您的試算表 ID",
    "SheetName": "工作表名稱",
    "ClientSecretPath": "服務帳號憑證檔案路徑",
    "UserEmail": "使用者電子郵件"
  }
}
```

### 下載設定
```json
{
  "Download": {
    "BaseUrl": "基礎 URL",
    "UrlFormat": "URL 格式字串",
    "YearsToDownload": "要下載的年數",
    "Filter": {
      "CustomerName": "客戶名稱",
      "ExcludeServiceType": "要排除的服務類型"
    }
  }
}
```

### 郵件設定
```json
{
  "Mail": {
    "SmtpServer": "smtp.gmail.com",
    "SmtpPort": 587,
    "NeedSSL": true,
    "SmtpCredential": true,
    "AuthAccount1": "您的 Gmail 帳號",
    "AuthAccount2": "您的應用程式密碼",
    "MailSender": "寄件者郵件地址",
    "MailSenderDisplay": "桃教_網路報修報表",
    "ToRecipients": "收件者郵件地址"
  }
}
```

## 使用方法

1. 確保所有配置都已完成
2. 執行程式：
```bash
dotnet run
```

程式會：
1. 下載指定年份範圍的報修資料
2. 處理和過濾資料
3. 將處理後的資料上傳到指定的 Google Sheet
4. 發送執行結果通知郵件

## 郵件通知

程式會在以下情況發送郵件通知：

1. 執行成功時：
   - 主旨：桃教_網路報修報表(環境)排程成功
   - 內容包含處理時間和上傳筆數

2. 執行失敗時：
   - 主旨：桃教_網路報修報表(環境)排程失敗
   - 內容包含錯誤時間、錯誤詳情和堆疊追蹤

## 日誌

程式執行過程中的日誌會記錄在應用程式目錄下的日誌檔案中。

## 注意事項

- 確保有足夠的磁碟空間用於暫存檔案
- 確保 Google Sheets API 已啟用
- 確保服務帳號有適當的權限
- 如果使用 Gmail 發送郵件，需要：
  - 在 Gmail 帳號中啟用「低安全性應用程式存取」或
  - 使用應用程式密碼（如果啟用了兩步驟驗證）
  - 應用程式密碼可以在 Google 帳號的「安全性」設定中生成

## 授權

[授權說明] 