using System;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Common.Helper;

namespace GoogleSheetUploader.Helper {
    /// <summary>
    /// 系統異常通知郵件發送輔助類別
    /// </summary>
    public class SystemAlertHelper {
        private readonly MailHelper _mailHelper;
        private readonly string _toRecipients;
        private readonly bool _alwaysSendMail;

        public SystemAlertHelper(IConfiguration configuration, bool alwaysSendMail = false) {
            _mailHelper = new MailHelper(configuration);
            _toRecipients = configuration["Mail:ToRecipients"] ?? throw new InvalidOperationException("ToRecipients 未設定");
            _alwaysSendMail = alwaysSendMail;
        }

        /// <summary>
        /// 發送系統異常通知郵件
        /// </summary>
        public async Task SendAlertMailAsync(string subject, string content) {
            try {
                LogHelper.Info($"準備發送系統異常通知郵件，主旨：{subject}");
                var result = await Task.Run(() => _mailHelper.SendMail(_toRecipients, subject, content));
                if (result) {
                    LogHelper.Info("系統異常通知郵件發送成功");
                } else {
                    LogHelper.Error("系統異常通知郵件發送失敗");
                }
            } catch (Exception ex) {
                LogHelper.Error("發送系統異常通知郵件時發生錯誤", ex);
            }
        }

        /// <summary>
        /// 發送資料數量異常通知郵件
        /// </summary>
        public async Task SendDataCountAlertAsync(int count) {
            if (!_alwaysSendMail && count >= 100) {
                return;
            }

            try {
                var subject = "Google Sheet 資料數量異常通知";
                var content = $"目前上傳的 Google Sheet 資料數量為 {count}，低於預期的 100 筆。";
                
                LogHelper.Info($"準備發送資料數量異常通知郵件，數量：{count}");
                var result = await Task.Run(() => _mailHelper.SendMail(_toRecipients, subject, content));
                if (result) {
                    LogHelper.Info("資料數量異常通知郵件發送成功");
                } else {
                    LogHelper.Error("資料數量異常通知郵件發送失敗");
                }
            } catch (Exception ex) {
                LogHelper.Error("發送資料數量異常通知郵件時發生錯誤", ex);
            }
        }

        /// <summary>
        /// 發送系統異常通知郵件（不等待結果）
        /// </summary>
        public void SendAlertMail(string subject, string content) {
            _ = SendAlertMailAsync(subject, content);
        }

        /// <summary>
        /// 發送資料數量異常通知郵件（不等待結果）
        /// </summary>
        public void SendDataCountAlert(int count) {
            _ = SendDataCountAlertAsync(count);
        }
    }
} 