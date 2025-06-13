using Common.Extension;
using log4net;
using System;
using System.Configuration;
using System.Net.Mail;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Common.Helper;

namespace GoogleSheetUploader.Helper {
   /// <summary>
   /// 寄信的動作和一般呼叫函數一樣，程式會等動作做完後﹝確定送到mail server﹞才繼續往下跑，
   /// 如果一次要寄出多封的郵件，將變得很沒效率﹝程式會卡住﹞，最好能採用非同步的方式來執行程式
   /// (PS:不過用非同步要注意收攏結果,不是把信送出就當作成功)先把全部改為同步送信
   /// </summary>
   public class MailHelper {
      private static readonly ILog log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

      public delegate bool SendMailDelegate(string toRecipients,
                                              string mailSubject,
                                              string content,
                                              bool useBCC = false,
                                              Attachment attachment = null);

      private readonly string _smtpServer;
      private readonly string _mailSender;
      private readonly int _smtpPort;
      private readonly bool _smtpCredential;
      private readonly bool _needSSL;
      private readonly string _authAccount1;
      private readonly string _authAccount2;
      private readonly string _mailSenderDisplay;
      private readonly string _toRecipients;

      public MailHelper(IConfiguration configuration) {
         _smtpServer = configuration["Mail:SmtpServer"] ?? throw new InvalidOperationException("SmtpServer 未設定");
         _smtpPort = int.Parse(configuration["Mail:SmtpPort"] ?? throw new InvalidOperationException("SmtpPort 未設定"));
         _needSSL = bool.Parse(configuration["Mail:NeedSSL"] ?? throw new InvalidOperationException("NeedSSL 未設定"));
         _smtpCredential = bool.Parse(configuration["Mail:SmtpCredential"] ?? throw new InvalidOperationException("SmtpCredential 未設定"));
         _authAccount1 = configuration["Mail:AuthAccount1"] ?? throw new InvalidOperationException("AuthAccount1 未設定");
         _authAccount2 = configuration["Mail:AuthAccount2"] ?? throw new InvalidOperationException("AuthAccount2 未設定");
         _mailSender = configuration["Mail:MailSender"] ?? throw new InvalidOperationException("MailSender 未設定");
         _mailSenderDisplay = configuration["Mail:MailSenderDisplay"] ?? throw new InvalidOperationException("MailSenderDisplay 未設定");
         _toRecipients = configuration["Mail:ToRecipients"] ?? throw new InvalidOperationException("ToRecipients 未設定");
      }

      /// <summary>
      /// 發信簡單版 (use SmtpClient)
      /// </summary>
      /// <param name="content"></param>
      /// <param name="mailSubject"></param>
      /// <returns></returns>
      public bool SendMail(List<string> content, string mailSubject="桃教_報修系統_排程異常") {
         try {
            bool useBCC = true;
            Attachment attachment = null;
            

            log.InfoEx($@"ready send mail to {_toRecipients}, mailSubject={mailSubject}");
            log.InfoEx($"......smtpServer={_smtpServer},smtpPort={_smtpPort},mailSender={_mailSender},mailSenderDisplay={_mailSenderDisplay}");
            log.InfoEx($"......needSSL={_needSSL},smtpCredential={_smtpCredential},authAccount1={_authAccount1},authAccount2=no display in log");

            using (MailMessage msg = new MailMessage()) {
               msg.From = new MailAddress(_mailSender, _mailSenderDisplay, System.Text.Encoding.UTF8);
               if (useBCC)
                  msg.Bcc.Add(_toRecipients);
               else
                  msg.To.Add(_toRecipients);

               msg.Subject = mailSubject;
               msg.SubjectEncoding = System.Text.Encoding.UTF8;
               msg.Body = String.Join("\r\n", content.ToArray());
               msg.IsBodyHtml = false;
               msg.BodyEncoding = System.Text.Encoding.UTF8;
               msg.Priority = MailPriority.Normal;

               if (attachment != null)
                  msg.Attachments.Add(attachment);

               using (SmtpClient client = new SmtpClient(_smtpServer, _smtpPort)) {
                  client.DeliveryMethod = SmtpDeliveryMethod.Network;
                  if (_smtpCredential) {
                     client.UseDefaultCredentials = false;
                     client.Credentials = new System.Net.NetworkCredential(_authAccount1, _authAccount2);
                  }

                  if (_needSSL)
                     client.EnableSsl = true;
                  client.Send(msg);
               }//using (SmtpClient client = new SmtpClient(smtpServer, smtpPort)) {

               msg.Attachments.Dispose();
               msg.Dispose();
            }//using (MailMessage msg = new MailMessage()) {

            log.InfoEx("done, successful send mail.");
            return true;
         } catch (Exception ex) {
            log.ErrorFormatEx(ex, "SendMail Error,smtpServer={0},smtpPort={1},mailSender={2},toRecipients={3}", _smtpServer, _smtpPort, _mailSender, _toRecipients);
            return false;
         }
      }

      /// <summary>
      /// 發送郵件
      /// </summary>
      public bool SendMail(string toRecipients, string mailSubject, string content, bool useBCC = false, Attachment? attachment = null) {
         try {
            LogHelper.Info($"準備發送郵件給 {toRecipients}, 主旨={mailSubject}");
            LogHelper.Info($"......smtpServer={_smtpServer},smtpPort={_smtpPort},mailSender={_mailSender},mailSenderDisplay={_mailSenderDisplay}");
            LogHelper.Info($"......needSSL={_needSSL},smtpCredential={_smtpCredential},authAccount1={_authAccount1}");

            using var msg = new MailMessage();
            msg.From = new MailAddress(_mailSender, _mailSenderDisplay, System.Text.Encoding.UTF8);
            if (useBCC)
               msg.Bcc.Add(toRecipients);
            else
               msg.To.Add(toRecipients);

            msg.Subject = mailSubject;
            msg.SubjectEncoding = System.Text.Encoding.UTF8;
            msg.Body = content;
            msg.IsBodyHtml = false;
            msg.BodyEncoding = System.Text.Encoding.UTF8;
            msg.Priority = MailPriority.Normal;

            if (attachment != null)
               msg.Attachments.Add(attachment);

            using var client = new SmtpClient(_smtpServer, _smtpPort);
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            if (_smtpCredential) {
               client.UseDefaultCredentials = false;
               client.Credentials = new System.Net.NetworkCredential(_authAccount1, _authAccount2);
            }

            if (_needSSL)
               client.EnableSsl = true;
            client.Send(msg);

            LogHelper.Info("郵件發送成功");
            return true;
         } catch (Exception ex) {
            LogHelper.Error("發送郵件時發生錯誤", ex);
            return false;
         }
      }

      /// <summary>
      /// 非同步寄信 (use SmtpClient)
      /// </summary>
      /// <param name="toRecipients"></param>
      /// <param name="mailSubject"></param>
      /// <param name="content"></param>
      /// <param name="useBCC"></param>
      /// <param name="attachment"></param>
      /// <returns></returns>
      public bool SendMailAsync(string toRecipients,
                                  string mailSubject,
                                  string content,
                                  bool useBCC = false,
                                  Attachment attachment = null) {
         try {
            SendMailDelegate dc = new SendMailDelegate(this.SendMail);
            IAsyncResult result = dc.BeginInvoke(toRecipients, mailSubject, content, useBCC, attachment, null, null);

            #region //可以用EndInvoke等待非同步的結果
            //// Poll while simulating work.
            //while (result.IsCompleted == false) {
            //    Thread.Sleep(250);
            //    Console.Write(".");
            //}

            //var res = dc.EndInvoke(result);
            #endregion
            return true;
         } catch (Exception ex) {
            log.ErrorFormatEx(ex, "SendMailAsync Error", null);
            return false;
         }
      }
   }
}
