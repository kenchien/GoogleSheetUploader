using log4net;
using System;
using System.Collections.Generic;
using System.Text;

namespace Common.Extension {

   public static class Log4NetExtensions {
      public static void InfoEx(this ILog logger, string msg) {
         //ken,為了防止Log Forging攻擊,可以用下面兩種方式
         //1.過濾一些特殊字元
         logger.Info(ConvertValidString(msg));

         //2.URLEncoder.encode(userInput, "utf-8");
      }

      /// <summary>
      /// 過濾一些特殊字元 "%0d", "\r", "%0a", "\n"
      /// </summary>
      /// <param name="log"></param>
      /// <returns></returns>
      public static string ConvertValidString(String log) {
         List<string> list = new List<string> { "%0d", "\r", "%0a", "\n" };

         string encode = log.Normalize(NormalizationForm.FormKC);
         foreach (string str in list) {
            encode = encode.Replace(str, "");
         }
         return encode;
      }

      /// <summary>
      /// 除了ex,還順便抓使用者當下的畫面值
      /// </summary>
      /// <param name="logger"></param>
      /// <param name="format"></param>
      /// <param name="exception"></param>
      /// <param name="args"></param>
      public static void ErrorEx(this ILog logger, Exception exception) {
         logger.Error(exception);
      }

      /// <summary>
      /// 除了ex,還能增加自己想要的錯誤說明並使用{0}{1}...,後面能補充參數值
      /// </summary>
      /// <param name="logger"></param>
      /// <param name="format"></param>
      /// <param name="exception"></param>
      /// <param name="args"></param>
      public static void ErrorFormatEx(this ILog logger, Exception exception, string format, params object[] args) {
         logger.Error(ConvertValidString(string.Format(format, args)), exception);
      }
   }

}