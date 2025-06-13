using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Common.Extension {
   public static class DateTimeExtecsion {
      public static string ToTaiwanDateTime(this DateTime dt, string format) {

         CultureInfo culture = new CultureInfo("zh-TW");
         culture.DateTimeFormat.Calendar = new TaiwanCalendar();

         string result = dt.ToString(format, culture);

         //ken,系統問題,format寫yyMM或是yMM 基本就只會回yyyMM,自己過濾調整吧(先用很笨拙的方式吧)
         string finalResult = result;
         if (format.Substring(0, 2) == "yy" && format.Substring(2, 1) != "y")
            finalResult = result.Substring(1);

         return finalResult;

      }

      public static string ToTaiwanDateTime(this string d, string format) {
         bool re = DateTime.TryParse(d, out DateTime dt);
         if (re)
            return dt.ToTaiwanDateTime(format);
         else
            return null;
      }

      public static string DateTimeSplitSymbol(this string d, char s) {
         string re = string.Empty;
         List<string> words = d.Split(s).ToList();

         foreach (string w in words) {
            re += w;
         }

         return re;
      }

      public static string NoSplit(this string source) {
         if (string.IsNullOrEmpty(source)) return null;

         return source.Replace("/", string.Empty);
      }

      public static bool DateCheck(this string d) {
         if (d == null)
            return false;
         d = d.Trim();
         if (d.Length != 8)
            return false;

         int iyear = Convert.ToInt16(d.Substring(0, 4));
         int imonth = Convert.ToInt16(d.Substring(4, 2));
         int iday = Convert.ToInt16(d.Substring(6, 2));
         if (iyear < 1900) { return false; }
         if (imonth > 12 || imonth < 1) { return false; }
         if (iday > DateTime.DaysInMonth(iyear, imonth) || iday < 1) { return false; };
         return true;
      }
   }
}