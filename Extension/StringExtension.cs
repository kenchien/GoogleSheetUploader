using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Common.Extension {
   public static class StringExtension {
      /// <summary>
      /// 把淺層class的所有屬性跟值,轉換成字串,方便輸出debug
      /// </summary>
      /// <typeparam name="T"></typeparam>
      /// <param name="classObj"></param>
      /// <param name="noLimit">substring</param>
      /// <returns></returns>
      public static string ToString3<T>(this T classObj, bool noLimit = false) {
         string res = string.Empty;
         foreach (var prop in classObj.GetType().GetProperties()) {
            if (prop.GetValue(classObj, null) != null)
               res += string.Format("{0}{1}={2}",
                                   string.IsNullOrEmpty(res) ? "" : ", ",
                                   prop.Name,
                                   prop.GetValue(classObj, null));
         }

         //ken,通常呼叫這個是需要把整個物件寫到紀錄table中,該欄位目前開varchar(3000)
         //欄位超過長度好像會自動截斷,這邊先不做縮限字元的處理

         return string.IsNullOrEmpty(res) ? "{}" : "{ " + res + " }";
      }

      public static string ToString2<T>(this List<T> l) {
         string retVal = string.Empty;
         foreach (T item in l)
            retVal += string.Format("{0}{1}", string.IsNullOrEmpty(retVal) ? "" : ", ", item);
         return string.IsNullOrEmpty(retVal) ? "{}" : "{ " + retVal + " }";
      }

      public static string ToString<T>(this List<T> l, string fmt) {
         string retVal = string.Empty;
         foreach (T item in l) {
            IFormattable ifmt = item as IFormattable;
            if (ifmt != null)
               retVal += string.Format("{0}{1}",
                                       string.IsNullOrEmpty(retVal) ?
                                          "" : ", ", ifmt.ToString(fmt, null));
            else
               retVal += ToString2(l);
         }
         return string.IsNullOrEmpty(retVal) ? "{}" : "{ " + retVal + " }";
      }


      public static bool IdnoCheck(this string v) {
         if (string.IsNullOrEmpty(v))
            return false; //id為空值或是null，回傳 ID 錯誤
         if (v.Length != 10)
            return false;

         v = v.ToUpper();
         var regex = new Regex("^[A-Z]{1}[0-9]{9}$");
         if (!regex.IsMatch(v))
            return false; //Regular Expression 驗證失敗，回傳 ID 錯誤

         int[] seed = new int[9]; //除了檢查碼外每個數字的存放空間
         string[] charMapping = new string[]
         {
                "A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "L", "M", "N", "P", "Q", "R", "S", "T", "U", "V", "X",
                "Y", "W", "Z", "I", "O"
         };
         //A=10 B=11 C=12 D=13 E=14 F=15 G=16 H=17 J=18 K=19 L=20 M=21 N=22
         //P=23 Q=24 R=25 S=26 T=27 U=28 V=29 X=30 Y=31 W=32  Z=33 I=34 O=35
         string target = v.Substring(0, 1);
         for (int index = 0; index < charMapping.Length; index++) {
            if (charMapping[index] == target) {
               index += 10;
               int n1 = index / 10; //十位數
               int n2 = index % 10; //個位數

               seed[0] = (n2 * 9 + n1) % 10; //
               break;
            }
         }

         for (int index = 1; index < 9; index++) {
            //將剩餘數字乘上權數後放入存放空間
            seed[index] = Convert.ToInt32(v.Substring(index, 1)) * (9 - index);
         }

         //檢查是否符合檢查規則，10減存放空間所有數字和除以10的餘數的個位數字是否等於
         //檢查碼
         //(10 - ((seed[0] + .... + seed[9]) % 10)) % 10 == 身分證字號的最後一碼
         return ((10 - (seed.Sum() % 10)) % 10) == Convert.ToInt32(v.Substring(9, 1));
      }

      /// <summary>
      /// 強制多少個字元就換行
      /// </summary>
      /// <param name="text"></param>
      /// <param name="lineLength"></param>
      /// <returns></returns>
      public static string SpliceText(this string text, int lineLength = 34) {
         return Regex.Replace(text, "(.{" + lineLength + "})", "$1" + Environment.NewLine);
      }

      public static string ToSafeString(this string? value) {
         return value ?? string.Empty;
      }

      public static bool IsNullOrEmpty(this string? value) {
         return string.IsNullOrEmpty(value);
      }

      public static bool IsNullOrWhiteSpace(this string? value) {
         return string.IsNullOrWhiteSpace(value);
      }

      public static string TrimSafe(this string? value) {
         return value?.Trim() ?? string.Empty;
      }

      public static string ToUpperSafe(this string? value) {
         return value?.ToUpper() ?? string.Empty;
      }

      public static string ToLowerSafe(this string? value) {
         return value?.ToLower() ?? string.Empty;
      }

      public static bool IsMatch(this string? value, string pattern) {
         if (value == null) return false;
         return Regex.IsMatch(value, pattern);
      }

      public static string ReplaceSafe(this string? value, string oldValue, string newValue) {
         if (value == null) return string.Empty;
         return value.Replace(oldValue, newValue);
      }

      public static string SubstringSafe(this string? value, int startIndex, int length) {
         if (value == null) return string.Empty;
         if (startIndex < 0) startIndex = 0;
         if (startIndex >= value.Length) return string.Empty;
         if (length < 0) length = 0;
         if (startIndex + length > value.Length) length = value.Length - startIndex;
         return value.Substring(startIndex, length);
      }

      public static string FormatWith(this string? format, params object?[] args) {
         if (format == null) return string.Empty;
         return string.Format(format, args);
      }

      public static string ToStringSafe(this object? value) {
         return value?.ToString() ?? string.Empty;
      }
   }
}