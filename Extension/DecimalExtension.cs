using System;

namespace Common.Extension {
   public static class BoolExtension {
      public static int ToInt(this bool v) {
         return v ? 1 : 0;
      }

      /// <summary>
      /// 無條件進位
      /// </summary>
      /// <param name="originValue">原有的值</param>
      /// <param name="decimals">小數點第幾位</param>
      /// <returns></returns>
      public static decimal Ceiling(this decimal originValue, int decimals = 0) {

         var temp = 1;
         if (decimals == 0)
            temp = 1;
         else
            for (int k = 0; k < decimals; k++) {
               temp = temp * 10;
            }

         return Math.Ceiling((decimal)originValue * temp) / temp;

      }

      /// <summary>
      /// 無條件進位
      /// </summary>
      /// <param name="originValue">原有的值</param>
      /// <param name="decimals">小數點第幾位</param>
      /// <returns></returns>
      public static decimal Ceiling(this decimal? originValue, int decimals = 0) {
         if (originValue == null) return 0;

         var temp = 1;
         if (decimals == 0)
            temp = 1;
         else
            for (int k = 0; k < decimals; k++) {
               temp = temp * 10;
            }

         return Math.Ceiling((decimal)originValue * temp) / temp;

      }
   }
}