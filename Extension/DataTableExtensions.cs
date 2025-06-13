using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;

namespace Common.Extension {
   public static class DataTableExtensions {
      public static DataTable ToDataTable<T>(this IEnumerable<T> list) {
         if (list == null) throw new ArgumentNullException(nameof(list));
         
         Type type = typeof(T);
         var properties = type.GetProperties();

         DataTable dataTable = new DataTable();
         foreach (PropertyInfo info in properties) {
            if (info == null) continue;
            dataTable.Columns.Add(new DataColumn(info.Name, Nullable.GetUnderlyingType(info.PropertyType) ?? info.PropertyType));
         }

         foreach (T entity in list) {
            if (entity == null) continue;
            
            object[] values = new object[properties.Length];
            for (int i = 0; i < properties.Length; i++) {
               if (properties[i] == null) continue;
               values[i] = properties[i].GetValue(entity) ?? DBNull.Value;
            }

            dataTable.Rows.Add(values);
         }

         return dataTable;
      }

      public static IList<T> ToList<T>(this DataTable table) where T : new() {
         if (table == null) throw new ArgumentNullException(nameof(table));
         
         IList<PropertyInfo> properties = typeof(T).GetProperties().ToList();
         IList<T> result = new List<T>();

         //取得DataTable所有的row data
         foreach (var row in table.Rows) {
            if (row == null) continue;
            
            var item = MappingItem<T>((DataRow)row, properties);
            result.Add(item);
         }

         return result;
      }

      private static T MappingItem<T>(DataRow row, IList<PropertyInfo> properties) where T : new() {
         if (row == null) throw new ArgumentNullException(nameof(row));
         if (properties == null) throw new ArgumentNullException(nameof(properties));
         
         T item = new T();
         foreach (var property in properties) {
            if (property == null) continue;
            if (row.Table.Columns.Contains(property.Name)) {
               //針對欄位的型態去轉換
               if (property.PropertyType == typeof(DateTime)) {
                  var dt = new DateTime();
                  if (DateTime.TryParse(row[property.Name].ToString(), out dt)) {
                     property.SetValue(item, dt, null);
                  } else {
                     property.SetValue(item, null, null);
                  }
               } else if (property.PropertyType == typeof(decimal)) {
                  var val = new decimal();
                  decimal.TryParse(row[property.Name].ToString(), out val);
                  property.SetValue(item, val, null);
               } else if (property.PropertyType == typeof(double)) {
                  var val = new double();
                  double.TryParse(row[property.Name].ToString(), out val);
                  property.SetValue(item, val, null);
               } else if (property.PropertyType == typeof(int)) {
                  var val = new int();
                  int.TryParse(row[property.Name].ToString(), out val);
                  property.SetValue(item, val, null);
               } else {
                  if (row[property.Name] != DBNull.Value) {
                     property.SetValue(item, row[property.Name], null);
                  }
               }
            }
         }

         return item;
      }

      /// <summary>
      /// 檢查每一筆Row.RowState是否為Added/Modified/Deleted,有一筆是就回傳true,否則回傳false
      /// </summary>
      /// <param name="table"></param>
      /// <returns></returns>
      public static bool CheckChange(this DataTable table) {
         if (table == null) throw new ArgumentNullException(nameof(table));
         
         foreach (DataRow row in table.Rows) {
            if (row.RowState == DataRowState.Added || row.RowState == DataRowState.Modified || row.RowState == DataRowState.Deleted) {
               return true;
            }
         }
         return false;
      }
   }
}