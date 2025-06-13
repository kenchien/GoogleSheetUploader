using System;

namespace GoogleSheetUploader.Extension {
    public static class DateTimeExtension {
        public static string ToSafeString(this DateTime? value, string format = "yyyy/MM/dd HH:mm:ss") {
            return value?.ToString(format) ?? string.Empty;
        }

        public static DateTime ToDateTimeSafe(this DateTime? value, DateTime defaultValue = default) {
            return value ?? defaultValue;
        }

        public static bool IsNullOrEmpty(this DateTime? value) {
            return !value.HasValue;
        }

        public static DateTime? ToNullableDateTime(this string? value) {
            if (string.IsNullOrWhiteSpace(value)) return null;
            return DateTime.TryParse(value, out var result) ? result : null;
        }

        public static DateTime ToDateTime(this string? value, DateTime defaultValue = default) {
            if (string.IsNullOrWhiteSpace(value)) return defaultValue;
            return DateTime.TryParse(value, out var result) ? result : defaultValue;
        }
    }
} 