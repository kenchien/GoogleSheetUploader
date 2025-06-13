using System;
using log4net;
using log4net.Config;

namespace Common.Helper {
    public static class LogHelper {
        private static readonly ILog _log = LogManager.GetLogger(typeof(LogHelper));

        static LogHelper() {
            XmlConfigurator.Configure(new FileInfo("log4net.config"));
        }

        public static void Info(string message) {
            _log.Info(message);
        }

        public static void Error(string message, Exception? ex = null) {
            if (ex != null) {
                _log.Error(message, ex);
            } else {
                _log.Error(message);
            }
        }

        public static void Debug(string message) {
            _log.Debug(message);
        }

        public static void Warn(string message) {
            _log.Warn(message);
        }

        public static void Warning(string message) {
            _log.Warn(message);
        }
    }
} 