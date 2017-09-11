using System;
using System.Diagnostics;

namespace Etk.Tools.Log
{
    public class DefaultLogger : ILogger
    {
        public LogType GetLogLevel() 
        { return LogType.Debug; }

        public void Log(LogType logType, string message)
        {
            Debug.WriteLine($"'{logType}'. {message ?? string.Empty}");        
        }

        public void LogFormat(LogType logType, string message, object o)
        {
            try
            {
                if (message != null)
                    message = string.Format(message, o);
                Debug.WriteLine($"'{logType}'. {message ?? string.Empty}");
            }
            catch
            {
                Debug.WriteLine($"'Error'. Can't log '{message ?? string.Empty}'");
            }
        }

        public void LogFormat(LogType logType, string message, object[] os)
        {
            try
            {
                if (message != null)
                    message = string.Format(message, os);
                Debug.WriteLine($"'{logType}'. {message ?? string.Empty}");
            }
            catch
            {
                Debug.WriteLine($"'Error'. Can't log '{message ?? string.Empty}'");
            }
        }

        public void LogException(LogType logType, Exception ex, string message)
        {
            try
            {
                Debug.WriteLine($"'{logType}'. {message ?? string.Empty}. {ex?.Message ?? string.Empty}");
            }
            catch
            {
                Debug.WriteLine($"'Error'. Can't log '{message ?? string.Empty}'");
            }
        }

        public void LogExceptionFormat(LogType logType, Exception ex, string message, object o)
        {
            try
            {
                if (message != null)
                    message = string.Format(message, o);
                Debug.WriteLine($"'{logType}'. {message ?? string.Empty}. {ex?.Message ?? string.Empty}");
            }
            catch
            {
                Debug.WriteLine($"'Error'. Can't log '{message ?? string.Empty}'");
            }
        }

        public void LogExceptionFormat(LogType logType, Exception ex, string message, object[] os)
        {
            try
            {
                if (message != null)
                    message = string.Format(message, os);
                Debug.WriteLine($"'{logType}'. {message ?? string.Empty}. {ex?.Message ?? string.Empty}");
            }
            catch
            {
                Debug.WriteLine($"'Error'. Can't log '{message ?? string.Empty}'");
            }
        }
    }
}
