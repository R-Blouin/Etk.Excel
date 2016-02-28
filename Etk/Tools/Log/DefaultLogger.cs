namespace Etk.Excel.UI.Log
{
    using System;
    using System.Diagnostics;

    public class DefaultLogger : ILogger
    {
        public LogType GetLogLevel() 
        { return LogType.Debug; }

        public void Log(LogType logType, string message)
        {
            Debug.WriteLine(string.Format("'{0}'. {1}", logType.ToString(), message ?? string.Empty));        
        }

        public void LogFormat(LogType logType, string message, object o)
        {
            try
            {
                if (message != null)
                    message = string.Format(message, o);
                Debug.WriteLine(string.Format("'{0}'. {1}", logType.ToString(), message ?? string.Empty));
            }
            catch
            {
                Debug.WriteLine(string.Format("'Error'. Can't log '{1}'", message ?? string.Empty));
            }
        }

        public void LogFormat(LogType logType, string message, object[] os)
        {
            try
            {
                if (message != null)
                    message = string.Format(message, os);
                Debug.WriteLine(string.Format("'{0}'. {1}", logType.ToString(), message ?? string.Empty));
            }
            catch
            {
                Debug.WriteLine(string.Format("'Error'. Can't log '{1}'", message ?? string.Empty));
            }
        }

        public void LogException(LogType logType, Exception ex, string message)
        {
            try
            {
                Debug.WriteLine(string.Format("'{0}'. {1}. {2}", logType.ToString(), message ?? string.Empty, ex == null ? string.Empty : ex.Message));
            }
            catch
            {
                Debug.WriteLine(string.Format("'Error'. Can't log '{1}'", message ?? string.Empty));
            }
        }

        public void LogExceptionFormat(LogType logType, Exception ex, string message, object o)
        {
            try
            {
                if (message != null)
                    message = string.Format(message, o);
                Debug.WriteLine(string.Format("'{0}'. {1}. {2}", logType.ToString(), message ?? string.Empty, ex == null ? string.Empty : ex.Message));
            }
            catch
            {
                Debug.WriteLine(string.Format("'Error'. Can't log '{1}'", message ?? string.Empty));
            }
        }

        public void LogExceptionFormat(LogType logType, Exception ex, string message, object[] os)
        {
            try
            {
                if (message != null)
                    message = string.Format(message, os);
                Debug.WriteLine(string.Format("'{0}'. {1}. {2}", logType.ToString(), message ?? string.Empty, ex == null ? string.Empty : ex.Message));
            }
            catch
            {
                Debug.WriteLine(string.Format("'Error'. Can't log '{1}'", message ?? string.Empty));
            }
        }
    }
}
