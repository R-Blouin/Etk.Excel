//[assembly: log4net.Config.XmlConfigurator()]

using System;
using System.ComponentModel.Composition;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Text;
using Etk.Tools.Log;
using log4net;
using log4net.Config;

namespace Etk.Log4NetWrapper
{
    /// <summary>
    /// <para>Log4Net implementation of 'Etk.Excel.UI.Log.ILogger'.</para>
    /// <para>If Log4Net is not yet configurated for the current process, the class looks for its configuration:</para>
    /// <para>&#160;&#160;First: In the file whose the path is given by the 'Log4NetWrapperConfigFile' key of the 'App.Config.AppSettings'.</para>
    /// <para>&#160;&#160;Second: In the App.Config, in the Log4Net section.</para>
    /// <para>If not configuration found: The logger will use a default configuration file that uses a 'RollingFileAppender' that will log in the 'Etk.Excel.UI.Log' file.</para>
    /// </summary>
    [Export(typeof(ILogger))]
    public class Logger : ILogger
    {
        static readonly ILog iLog = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType.Name);

        private readonly LogType logLevel;

        #region .ctors
        /// <summary>
        /// Default constructor.
        /// </summary>
        public Logger()
        {
            LoadConfiguration();

            if (iLog.IsDebugEnabled)
                logLevel = LogType.Debug;
            else if (iLog.IsWarnEnabled)
                logLevel = LogType.Warn;
            else if (iLog.IsInfoEnabled)
                logLevel = LogType.Info;
            else if (iLog.IsErrorEnabled)
                logLevel = LogType.Error;
            else if (iLog.IsFatalEnabled)
                logLevel = LogType.Fatal;
            else
                logLevel = LogType.None;
        }
        #endregion

        #region ILog Membres
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public LogType GetLogLevel()
        {
            return logLevel;
        }

        /// <summary>
        /// Log message
        /// </summary>
        /// <param name="logType"></param>
        /// <param name="message"></param>
        public void Log(LogType logType, string message)
        {
            Log4NetLog(logType, null, message, null);
        }

        /// <summary>
        /// Log message with format
        /// </summary>
        /// <param name="logType"></param>
        /// <param name="messageFormat"></param>
        /// <param name="o"></param>
        public void LogFormat(LogType logType, string messageFormat, object o)
        {
            Log4NetLog(logType, null, messageFormat, new [] {o});
        }

        /// <summary>
        /// Log message with format
        /// </summary>
        /// <param name="logType"></param>
        /// <param name="messageFormat"></param>
        /// <param name="os"></param>
        public void LogFormat(LogType logType, string messageFormat, object[] os)
        {
            Log4NetLog(logType, null, messageFormat, os);
        }

        /// <summary>
        /// Log exception
        /// </summary>
        /// <param name="logType"></param>
        /// <param name="ex"></param>
        /// <param name="message"></param>
        public void LogException(LogType logType, Exception ex, string message)
        {
            Log4NetLog(logType, ex, message, null);
        }

        /// <summary>
        /// Log exception with format
        /// </summary>
        /// <param name="logType"></param>
        /// <param name="ex"></param>
        /// <param name="messageFormat"></param>
        /// <param name="o"></param>
        public void LogExceptionFormat(LogType logType, Exception ex, string messageFormat, object o)
        {
            Log4NetLog(logType, ex, messageFormat, new [] {o});
        }

        /// <summary>
        /// Log exception with format
        /// </summary>
        /// <param name="logType"></param>
        /// <param name="ex"></param>
        /// <param name="messageFormat"></param>
        /// <param name="os"></param>
        public void LogExceptionFormat(LogType logType, Exception ex, string messageFormat, object[] os)
        {
            Log4NetLog(logType, ex, messageFormat, os);
        }
        #endregion

        #region private methods
        #endregion

        #region private methods
        private void Log4NetLog(LogType logType, Exception ex, string message, object[] os)
        {
            try
            {
                if (!string.IsNullOrEmpty(message) || ex != null)
                {
                    switch (logType)
                    {
                        case LogType.Fatal:
                            if (iLog.IsFatalEnabled)
                            {
                                if (os != null && os.Length > 0)
                                    message = string.Format(message, os);
                                iLog.Fatal(message, ex);
                            }
                            break;
                        case LogType.Error:
                            if (iLog.IsErrorEnabled)
                            {
                                if (os != null && os.Length > 0)
                                    message = string.Format(message, os);
                                iLog.Error(message, ex);
                            }
                            break;
                        case LogType.Warn:
                            if (iLog.IsWarnEnabled)
                            {
                                if (os != null && os.Length > 0)
                                    message = string.Format(message, os);
                                iLog.Warn(message, ex);
                            }
                            break;
                        case LogType.Info:
                            if (iLog.IsInfoEnabled)
                            {
                                if (os != null && os.Length > 0)
                                    message = string.Format(message, os);
                                iLog.Info(message, ex);
                            }
                            break;
                        case LogType.Debug:
                            if (iLog.IsDebugEnabled)
                            {
                                if (os != null && os.Length > 0)
                                    message = string.Format(message, os);
                                iLog.Debug(message, ex);
                            }
                            break;
                    }
                }
            }
            catch (Exception exx)
            {
                string exceptionMessage = $"Cannot log the message '{message ?? string.Empty}'.";
                iLog.Error(exceptionMessage, exx);
            }
        }

        private void LoadConfiguration()
        {
            if (! log4net.LogManager.GetRepository().Configured)
            {
                string configFilePath = ConfigurationManager.AppSettings.Get("Log4NetWrapperConfigFile");
                try
                {
                    if(! string.IsNullOrEmpty(configFilePath))
                    {
                        XmlConfigurator.Configure(new FileInfo(configFilePath));
                        this.LogFormat(LogType.Info, "'Log4NetWrapper' initialyzed from configuration file '{0}'.", configFilePath);
                    }
                    else
                    {
                        XmlConfigurator.Configure();
                        if(log4net.LogManager.GetRepository().Configured)
                            Log(LogType.Info, "'Log4NetWrapper' initialyzed from the 'Log4Net' 'App.Config' configuration section.");
                        else
                            UseDefaultConfiguration();
                    }
                }
                catch (Exception ex)
                {
                    UseDefaultConfiguration();
                    LogException(LogType.Error, ex, "'Log4NetWrapper' initialyzation failed. Its embedded default configuration will be used.");
                }
            }
        }
        
        private void UseDefaultConfiguration()
        {
            try
            {
                const string defaultConfig = @"<log4net>
                                                <appender name=""RollingFileAppender"" UnderlyingType=""log4net.Appender.RollingFileAppender"">
                                                  <file value=""EtkLogs.Log"" />
                                                  <appendToFile value=""true"" />
                                                  <rollingStyle value=""Date"" />
                                                  <datePattern value=""yyyyMMdd"" />
                                                  <maxSizeRollBackups value=""5"" />
                                                  <staticLogFileName value=""true"" />
                                                  <layout UnderlyingType=""log4net.Layout.PatternLayout"">
                                                    <conversionPattern value=""%date %-5level - %message%newline""/>
                                                  </layout>
                                                </appender>
                                                <root>
                                                <level value=""DEBUG"" />
                                                <appender-ref ref=""RollingFileAppender"" />
                                                </root>
                                                </log4net>";

                MemoryStream memoryStream = new MemoryStream(UnicodeEncoding.Default.GetBytes(defaultConfig));
                XmlConfigurator.Configure(memoryStream);
                this.Log(LogType.Info, "Using Log4net default configuration.");
            }
            catch (Exception ex)
            {
                new EtkException("Log4Net: Cannot load the default configuration", ex);
            }
        }  
        #endregion
    }
}
