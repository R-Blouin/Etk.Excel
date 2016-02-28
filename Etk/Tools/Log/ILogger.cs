namespace Etk.Excel.UI.Log
{
    using System;

    /// <summary>
    /// Application logging interface
    /// </summary>
    public interface ILogger
    {
        LogType GetLogLevel();

        /// <summary>
        /// Log a message using the application logger.
        /// </summary>
        /// <param name="logType">Criticality of the message.</param>
        /// <param name="message">Message to log</param>
        void Log(LogType logType, string message);

        /// <summary>
        /// Log a message using the application logger and <see cref="M:System.String.Format(System.String,System.Object)"/>
        /// </summary>
        /// <param name="logType">Criticality of the message.</param>
        /// <param name="messageFormat">Format of the message to log.</param>
        /// <param name="o">Parameter of the message.</param>
        /// <remarks> 
        /// <para>The message is formatted using the <c>String.Format</c> method.</para>
        /// <para>See <see cref="M:System.String.Format(System.String,System.Object[])"/> for details of the syntax of the format string and the behavior of the formatting. </para> 
        /// </remarks>
        void LogFormat(LogType logType, string messageFormat, object o);

        /// <summary>
        /// Log a message using the application logger and <see cref="M:System.String.Format(System.String,System.Object[])"/>
        /// </summary>
        /// <param name="logType">Criticality of the message.</param>
        /// <param name="messageFormat">Format of the message to log.</param>
        /// <param name="os">Parameters of the message.</param>
        /// <remarks> 
        /// <para>The message is formatted using the <c>String.Format</c> method.</para>
        /// <para>See <see cref="M:System.String.Format(System.String,System.Object[])"/> for details of the syntax of the format string and the behavior of the formatting. </para> 
        /// </remarks>
        void LogFormat(LogType logType, string messageFormat, params object[] os);

        /// <summary>
        /// Log an exception message using the application logger.
        /// </summary>
        /// <param name="logType">Criticality of the message.</param>
        /// <param name="ex">Exception to log.</param>
        /// <param name="message">Message to log.</param>
        void LogException(LogType logType, Exception ex, string message );

        /// <summary>
        /// Log a message using the application logger and <see cref="M:System.String.Format(System.String,System.Object)"/>
        /// </summary>
        /// <param name="logType">Criticality of the message.</param>
        /// <param name="ex">Exception to log.</param>
        /// <param name="messageFormat">Format of the message to log.</param>
        /// <param name="o">Parameter of the message.</param>
        /// <remarks> 
        /// <para>The message is formatted using the <c>String.Format</c> method.</para>
        /// <para>See <see cref="M:System.String.Format(System.String,System.Object[])"/> for details of the syntax of the format string and the behavior of the formatting. </para> 
        /// </remarks>
        void LogExceptionFormat(LogType logType, Exception ex, string messageFormat, object o);

        /// <summary>
        /// Log a message using the application logger and <see cref="M:System.String.Format(System.String,System.Object[])"/>
        /// </summary>
        /// <param name="logType">Criticality of the message.</param>
        /// <param name="ex">Exception to log.</param>
        /// <param name="messageFormat">Format of the message to log.</param>
        /// <param name="os">Parameters of the message.</param>
        /// <remarks> 
        /// <para>The message is formatted using the <c>String.Format</c> method.</para>
        /// <para>See <see cref="M:System.String.Format(System.String,System.Object[])"/> for details of the syntax of the format string and the behavior of the formatting. </para> 
        /// </remarks>
        void LogExceptionFormat(LogType logType, Exception ex, string messageFormat, params object[] os);
    }
}