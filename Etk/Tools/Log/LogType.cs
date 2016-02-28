namespace Etk.Excel.UI.Log
{
    /// <summary>
    /// Criticality of the log
    /// </summary>
    public enum LogType
    {
        /// <summary>Level of Criticality None: No logging</summary>
        None,
        /// <summary>Level of Criticality Fatal</summary>
        Fatal,
        /// <summary>Level of Criticality Error</summary>
        Error,
        /// <summary>Level of Criticality Warn</summary>
        Warn,
        /// <summary>Level of Criticality Info</summary>
        Info,
        /// <summary>Level of Criticality Debug</summary>
        Debug,
    }
}
