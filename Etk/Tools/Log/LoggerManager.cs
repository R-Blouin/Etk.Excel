namespace Etk.Excel.UI.Log
{
    using System.ComponentModel.Composition;

    /// <summary> Application logging manager (lazy singleton).</summary>
    [Export]
    class LoggerManager
    {
        public ILogger Instance
        { get; private set; }

        [ImportingConstructor]
        private LoggerManager([Import(AllowDefault = true)] ILogger logger)
        {
            Instance = logger == null ? new DefaultLogger() : logger;
        }
    }
}
