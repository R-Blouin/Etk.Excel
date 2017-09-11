using System.ComponentModel.Composition;

namespace Etk.Tools.Log
{
    /// <summary> Application logging manager (lazy singleton).</summary>
    [Export]
    class LoggerManager
    {
        public ILogger Instance
        { get; private set; }

        [ImportingConstructor]
        private LoggerManager([Import(AllowDefault = true)] ILogger logger)
        {
            Instance = logger ?? new DefaultLogger();
        }
    }
}
