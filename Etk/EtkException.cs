namespace Etk
{
    using System;
    using System.Runtime.Serialization;
    using Etk.Excel.UI.Log;

    [Serializable]
    public class EtkException : Exception
    {
        private ILogger log = Logger.Instance;

        public EtkException() : base() { }
        
        public EtkException(string message) 
                                   : base(message) 
        {
            log.LogException(LogType.Error, this, message);
        }

        public EtkException(string message, bool logException)
                                   : base(message)
        {
            if (logException)
                log.LogException(LogType.Error, this, message);
        }


        public EtkException(string message, Exception innerException)
                                   : base(message, innerException)
        {
            log.LogException(LogType.Error, this, base.Message);
        }

        public EtkException(string message, Exception innerException, bool logException)
                                   : base(message, innerException)
        {
            if (logException)
                log.LogException(LogType.Error, this, message);
        }

        protected EtkException(SerializationInfo info, StreamingContext context)
                                      : base(info, context)
        {}
    }
}
