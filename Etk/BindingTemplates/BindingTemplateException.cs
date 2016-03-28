using System;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization;

[assembly: InternalsVisibleTo("Etk.UI")]

namespace Etk.BindingTemplates
{
    [Serializable]
    public class BindingTemplateException : EtkException
    {
        public BindingTemplateException() : base() { }

        public BindingTemplateException(string message) : base(message) 
        { }
        
        public BindingTemplateException(string message, Exception innerException)
                                       : base(message, innerException) 
        { }

        public BindingTemplateException(string message, bool log) 
                                       : base(message, log)
        { }

        public BindingTemplateException(string message, Exception innerException, bool log)
                                       : base(message, innerException, log)
        { }


        protected BindingTemplateException(SerializationInfo info, StreamingContext context)
                                          : base(info, context) 
        { }
    }
}
