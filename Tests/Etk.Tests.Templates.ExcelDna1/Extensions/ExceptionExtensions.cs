using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Etk.Tests.Templates.ExcelDna1.Extensions
{
    static public class ExceptionExtensions
    {
        public static string ToString(this Exception ex, string message)
        {
            StringBuilder sb = new StringBuilder(message);
            while (ex != null)
            {
                sb.AppendFormat("\r\n{0}", ex.Message);
                ex = ex.InnerException;
            }
            return sb.ToString();
        }
    }
}
