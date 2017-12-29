using System;
using System.Runtime.InteropServices;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.Extensions
{
    public static class RangeExtension
    {
        #region Static Fields

        /// <summary>
        /// The missing.
        /// </summary>
        private static readonly object missing = Type.Missing;
        #endregion

        #region Public static Methods
        /// <summary>Return the intersection range.</summary>
        public static ExcelInterop.Range IntersectHelper(this ExcelInterop.Range me, ExcelInterop.Range target)
        {
            ExcelInterop.Application application = null;
            ExcelInterop.Range ret;
            try
            {
                application = me.Application;
                ret = me.Application.Intersect(me, target);
                return ret;
            }
            finally
            {
                if (application != null)
                    Marshal.ReleaseComObject(application);
            }
        }

        /// <summary> Return true if an intersection exists between the current range and the one passed in parameter.</summary>
        public static bool IsInRange(this ExcelInterop.Range me, ExcelInterop.Range target)
        {
            ExcelInterop.Range inter = me.IntersectHelper(target);
            return inter != null && inter.Cells.Count != 0;
        }
        #endregion
    }
}
