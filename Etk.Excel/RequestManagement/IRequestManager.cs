namespace Etk.Excel.RequestManagement
{
    using System;
    using Excel = Microsoft.Office.Interop.Excel;
    
    public interface IRequestManager
    {
        object GDA(Excel.Range caller, string dataAccessor, object[] parameters);
    }
}
