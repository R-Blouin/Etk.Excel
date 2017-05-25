namespace Etk.Excel.UI.RequestManagement
{
    using ExcelInterop = Microsoft.Office.Interop.Excel;

    public interface IRequestManager
    {
        object GDA(ExcelInterop.Range caller, string dataAccessor, object[] parameters);
    }
}
