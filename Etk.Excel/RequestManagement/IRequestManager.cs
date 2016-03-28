namespace Etk.Excel.RequestManagement
{
    public interface IRequestManager
    {
        object GDA(Microsoft.Office.Interop.Excel.Range caller, string dataAccessor, object[] parameters);
    }
}
