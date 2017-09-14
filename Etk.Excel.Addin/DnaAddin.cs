using ExcelDna.ComInterop;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Etk.Excel.Addin
{
    [ComVisible(false)]
    class DnaAddin : IExcelAddIn
    {
        private readonly string etkTlbName = "Etk.Excel.Addin.tlb";
        private readonly string etkXllName = "Etk_Excel_Addin";

        public void AutoOpen()
        {
            Microsoft.Office.Interop.Excel.Application excelApplication = ExcelDnaUtil.Application as Microsoft.Office.Interop.Excel.Application;
            ETKExcel.Init(excelApplication);

            ComServer.DllRegisterServer();

            //excelApplication.WorkbookOpen += (wb) => Register(wb);
            //excelApplication.WorkbookBeforeClose += (Workbook wb, ref bool cancel) => UnRegisterTlb(wb);

            if (excelApplication.ActiveWorkbook != null)
                Register(excelApplication.ActiveWorkbook);
        }

        public void AutoClose()
        {
            ComServer.DllUnregisterServer();
        }

        private void Register(Workbook workbook)
        {
            UnRegisterTlb(workbook);
            RegisterTlb(workbook);
        }

        private void UnRegisterTlb(Workbook workbook)
        {
            try
            {
                foreach (Microsoft.Vbe.Interop.Reference reference in workbook.VBProject.References)
                {
                    if (reference.Name.Equals(etkXllName))
                    {
                        workbook.VBProject.References.Remove(reference);
                        break;
                    }
                }
            }
            catch
            { }
        }

        private void RegisterTlb(Workbook workbook)
        {
            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            string tlbPath = Path.Combine(Path.GetDirectoryName(assemblyPath), etkTlbName);
            workbook.VBProject.References.AddFromFile(tlbPath);
        }
    }
}
