using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Etk.Excel.Application;
using Etk.Excel.BindingTemplates;
using Etk.Excel.ContextualMenus;
using Etk.ModelManagement;
using ExcelInterop = Microsoft.Office.Interop.Excel;

[assembly: InternalsVisibleTo("Etk.Tests.Templates.ExcelDna1")]
[assembly: InternalsVisibleTo("Etk.Excel.UI")]
namespace Etk.Excel
{
    /// <summary> 
    /// Framework main class. 
    /// Gateway to all the Etk framework fonctionalities 
    /// </summary>
    [Export]
    public sealed class ETKExcel
    {
        #region properties
        private bool isDisposed;
        private static readonly object syncObj = new object();

        private readonly List<ExcelInterop.Workbook> managedWorkbooks = new List<ExcelInterop.Workbook>();

        [Import(AllowDefault = false)]
        private ExcelApplication excelApplication;
        [Import(AllowDefault = false)]
        private ExcelTemplateManager templateManager;
        [Import(AllowDefault = false)]
        private ContextualMenuManager contextualMenuManager;

        [Import(AllowDefault = false)]
        private ModelDefinitionManager modelDefinitionManager;
        //[Import(AllowDefault = false)]
        //private RequestsManager RequestsManager = null;

        #region singleton
        private static ETKExcel Instance;
        #endregion

        /// <summary>Give acces to the <see cref="IExcelTemplateManager"/> part of the  framework</summary>
        public static IExcelTemplateManager TemplateManager
        {
            get 
            {
                if (Instance == null || Instance.isDisposed)
                    throw new EtkException("'ETKExcel' is not initialyzed or was disposed");
                return Instance.templateManager; 
            }
        }

        /// <summary>Give acces to the <see cref="IExcelApplication"/> part of the  framework</summary>
        public static IExcelApplication ExcelApplication
        {
            get 
            {
                if (Instance == null || Instance.isDisposed)
                    throw new EtkException("'ETKExcel' is not initialyzed or was disposed");
                return Instance.excelApplication; 
            }
        }

        /// <summary>Give acces to the <see cref="IContextualMenuManager"/> part of the  framework</summary>
        public static IContextualMenuManager ContextualMenuManager
        {
            get
            {
                if (Instance == null || Instance.isDisposed)
                    throw new EtkException("'ETKExcel' is not initialyzed or was disposed");
                return Instance.contextualMenuManager;
            }
        }

        /// <summary>Give acces to the <see cref="IModelDefinitionManager"/> part of the  framework</summary>
        public static IModelDefinitionManager ModelDefinitionManager
        {
            get
            {
                if (Instance == null || Instance.isDisposed)
                    throw new EtkException("'ETKExcel' is not initialyzed or was disposed");
                return Instance.modelDefinitionManager;
            }
        }
        #endregion

        #region .ctors

        private ETKExcel()
        {}

        ~ETKExcel()
        {
            InternalDispose();
        }
        #endregion

        #region public methods
        /// <summary>Init the framework. Must be called before any other uses of the framework</summary>
        /// <param name="application">A reference to the current Excel application instance</param>
        public static void Init(ExcelInterop.Application application)
        {
            try
            {
                if (application == null)
                    throw new ArgumentException("the 'application' parameter is mandatory");

                lock (syncObj)
                {
                    if (Instance == null)
                    {
                        // Inject the Excel application reference
                        CompositionManager.Instance.ComposeExportedValue(application);

                        Instance = new ETKExcel();
                        // Compose the current instance
                        CompositionManager.Instance.ComposeParts(Instance);

                        Instance.AddManagedWorkbook(application.ActiveWorkbook);

                        application.WorkbookOpen += Instance.AddManagedWorkbook;
                        application.WorkbookBeforeClose += Instance.OnWorkbookBeforeClose;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new EtkException($"'ETKExcel' initialization failed:{ex.Message}", ex);
            }
        }

        public static void Dispose()
        {
            Instance.InternalDispose();
        }
        #endregion

        #region private methods
        private void AddManagedWorkbook(ExcelInterop.Workbook workbook)
        {
            ExcelInterop.Workbook managedWorkbook = Instance.managedWorkbooks.FirstOrDefault(w => w == workbook);
            if (managedWorkbook == null)
            {
                managedWorkbooks.Add(workbook);
                contextualMenuManager.RegisterWorkbook(workbook);
                //workbook.SheetActivate += OnActivateSheetViewsManagement;
            }
            else
            {
                Marshal.ReleaseComObject(managedWorkbook);
                managedWorkbook = null;
            }
        }

        private void OnWorkbookBeforeClose(ExcelInterop.Workbook workbook, ref bool cancel)
        {
            if (!cancel && workbook.Application.Workbooks.Count >= 1)
                Instance.managedWorkbooks.Remove(workbook);
        }

        private void InternalDispose()
        {
            lock (syncObj)
            {
                if (!isDisposed)
                {
                    templateManager?.Dispose();
                    //  RequestsManager?.Dispose();
                    contextualMenuManager?.Dispose();

                    managedWorkbooks.Clear();

                    if (excelApplication != null)
                    {
                        if (excelApplication.Application != null)
                        {
                            excelApplication.Application.WorkbookBeforeClose -= OnWorkbookBeforeClose;
                            excelApplication.Application.WorkbookOpen -= Instance.AddManagedWorkbook;
                        }
                        excelApplication.Dispose();
                    }

                    isDisposed = true;
                    Instance = null;
                }
            }
        }
        #endregion
    }
}
