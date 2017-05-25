using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Runtime.CompilerServices;
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
        private bool isDisposed = false;
        private static readonly object syncObj = new object();

        private readonly List<ExcelInterop.Workbook> managedWorkbooks = new List<ExcelInterop.Workbook>();

        [Import(AllowDefault = false)]
        private ExcelApplication excelApplication = null;
        [Import(AllowDefault = false)]
        private ExcelTemplateManager templateManager = null;
        [Import(AllowDefault = false)]
        private ContextualMenuManager contextualMenuManager = null;

        [Import(AllowDefault = false)]
        private ModelDefinitionManager modelDefinitionManager = null;
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
        private ETKExcel(ExcelInterop.Application application)
        {
            if (application == null)
                throw new EtkException("ETKExcel initialization: the 'application' parameter is mandatory");

            // Init System.Windows.Application (Wpf)
            ////////////////////////////////////////
            //if (System.Windows.Application.Current == null)
            //{
            //    new System.Windows.Application();
            //    System.Windows.Application.Current.ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown;
            //}
        }

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
                        // Init ETKExcel
                        ////////////////
                        Instance = new ETKExcel(application);

                        // Inject the Excel application reference
                        CompositionManager.Instance.ComposeExportedValue(application);
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
                throw new EtkException(string.Format("'ETKExcel' initialization failed:{0}", ex.Message), ex);
            }
        }

        public static void Dispose()
        {
            Instance.InternalDispose();
        }
        #endregion

        #region internal methods

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
            managedWorkbook = null;
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
                    if (templateManager != null)
                        templateManager.Dispose();
                    //@@if (RequestsManager != null)
                    //    RequestsManager.Dispose();
                    if (contextualMenuManager != null)
                        contextualMenuManager.Dispose();

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
