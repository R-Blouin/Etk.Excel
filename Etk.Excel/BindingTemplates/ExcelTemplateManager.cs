using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Runtime.InteropServices;
using Etk.BindingTemplates;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.EventCallBacks;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.BindingTemplates.Views;
using Etk.Excel.Application;
using Etk.Excel.BindingTemplates.Decorators;
using Etk.Excel.BindingTemplates.Definitions;
using Etk.Excel.BindingTemplates.SortSearchAndFilter;
using Etk.Excel.BindingTemplates.Views;
using Etk.Excel.ContextualMenus;
using Etk.Excel.TemplateManagement;
using Etk.Tools.Extensions;
using Etk.Tools.Log;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.BindingTemplates
{
    [Export]
    [PartCreationPolicy(CreationPolicy.Shared)]
    class ExcelTemplateManager : IExcelTemplateManager, IDisposable
    {
        private const string TEMPLATE_START_FORMAT = "<Template*Name='{0}'";

        private bool disposed;
        private static object syncRoot = new object();

        #region attributes and properties

        internal ExcelNotifyPropertyManager ExcelNotifyPropertyManager
        { get; private set; }

        internal ExcelApplication ExcelApplication
        { get; private set; }

        internal EventCallbacksManager CallbacksManager
        { get; private set; }

        private ILogger log = Logger.Instance;
        private Dictionary<ExcelInterop.Worksheet, List<ExcelTemplateView>> viewsBySheet = new Dictionary<ExcelInterop.Worksheet, List<ExcelTemplateView>>();

        private ContextualMenuManager contextualMenuManager;
        private ExcelDecoratorsManager excelDecoratorsManager;
        private BindingTemplateManager bindingTemplateManager;

        private TemplateContextualMenuManager templateContextualMenuManager;
        private SortSearchAndFilterMenuManager sortSearchAndFilterMenuManager; 
        #endregion

        #region .ctors
        [ImportingConstructor]
        public ExcelTemplateManager([Import] ExcelApplication excelApplication,
                                    [Import] ContextualMenuManager contextualMenuManager,
                                    [Import] ExcelDecoratorsManager excelDecoratorsManager,
                                    [Import] EventCallbacksManager eventCallbacksManager,
                                    [Import] BindingTemplateManager bindingTemplateManager)
        {
            if (excelApplication == null)
                throw new EtkException("'ExcelBindingTemplateManager' initialization: the 'application' parameter is mandatory");

            ExcelApplication = excelApplication;
            CallbacksManager = eventCallbacksManager;
            this.excelDecoratorsManager = excelDecoratorsManager;
            this.contextualMenuManager = contextualMenuManager;
            this.bindingTemplateManager = bindingTemplateManager;

            // Create the notify property manager . 
            ExcelNotifyPropertyManager = new ExcelNotifyPropertyManager(ExcelApplication);
            // Create the template contextual menus manager. 
            templateContextualMenuManager = new TemplateContextualMenuManager(contextualMenuManager);
            // Declare the contextual menus activators for the template views. 
            contextualMenuManager.OnContextualMenusRequested += ManageViewsContextualMenu;

            sortSearchAndFilterMenuManager = new SortSearchAndFilterMenuManager();
        }

        ~ExcelTemplateManager()
        {
            Dispose();
        }
        #endregion

        #region private methods
        private ExcelTemplateView CreateView(ExcelInterop.Worksheet sheetContainer, string templateName, ExcelInterop.Worksheet sheetDestination, ExcelInterop.Range firstOutputCell, string clearingCellAddress)
        {
            if (sheetContainer == null)
                throw new ArgumentNullException("Template container sheet is mandatory");
            if (string.IsNullOrEmpty(templateName))
                throw new ArgumentNullException("Template name is mandatory");
            if (sheetDestination == null)
                throw new ArgumentNullException("Template destination sheet is mandatory");
            if (firstOutputCell == null)
                throw new ArgumentNullException("Template first output cell is mandatory");

            ExcelInterop.Range clearingCell = null;
            if (!string.IsNullOrEmpty(clearingCellAddress))
            {
                try
                {
                    ExcelInterop.Range workingCell = sheetDestination.Range[clearingCellAddress];
                    clearingCell = workingCell.Cells[1, 1];
                    workingCell = null;
                }
                catch
                {
                    throw new ArgumentException(string.Format("The clearing cell value '{0}' is not valid", clearingCellAddress));
                }
            }

            ExcelInterop.Range range = sheetContainer.Cells.Find(string.Format(TEMPLATE_START_FORMAT, templateName), Type.Missing, ExcelInterop.XlFindLookIn.xlValues, ExcelInterop.XlLookAt.xlPart, ExcelInterop.XlSearchOrder.xlByRows, ExcelInterop.XlSearchDirection.xlNext, false);
            if (range == null)
                throw new EtkException(string.Format("Cannot find the template '{0}' in sheet '{1}'", templateName.EmptyIfNull(), sheetContainer.Name.EmptyIfNull()));

            string templateDescriptionKey = string.Format("{0}-{1}", sheetContainer.Name, templateName);
            TemplateDefinition templateDefinition = bindingTemplateManager.GetTemplateDefinition(templateDescriptionKey);
            if (templateDefinition == null)
            {
                templateDefinition = ExcelTemplateDefinitionFactory.CreateInstance(templateName, range);
                bindingTemplateManager.RegisterTemplateDefinition(templateDefinition);
            }

            ExcelTemplateView view = new ExcelTemplateView(templateDefinition, sheetDestination, firstOutputCell, clearingCell);
            bindingTemplateManager.AddView(view);
            log.LogFormat(LogType.Debug, "Sheet '{0}', View '{1}'.'{2}' created.", sheetDestination.Name.EmptyIfNull(), sheetContainer.Name.EmptyIfNull(), templateName.EmptyIfNull());
            range = null;

            return view;
        }

        private void RegisterView(ExcelTemplateView view)
        {
            if (view == null)
                return;

            try
            {
                if (!viewsBySheet.ContainsKey(view.SheetDestination))
                {
                    viewsBySheet[view.SheetDestination] = new List<ExcelTemplateView>();

                    view.SheetDestination.Change += OnSheetChange;
                    view.SheetDestination.SelectionChange += OnSelectionChange;
                    view.SheetDestination.BeforeDoubleClick += OnBeforeBoubleClick;

                    ExcelInterop.Workbook book = view.SheetDestination.Parent as ExcelInterop.Workbook;
                    if (book != null)
                    {
                        book.SheetCalculate -= OnSheetCalculate;
                        book.SheetCalculate += OnSheetCalculate;
                        Marshal.ReleaseComObject(book);
                    }
                }
                viewsBySheet[view.SheetDestination].Add(view);
            }
            catch (Exception ex)
            {
                throw new EtkException("View registration failed", ex);
            }
        }

        private void OnSelectionChange(ExcelInterop.Range target)
        {
            //Excel.Range realTarget = target.Cells.Count > 1 ? target.Resize[1, 1] : target;
            ExcelInterop.Range realTarget = target.Cells[1, 1];
            List<ExcelTemplateView> views;
            if (viewsBySheet.TryGetValue(realTarget.Worksheet, out views))
            {
                IEnumerable<ExcelTemplateView> viewToWorkWith = views.Select(v => v).ToList();
                foreach (ExcelTemplateView view in viewToWorkWith)
                {
                    if (view.OnSelectionChange(ExcelApplication, realTarget))
                        break;
                }
            }
            Marshal.ReleaseComObject(realTarget);
            realTarget = null;
        }

        private void OnSheetCalculate(object sheet)
        {
            List<ExcelTemplateView> views = null;
            viewsBySheet.TryGetValue(sheet as ExcelInterop.Worksheet, out views);
            if (views != null)
            {
                foreach (ExcelTemplateView view in views)
                    view.OnSheetCalculate();
            }
        }

        /// <summary>
        /// Manage the views contextual menus (those that are defined in the templates)
        /// </summary>
        private IEnumerable<IContextualMenu> ManageViewsContextualMenu(ExcelInterop.Worksheet sheet, ExcelInterop.Range range)
        {
            List<IContextualMenu> menus = new List<IContextualMenu>();
            if (sheet != null && range != null)
            {
                ExcelInterop.Range targetRange = range.Cells[1, 1];

                lock (syncRoot)
                {
                    List<ExcelTemplateView> views;
                    if (viewsBySheet.TryGetValue(sheet, out views))
                    {
                        if (views != null)
                        {
                            foreach (ExcelTemplateView view in views.Where(v => v.IsRendered).Select(v => v))
                            {
                                ExcelInterop.Range intersect = ExcelApplication.Application.Intersect(view.RenderedRange, targetRange);
                                if (intersect != null)
                                {
                                    IBindingContextItem contextItem = view.GetConcernedContextItem(targetRange);
                                    if (contextItem != null)
                                    {
                                        // User contextual menu
                                        IBindingContextElement currentContextElement = contextItem.ParentElement;
                                        if (currentContextElement != null)
                                        {
                                            IBindingContextElement targetedContextElement = currentContextElement;
                                            do
                                            {
                                                ExcelTemplateDefinitionPart currentTemplateDefinition = currentContextElement.ParentPart.TemplateDefinitionPart as ExcelTemplateDefinitionPart;
                                                if ((currentTemplateDefinition.Parent as ExcelTemplateDefinition).ContextualMenu != null)
                                                {
                                                    ContextualMenu contextualMenu = (currentTemplateDefinition.Parent as ExcelTemplateDefinition).ContextualMenu as ContextualMenu;
                                                    contextualMenu.SetAction(targetRange, currentContextElement, targetedContextElement);
                                                    menus.Insert(0, contextualMenu);
                                                }
                                                currentContextElement = currentContextElement.ParentPart.ParentContext == null ? null : currentContextElement.ParentPart.ParentContext.Parent;
                                            }
                                            while (currentContextElement != null);
                                        }
                                    
                                        // Etk sort, search and filter
                                        IContextualMenu searchSortAndFilterMenu = sortSearchAndFilterMenuManager.GetMenus(view, targetRange, contextItem);
                                        if (searchSortAndFilterMenu != null)
                                            menus.Insert(0, searchSortAndFilterMenu);
                                    }
                                }
                            }
                        }
                    }
                }
                targetRange = null;
            }
            return menus;
        }

        private void OnActivateSheetViewsManagement(object sheet)
        {
            ExcelInterop.Worksheet worksheet = sheet as ExcelInterop.Worksheet;
            try
            {
                lock (syncRoot)
                {
                    List<ExcelTemplateView> views;
                    if (viewsBySheet.TryGetValue(worksheet, out views))
                    {
                        if (views != null)
                        {
                            //@@using (FreezeExcel freeze = new FreezeExcel())
                            {
                                foreach (ExcelTemplateView view in views)
                                    view.OnSheetActivation();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string message = string.Format("Sheet '{0}': failed to render its views. {1}", worksheet == null ? string.Empty : worksheet.Name, ex.Message);
                throw new EtkException(message, ex);
            }
            finally
            {
                worksheet = null;
            }
        }

        /// <summary>Manege the change done on the sheet</summary>
        private void OnSheetChange(ExcelInterop.Range target)
        {
            List<ExcelTemplateView> views;
            bool inError = false;
            lock (syncRoot)
            {
                if (viewsBySheet.TryGetValue(target.Worksheet, out views))
                {
                    if (views != null)
                    {
                        foreach (ExcelTemplateView view in views)
                        {
                            try
                            {
                                if (view.OnSheetChange(ExcelApplication, target))
                                    break;
                            }
                            catch (Exception ex)
                            {
                                string message = string.Format("Sheet '{0}', Template '{1}'. Sheet change failed: '{2}'",
                                                                target.Worksheet.Name, view.TemplateDefinition.Name, ex.Message);
                                log.LogException(LogType.Error, ex, message);
                                inError = true;
                            }
                        }
                    }
                }
            }

            if (inError)
            {
                ExcelInterop.Worksheet worksheet = target.Worksheet;
                string message = string.Format("Sheet '{0}', At least one sheet change failed. Please, checked the log", worksheet.Name);
                throw new EtkException(message);
            }
        }

        /// <summary> MAnage the double click on a cell</summary>
        private void OnBeforeBoubleClick(ExcelInterop.Range target, ref bool cancel)
        {
            ExcelInterop.Range realTarget = target.Cells[1, 1];
            ExcelInterop.Worksheet worksheet = target.Worksheet;
            try
            {
                List<ExcelTemplateView> views;
                if (viewsBySheet.TryGetValue(worksheet, out views))
                {
                    if (views != null)
                    {
                        foreach (ExcelTemplateView view in views)
                        {
                            if (!view.IsDisposed && view.IsRendered)
                            { 
                                if (view.OnBeforeBoubleClick(realTarget, ref cancel))
                                    break;
                            }
                        }
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(worksheet);
                worksheet = null;
            }
        }

        private void ProtectSheet(ExcelTemplateView view)
        {
            if (view == null)
                return;

            view.SheetDestination.Cells.Locked = false;
            view.SheetDestination.Protect(System.Type.Missing, true, true, System.Type.Missing, false, true,
                                          true, true,
                                          false, false,
                                          false,
                                          false, false, false, true,
                                          true);
        }
        #endregion

        #region internal methods
        internal ExcelTemplateDefinition GetTemplateDefinitionFromLink(ExcelTemplateDefinitionPart parent, TemplateLink templateLink)
        {
            try
            {
                string[] tos = templateLink.To.Split('.');
                ExcelInterop.Worksheet sheetContainer = null;
                string templateName;
                if (tos.Count() == 1)
                {
                    sheetContainer = parent.DefinitionCells.Worksheet;
                    templateName = tos[0].EmptyIfNull().Trim();
                }
                else
                {
                    string worksheetContainerName = tos[0].EmptyIfNull().Trim();
                    templateName = tos[1].EmptyIfNull().Trim();

                    ExcelInterop.Worksheet parentWorkSheet = parent.DefinitionCells.Worksheet;
                    ExcelInterop.Workbook workbook = parentWorkSheet.Parent as ExcelInterop.Workbook;
                    if (workbook == null)
                        throw new EtkException("Cannot retrieve the workbook that owned the template destination sheet");

                    List<ExcelInterop.Worksheet> sheets = new List<ExcelInterop.Worksheet>(workbook.Worksheets.Cast<ExcelInterop.Worksheet>());
                    sheetContainer = sheets.FirstOrDefault(s => !string.IsNullOrEmpty(s.Name) && s.Name.Equals(worksheetContainerName));
                    if (sheetContainer == null)
                        throw new EtkException(string.Format("Cannot find the sheet '{0}' in the current workbook", worksheetContainerName), false);

                    Marshal.ReleaseComObject(workbook);
                    workbook = null;
                }

                string templateDescriptionKey = string.Format("{0}-{1}", sheetContainer.Name, templateName);
                ExcelTemplateDefinition templateDefinition = bindingTemplateManager.GetTemplateDefinition(templateDescriptionKey) as ExcelTemplateDefinition;
                if (templateDefinition == null)
                {
                    ExcelInterop.Range range = sheetContainer.Cells.Find(string.Format(TEMPLATE_START_FORMAT, templateName), Type.Missing, ExcelInterop.XlFindLookIn.xlValues, ExcelInterop.XlLookAt.xlPart, ExcelInterop.XlSearchOrder.xlByRows, ExcelInterop.XlSearchDirection.xlNext, false);
                    if (range == null)
                        throw new EtkException(string.Format("Cannot find the template '{0}' in sheet '{1}'", templateName.EmptyIfNull(), sheetContainer.Name.EmptyIfNull()));
                    templateDefinition = ExcelTemplateDefinitionFactory.CreateInstance(templateName, range);
                    bindingTemplateManager.RegisterTemplateDefinition(templateDefinition);

                    range = null;
                }

                Marshal.ReleaseComObject(sheetContainer);
                sheetContainer = null;
                return templateDefinition;
            }
            catch (Exception ex)
            {
                string message = string.Format("Cannot create the template dataAccessor. {0}", ex.Message);
                throw new EtkException(message, false);
            }
        }
        #endregion

        #region IExcelTemplateManager Members
        /// <summary> Implements <see cref="IExcelTemplateManager.AddView"/> </summary> 
        public IExcelTemplateView AddView(ExcelInterop.Worksheet sheetContainer, string templateName, ExcelInterop.Worksheet sheetDestination, ExcelInterop.Range destinationRange, string clearingCell)
        {
            try
            {
                lock (syncRoot)
                {
                    ExcelTemplateView view = CreateView(sheetContainer, templateName, sheetDestination, destinationRange, clearingCell);
                    RegisterView(view);
                    return view;
                }
            }
            catch (Exception ex)
            {
                string message = string.Format("Sheet '{0}', cannot add the View from template '{1}.{2}'", sheetDestination != null ? sheetDestination.Name.EmptyIfNull() : string.Empty
                                                                                                                    , sheetContainer != null ? sheetContainer.Name.EmptyIfNull() : string.Empty
                                                                                                                    , templateName.EmptyIfNull());
                Logger.Instance.LogException(LogType.Error, ex, message);
                throw new EtkException(message, ex);
                //ExcelApplication.DisplayException(null, message, ex);
                //return null;
            }
        }

        /// <summary> Implements <see cref="IExcelTemplateManager.AddView"/> </summary> 
        public IExcelTemplateView AddView(string sheetTemplateName, string templateName, string sheetDestinationName, string destinationRange, string clearingCell)
        {
            ExcelInterop.Workbook workbook = null;
            ExcelInterop.Worksheet sheetContainer = null;
            ExcelInterop.Worksheet sheetDestination = null;
            try
            {
                if (string.IsNullOrEmpty(sheetTemplateName))
                    throw new ArgumentNullException("the sheet container name is mandatory");

                if (sheetDestinationName == null)
                    throw new ArgumentNullException("Destination sheet name is mandatory");

                workbook = ETKExcel.ExcelApplication.Application.ActiveWorkbook;

                sheetContainer = ETKExcel.ExcelApplication.GetWorkSheetFromName(workbook, sheetTemplateName);
                if (sheetContainer == null)
                    throw new ArgumentException(string.Format("Cannot find the Destination sheet '{0}'"), sheetTemplateName);
                sheetDestination = ETKExcel.ExcelApplication.GetWorkSheetFromName(workbook, sheetDestinationName);
                if (sheetDestination == null)
                    throw new ArgumentException(string.Format("Cannot find the Destination sheet '{0}'"), sheetDestinationName);

                ExcelInterop.Range destinationRangeRange = sheetDestination.Range[destinationRange];
                IExcelTemplateView view = AddView(sheetContainer, templateName, sheetDestination, destinationRangeRange, clearingCell);
                return view;
            }
            catch (Exception ex)
            {
                string message = string.Format("Sheet '{0}', cannot add the View from template '{1}.{2}'", sheetDestination != null ? sheetDestination.Name.EmptyIfNull() : string.Empty
                                                                                                         , sheetTemplateName.EmptyIfNull()
                                                                                                         , templateName.EmptyIfNull());
                Logger.Instance.LogException(LogType.Error, ex, message);
                Logger.Instance.LogException(LogType.Error, ex, message);

                throw new EtkException(message, ex);
                //ExcelApplication.DisplayException(null, message, ex);
                //return null;
            }
            finally
            {
                if (sheetContainer != null)
                {
                    Marshal.ReleaseComObject(sheetContainer);
                    sheetContainer = null;
                }
                if (workbook != null)
                {
                    Marshal.ReleaseComObject(workbook);
                    workbook = null;
                }
            }
        }

        /// <summary> Implements <see cref="IExcelTemplateManager.RemoveView"/> </summary> 
        public void RemoveView(IExcelTemplateView view)
        {
            if (view == null)
                return;

            ExcelTemplateView excelView = view as ExcelTemplateView;
            if (excelView == null)
                return;

            try
            {
                lock (syncRoot)
                {
                    //ExcelInterop.Workbook book = excelView.SheetDestination.Parent as ExcelInterop.Workbook;
                    //if (book != null)
                    //{
                    //    book.SheetCalculate -= OnSheetCalculate;
                    //    Marshal.ReleaseComObject(book);
                    //}

                    ClearView(excelView);

                    KeyValuePair<ExcelInterop.Worksheet, List<ExcelTemplateView>> kvp = viewsBySheet.FirstOrDefault(s => s.Value.FirstOrDefault(v => v.Equals(view)) != null);
                    if (kvp.Key != null && kvp.Value != null && kvp.Value.Count > 0)
                        viewsBySheet[kvp.Key].Remove(excelView);

                    if (log.GetLogLevel() == LogType.Debug)
                        log.LogFormat(LogType.Debug, "View '{0}' from '{1}' removed.", excelView.Ident, excelView.TemplateDefinition.Name);
                    bindingTemplateManager.RemoveView(excelView);
                }
            }
            catch (Exception ex)
            {
                string message = "Remove View failed.";
                Logger.Instance.LogException(LogType.Error, ex, message);
                throw new EtkException(message, ex);
                //ExcelApplication.DisplayException(null, message, ex);
            }
        }

        /// <summary> Implements <see cref="IExcelTemplateManager.RemoveViews"/> </summary> 
        public void RemoveViews(IEnumerable<IExcelTemplateView> views)
        {
            if (views == null)
                return;

            try
            {
                lock (syncRoot)
                {
                    bool success = true;
                    foreach (IExcelTemplateView view in views)
                    {
                        try { RemoveView(view); }
                        catch { success = false; }
                    }
                    if (!success)
                        throw new EtkException("No all views have been removed. Please check the logs.");
                }
            }
            catch (Exception ex)
            {
                string message = "'Remove views' failed.";
                Logger.Instance.LogException(LogType.Error, ex, message);
                throw new EtkException(message, ex);
                //ExcelApplication.DisplayException(null, message, ex);
            }
        }

        /// <summary> Implements <see cref="IExcelTemplateManager.GetSheetViews"/> </summary> 
        public List<IExcelTemplateView> GetSheetViews(ExcelInterop.Worksheet sheet)
        {
            List<IExcelTemplateView> iViews = new List<IExcelTemplateView>();
            try
            {
                if (sheet != null)
                {
                    lock (syncRoot)
                    {
                        List<ExcelTemplateView> views = null;
                        if (viewsBySheet.TryGetValue(sheet, out views))
                        {
                            IEnumerable<ITemplateView> templateViews = bindingTemplateManager.GetAllViews().Where(v => views.Contains(v) && v is ExcelTemplateView);
                            iViews.AddRange(templateViews.Cast<IExcelTemplateView>());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string message = "'GetSheetTemplates' failed";
                Logger.Instance.LogException(LogType.Error, ex, message);
                throw new EtkException(message, ex);
                //ExcelApplication.DisplayException(null, message, ex);
            }
            return iViews;
        }

        /// <summary> Implements <see cref="IExcelTemplateManager.GetSheetViews"/> </summary> 
        public List<IExcelTemplateView> GetActiveSheetViews()
        {
            List<IExcelTemplateView> iViews = new List<IExcelTemplateView>();
            ExcelInterop.Worksheet activeSheet = null;
            try
            {
                activeSheet = ExcelApplication.GetActiveSheet();
                if (activeSheet != null)
                    iViews = GetSheetViews(activeSheet);
            }
            catch (Exception ex)
            {
                string message = "'GetActiveSheetViews' failed";
                Logger.Instance.LogException(LogType.Error, ex, message);
                throw new EtkException(message, ex);
                //ExcelApplication.DisplayException(null, message, ex);
            }
            finally
            {
                if (activeSheet != null)
                    Marshal.ReleaseComObject(activeSheet);
            }
            return iViews;
        }

        /// <summary> Implements <see cref="IExcelTemplateManager.Render"/> </summary> 
        public void Render(IEnumerable<IExcelTemplateView> views)
        {
            if (views == null)
                return;
            try
            {
                lock (syncRoot)
                {
                    if (ExcelApplication.IsInEditMode())
                        ExcelApplication.DisplayMessageBox(null, "'Render' is not allowed: Excel is in Edit mode", System.Windows.Forms.MessageBoxIcon.Warning);
                    else
                    {
                        ExcelInterop.Range selectedRange = ExcelApplication.Application.Selection as ExcelInterop.Range;
                        using (FreezeExcel freezeExcel = new FreezeExcel())
                        {
                            foreach (IExcelTemplateView view in views)
                            {
                                if (view != null)
                                {
                                    ExcelTemplateView excelTemplateView = view as ExcelTemplateView;
                                    try
                                    {
                                        excelTemplateView.SheetDestination.Unprotect(System.Type.Missing);
                                        excelTemplateView.Render();
                                    }
                                    finally
                                    {
                                        ProtectSheet(excelTemplateView as ExcelTemplateView);
                                    }
                                }
                            }
                        }
                        if (selectedRange != null)
                            selectedRange.Select();
                    }
                }
            }
            catch (Exception ex)
            {
                string message = "'Render' failed.";
                Logger.Instance.LogException(LogType.Error, ex, message);
                throw new EtkException(message, ex);
                //ExcelApplication.DisplayException(null, message, ex);
            }
        }

        /// <summary> Implements <see cref="IExcelTemplateManager.Render"/> </summary> 
        public void Render(IExcelTemplateView view)
        {
            if (view != null)
                Render(new IExcelTemplateView[] { view });
        }

        /// <summary> Implements <see cref="IExcelTemplateManager.RenderDataOnly"/> </summary> 
        public void RenderDataOnly(IExcelTemplateView view)
        {
            if (view != null)
                RenderDataOnly(new IExcelTemplateView[] { view });
        }

        /// <summary> Implements <see cref="IExcelTemplateManager.RenderDataOnly"/> </summary> 
        public void RenderDataOnly(IEnumerable<IExcelTemplateView> views)
        {
            if (views == null)
                return;
            try
            {
                lock (syncRoot)
                {
                    if (ExcelApplication.IsInEditMode())
                        ExcelApplication.DisplayMessageBox(null, "'Render data only' is not allowed: Excel is in Edit mode", System.Windows.Forms.MessageBoxIcon.Warning);
                    else
                    {
                        ExcelInterop.Range selectedRange = ExcelApplication.Application.Selection as ExcelInterop.Range;
                        using (FreezeExcel freezeExcel = new FreezeExcel())
                        {
                            foreach (IExcelTemplateView view in views)
                            {
                                if (view != null)
                                {
                                    ExcelTemplateView excelTemplateView = view as ExcelTemplateView;
                                    try
                                    {
                                        excelTemplateView.SheetDestination.Unprotect(System.Type.Missing);
                                        excelTemplateView.RenderDataOnly();
                                    }
                                    finally
                                    {
                                        ProtectSheet(excelTemplateView as ExcelTemplateView);
                                    }
                                }
                            }
                        }
                        if (selectedRange != null)
                            selectedRange.Select();
                    }
                }
            }
            catch (Exception ex)
            {
                string message = "'RenderDataOnly' failed.";
                Logger.Instance.LogException(LogType.Error, ex, message);
                throw new EtkException(message, ex);
                //ExcelApplication.DisplayException(null, message, ex);
            }
        }

        /// <summary> Implements <see cref="IExcelTemplateManager.ClearView"/> </summary> 
        public void ClearView(IExcelTemplateView view)
        {
            if (view != null)
                ClearViews(new IExcelTemplateView[] { view });
        }

        /// <summary> Implements <see cref="IExcelTemplateManager.ClearViews"/> </summary> 
        public void ClearViews(IEnumerable<IExcelTemplateView> views)
        {
            if (views == null)
                return;

            views = views.Where(v => v != null);
            if (views.Count() == 0)
                return;

            try
            {
                lock (syncRoot)
                {
                    if (ExcelApplication.IsInEditMode())
                        ExcelApplication.DisplayMessageBox(null, "'Clear views' is not allowed: Excel is in Edit mode", System.Windows.Forms.MessageBoxIcon.Warning);
                    else
                    {
                        using (FreezeExcel freezeExcel = new FreezeExcel())
                        {
                            foreach (IExcelTemplateView view in views)
                            {
                                ExcelTemplateView excelView = view as ExcelTemplateView;
                                if (excelView != null)
                                {
                                    try
                                    {
                                        excelView.SheetDestination.Unprotect(System.Type.Missing);
                                        excelView.Clear();
                                    }
                                    finally
                                    {
                                        ProtectSheet(excelView as ExcelTemplateView);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string message = "'Clear views' failed.";
                Logger.Instance.LogException(LogType.Error, ex, message);
                throw new EtkException(message, ex);
                //ExcelApplication.DisplayException(null, message, ex);
            }
        }

        /// <summary> Implements <see cref="IExcelTemplateManager.RegisterDecoratorsFromXml"/> </summary> 
        public void RegisterDecoratorsFromXml(string xmLDefinition)
        {
            excelDecoratorsManager.RegisterDecoratorsFromXml(xmLDefinition);
        }

        /// <summary> Implements <see cref="IExcelTemplateManager.RegisterDecorator"/> </summary> 
        public void RegisterDecorator(ExcelRangeDecorator rangeDecorator)
        {
            excelDecoratorsManager.RegisterDecorator(rangeDecorator);
        }

        /// <summary> Implements <see cref="IExcelTemplateManager.RegisterEventCallbacksFromXml"/> </summary> 
        public void RegisterEventCallbacksFromXml(string xmLDefinition)
        {
            CallbacksManager.RegisterEventCallbacksFromXml(xmLDefinition);
        }

        /// <summary> Register Event callback definitions 
        /// <param name="callback">The callback to register</param>
        public void RegisterEventCallback(EventCallback callback)
        {
            CallbacksManager.RegisterCallback(callback);
        }
        #endregion

        public void Dispose()
        {
            lock (syncRoot)
            {
                if (!disposed)
                {
                    disposed = true;

                    //if (viewsBySheet != null)
                    //{
                    //    viewsBySheet.Values.Where(l => l != null)
                    //                       .SelectMany(v => v)
                    //                       .Where(v => v != null)
                    //                       .ToList()
                    //                       .ForEach(v => v.Dispose());
                    //}

                    contextualMenuManager.OnContextualMenusRequested -= ManageViewsContextualMenu;

                    if (templateContextualMenuManager != null)
                    {
                        templateContextualMenuManager.Dispose();
                        templateContextualMenuManager = null;
                    }

                    if (ExcelNotifyPropertyManager != null)
                    {
                        ExcelNotifyPropertyManager.Dispose();
                        ExcelNotifyPropertyManager = null;
                    }
                    ExcelApplication.Dispose();
                    ExcelApplication = null;
                }
            }
        }
    }
}
