using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.BindingTemplates.Views;
using Etk.Excel.Application;
using Etk.Excel.BindingTemplates.Definitions;
using Etk.Excel.BindingTemplates.Renderer;
using Etk.Excel.BindingTemplates.SortSearchAndFilter;
using Etk.Tools.Log;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Etk.BindingTemplates.Definitions.EventCallBacks;

namespace Etk.Excel.BindingTemplates.Views
{
    public enum AutoFitMode
    {
        None = 0,
        Width = 1,
        Height = 2,
        WidthHeight = 3,
        HeightWidth = 4,
    }

    class SelectionPattern
    {
        public ExcelInterop.XlPattern Pattern;
        public int PatternColor;
        public int PatternColorIndex;
        public int PatternThemeColor;
        public double PatternTintAndShade;

        public SelectionPattern(ref ExcelInterop.Interior interior)
        {
            try
            {
                Pattern = (ExcelInterop.XlPattern)interior.Pattern;
                PatternColor = interior.PatternColor;
                PatternColorIndex = interior.PatternColorIndex;
                PatternThemeColor = interior.PatternThemeColor;
                PatternTintAndShade = interior.PatternTintAndShade;
            }
            catch
            { }
        }
    }

    class ExcelTemplateView : TemplateView, IExcelTemplateView
    {
        #region attributes and properties
        private const int AutoFitMaxIterationCount = 10;
        private readonly ILogger log = Logger.Instance;
        private readonly List<SelectionPattern> currentSelectedRangePattern = new List<SelectionPattern>();
        private ExcelInterop.Range currentSelectedRange;

        internal ExcelInterop.Range CurrentSelectedCell
        { get; private set; }

        internal List<ExcelBindingSearchContextItem> CellsThatContainSearchValue
        { get;  }

        //public event Action<object, object> DataChanged;
        public event Action DataChanged;
        public event Action<bool> BeforeRendering;
        public event Action<bool> AfterRendering;
        public event Action ViewSheetIsDeactivated;

        private event Action viewSheetIsActivated;
        public event Action ViewSheetIsActivated
        {
            add { viewSheetIsActivated += value; }
			remove { viewSheetIsActivated -= value;}
        }

        public AutoFitMode AutoFit
        { get; set; }

        public ExcelInterop.Worksheet ViewSheet
        { get; private set; }

        public ExcelInterop.Range FirstOutputCell
        { get; set; }

        public ExcelInterop.Range ClearingCell
        { get; set; }

        public ExcelRootRenderer Renderer
        { get; private set; }

        public bool IsRendered => Renderer != null && Renderer.RenderedRange != null;

        public ExcelInterop.Range RenderedRange => Renderer?.RenderedRange;

        public RenderedArea RenderedArea => Renderer?.RenderedArea;

        public AccessorParametersManager AccessorParametersManager
        { get; private set; }

        //public ExcelPartRenderer Expander
        //{ get; set; }

        public override string SearchValue
        {
            get { return searchValue; } 
            set 
            {
                searchValue = value;
                foreach (ExcelBindingSearchContextItem ctrl in CellsThatContainSearchValue)
                {
                    try
                    {
                        ctrl.ExecuteSearch = false;
                        ctrl.DestinationRange.Value = searchValue;
                    }
                    finally
                    {
                        ctrl.ExecuteSearch = true;
                    }
                }
            }
        }
        #endregion

        #region .ctors
        public ExcelTemplateView(ITemplateDefinition templateDefinition, ExcelInterop.Worksheet sheetDestination, ExcelInterop.Range firstOutputCell, ExcelInterop.Range clearingCell)
            : base(templateDefinition)
        {
            ViewSheet = sheetDestination;
            FirstOutputCell = firstOutputCell;
            ClearingCell = clearingCell;
            AutoFit = AutoFitMode.WidthHeight;
            CellsThatContainSearchValue = new List<ExcelBindingSearchContextItem>();
        }
        #endregion

        #region public methods
        /// <summary> Clear the execution previous rendering.</summary>
        public override void Clear()
        {
            lock (syncRoot)
            {
                currentSelectedRangePattern.Clear();

                currentSelectedRange = null;
                CurrentSelectedCell = null;
                CellsThatContainSearchValue.Clear();
                //@@ searchValue = null;

                base.Clear();
                if (!IsDisposed && Renderer != null)
                {
                    if (ETKExcel.ExcelApplication.IsInEditMode())
                        throw new COMException("Excel is on Edit mode");
                    Renderer.Clear();
                    if (log.GetLogLevel() == LogType.Debug)
                        log.LogFormat(LogType.Debug, "Sheet '{0}', View '{1}' from '{2}' cleared.", ViewSheet.Name, this.Ident, TemplateDefinition.Name);
                }
            }
        }

        public override void CreateBindingContext(object dataSource)
        {
            lock (syncRoot)
            {
                if (!IsDisposed)
                {
                    try
                    {
                        base.CreateBindingContext(dataSource);

                        if (Renderer != null)
                            Renderer.Dispose();
                        if (dataSource != null)
                            Renderer = new ExcelRootRenderer(this);
                    }
                    catch (Exception ex)
                    {
                        string message = $"Sheet '{ViewSheet.Name}', View '{this.Ident}' from '{TemplateDefinition.Name}' Set data source failed.";
                        throw new EtkException(message, ex, false);
                    }
                }
            }
        }

        public void SetAccessorParameters(IEnumerable<object> parameters)
        {
            lock (syncRoot)
            {
                if (!IsDisposed)
                {
                    if (AccessorParametersManager != null)
                        AccessorParametersManager.Dispose();

                    AccessorParametersManager = new AccessorParametersManager(this, parameters);
                }
            }
        }

        public override void Dispose()
        {
            lock (syncRoot)
            {
                if (!IsDisposed)
                {
                    //Expander = null;
                    if (AccessorParametersManager != null)
                    {
                        AccessorParametersManager.Dispose();
                        AccessorParametersManager = null;
                    }

                    if (Renderer != null)
                    {
                        Renderer.Dispose();
                        Renderer = null;
                    }

                    CellsThatContainSearchValue.Clear();
                    if (ViewSheet != null)
                    {
                        ExcelApplication.ReleaseComObject(ViewSheet);
                        ViewSheet = null;
                    }

                    FirstOutputCell = null;
                    currentSelectedRange = null;
                    ClearingCell = null;

                    base.Dispose();
                }
            }
        }

        public void OnSheetCalculate()
        {
            if (IsRendered)
                Renderer.OnCalculate();
        }

        public void RegisterSearchControl(ExcelBindingSearchContextItem searchControl)
        {
            CellsThatContainSearchValue.Add(searchControl);
        }

        public override void ExecuteSearch()
        {
            using (var freezeExcel = new FreezeExcel(ETKExcel.ExcelApplication.KeepStatusVisible))
            {
                if (Renderer?.BodyPartRenderer?.RenderedArea == null)
                    return;

                var toShowOrHide = new List<KeyValuePair<ExcelInterop.Range, bool>>();

                ExcelInterop.Range firstRange = ViewSheet.Cells[Renderer.BodyPartRenderer.RenderedArea.YPos, Renderer.BodyPartRenderer.RenderedArea.XPos];
                ExcelInterop.Range lastRange = ViewSheet.Cells[Renderer.BodyPartRenderer.RenderedArea.YPos + Renderer.BodyPartRenderer.RenderedArea.Height - 1, Renderer.BodyPartRenderer.RenderedArea.XPos + Renderer.BodyPartRenderer.RenderedArea.Width - 1];
                ExcelInterop.Range renderedRange = ViewSheet.Range[firstRange, lastRange];
                ExcelInterop.Range rowsOrColumns = TemplateDefinition.Orientation == Orientation.Horizontal ? renderedRange.Columns : renderedRange.Cells.Rows;
                if (string.IsNullOrEmpty(SearchValue))
                {
                    foreach (ExcelInterop.Range rowOrColumn in rowsOrColumns)
                        toShowOrHide.Add(new KeyValuePair<ExcelInterop.Range, bool>(rowOrColumn, false));
                }
                else
                {
                    string searchValueUpper = SearchValue.ToUpper();
                    foreach (ExcelInterop.Range rowOrColumn in rowsOrColumns)
                    {
                        bool toHide = true;
                        foreach (ExcelInterop.Range cell in rowOrColumn.Cells)
                        {
                            string cellText;
                            if (cell.MergeCells)
                                cellText = cell.MergeArea[1.1].Text;
                            else
                                cellText = cell.Text;
                            if (!string.IsNullOrEmpty(cellText) && cellText.ToUpper().Contains(searchValueUpper))
                            {
                                toHide = false;
                                break;
                            }
                        }
                        toShowOrHide.Add(new KeyValuePair<ExcelInterop.Range, bool>(rowOrColumn, toHide));
                    }
                }

                foreach (KeyValuePair<ExcelInterop.Range, bool> showOrHide in toShowOrHide)
                {
                    ExcelInterop.Range cells;
                    if (TemplateDefinition.Orientation == Orientation.Horizontal)
                        cells = ViewSheet.Columns[showOrHide.Key.Column];
                    else
                        cells = ViewSheet.Rows[showOrHide.Key.Row];
                    cells.Hidden = showOrHide.Value;
                    cells = null;
                }

                if (string.IsNullOrEmpty(SearchValue))
                    ManageExpander();

                firstRange = null;
                lastRange = null;
                renderedRange = null;
                rowsOrColumns = null;
            }
        }

        public override void SetDataSource(object dataSource)
        {
            try
            {
                if(ViewSheet.ProtectContents)
                    ViewSheet.Unprotect(Type.Missing);

                //searchValue = null;
                CellsThatContainSearchValue.Clear();
                base.SetDataSource(dataSource);
            }
            finally
            {
                ProtectSheet();
            }
        }

        public void Render()
        {
            ETKExcel.TemplateManager.Render(this);
        }

        public void RenderDataOnly()
        {
            ETKExcel.TemplateManager.RenderDataOnly(this);
        }

        public void ClearView()
        {
            ETKExcel.TemplateManager.ClearView(this);
        }

        public void ExecuteAutoFit()
        {
            ExcelInterop.Range range = null;
            switch (AutoFit)
            {
                case AutoFitMode.Width:
                case AutoFitMode.WidthHeight:
                    {
                        range = ViewSheet?.Columns ?? (null != Renderer.RenderedRange && null != ViewSheet.Columns
                                                       ? Renderer.RenderedRange.Columns
                                                       : null);
                        if (null != range)
                        {
                            AutoFitColumns(range);
                            if (AutoFit == AutoFitMode.WidthHeight)
                                AutoFitRows(range);
                        }
                    }
                 break;

                case AutoFitMode.Height:
                case AutoFitMode.HeightWidth:
                    {
                        range = null != ViewSheet && null != ViewSheet.Rows
                                        ? ViewSheet.Rows
                                        : (null != Renderer.RenderedRange && null != ViewSheet.Rows
                                            ? Renderer.RenderedRange.Rows
                                            : null);
                        if (null != range)
                        {
                            AutoFitRows(range);
                            if (AutoFit == AutoFitMode.HeightWidth)
                                AutoFitColumns(range);
                        }
                    }
               break;
            }

            range = null;
        }

        private void AutoFitRows(ExcelInterop.Range rows)
        {
            double previousSize = -2;
            double currentSize = -1;
            var iteration = 0;
            while (AutoFitMaxIterationCount > iteration && currentSize != previousSize)
            {
                iteration++;
                previousSize = rows.Height;
                rows.Rows.AutoFit();
                currentSize = rows.Height;
            }
        }

        private void AutoFitColumns(ExcelInterop.Range columns)
        {
            double previousSize = -2;
            double currentSize = -1;
            var iteration = 0;
            while (AutoFitMaxIterationCount > iteration && currentSize != previousSize)
            {
                iteration++;
                previousSize = columns.Width;
                columns.Columns.AutoFit();
                currentSize = columns.Width;
            }
        }

        public void ProtectSheet()
        {
            ((ExcelApplication) ETKExcel.ExcelApplication).ProtectSheet(ViewSheet);
        }
        #endregion

        #region internal methods
        //internal void ResolveExpander()
        //{
        //    //if (!FilterOwner.HeaderAsExpander)
        //    //    return;

        //    //if (FilterOwner.ExpanderMode == ExpanderMode.Hide)
        //    //{
        //    //    Worksheet worksheet = Expander.OutputRange.Worksheet;
        //    //    Range toShowHide;
        //    //    try
        //    //    {
        //    //        if (FilterOwner.Orientation == Orientation.Vertical)
        //    //        {
        //    //            int headerHeight = Expander.OutputRange.Rows.Count;
        //    //            toShowHide = worksheet.Cells[RenderedArea.YFirstCell + headerHeight, 1];
        //    //            toShowHide = toShowHide.Resize[RenderedArea.Height - headerHeight, 1];
        //    //        }
        //    //        else
        //    //        {
        //    //            int headerWidth = Expander.OutputRange.Columns.Count;
        //    //            toShowHide = worksheet.Cells[1, RenderedArea.XFirstCell + headerWidth];
        //    //            toShowHide = toShowHide.Resize[1, RenderedArea.Width - headerWidth];
        //    //        }
        //    //        toShowHide.EntireRow.Hidden = IsExpanded;
        //    //        IsExpanded = !IsExpanded;
        //    //    }
        //    //    finally
        //    //    {
        //    //        ExcelApplication.ReleaseComObject(worksheet);
        //    //        worksheet = null;
        //    //        toShowHide = null;
        //    //    }
        //    //}
        //    //else
        //    //{
        //    //    IsExpanded = !IsExpanded;
        //    //    ITemplateView viewToRender = this;
        //    //    while (viewToRender.ParentElement != null)
        //    //    {
        //    //        viewToRender = viewToRender.ParentElement;
        //    //    }
        //    //    ETKExcel.TemplateManager.RenderView((IExcelTemplateView)viewToRender);
        //    //}
        //}

        internal void OnViewSheetIsActivated()
        {
            if (viewSheetIsActivated == null || IsDisposed || Renderer?.RenderedRange == null)
                return;

            try
            {
                viewSheetIsActivated();
            }
            catch (Exception ex)
            {
                string message = $"Sheet '{ViewSheet.Name}', Template '{TemplateDefinition.Name}'. 'ViewSheetIsActivated' failed: '{ex.Message}'";
                log.LogException(LogType.Error, ex, message);
            }
        }

        internal void OnViewSheetIsDeactivated()
        {
            if (ViewSheetIsDeactivated == null || IsDisposed || Renderer?.RenderedRange == null)
                return;

            try
            {
                ViewSheetIsDeactivated();
            }
            catch (Exception ex)
            {
                string message = $"Sheet '{ViewSheet.Name}', Template '{TemplateDefinition.Name}'. 'ViewSheetIsDeactivated' failed: '{ex.Message}'";
                log.LogException(LogType.Error, ex, message);
            }
        }

        /// <summary>
        /// Bind the template to Excel => Refresh Excel cells from the datasource currently injected. 
        /// </summary>
        internal void RenderView()
        {
            lock (syncRoot)
            {
                if (!IsDisposed && Renderer != null)
                {
                    if (ETKExcel.ExcelApplication.IsInEditMode())
                        throw new COMException("Excel is on Edit mode");

                    try
                    {
                        using (var freezeExcel = new FreezeExcel(ETKExcel.ExcelApplication.KeepStatusVisible))
                        {
                            BeforeRendering?.Invoke(false);

                            // Clear the previous rendering.
                            ////////////////////////////////
                            CurrentSelectedCell = null;
                            currentSelectedRange = null;
                            Renderer.Clear();

                            Renderer.Render();

                            ExecuteAutoFit();

                            if (log.GetLogLevel() == LogType.Debug)
                                log.LogFormat(LogType.Debug, "Sheet '{0}', View '{1}' from '{2}' rendered.", ViewSheet.Name, this.Ident, TemplateDefinition.Name);

                            AfterRendering?.Invoke(false);
                        }
                    }
                    catch (Exception ex)
                    {
                        string message = $"Sheet '{ViewSheet.Name}', View '{this.Ident}' from '{TemplateDefinition.Name}' render failed.";
                        throw new EtkException(message, ex, false);
                    }
                }
            }
        }

        internal void ManageExpander()
        {
            if(Renderer != null)
                ManageExpander(Renderer);
        }

        /// <summary>
        /// Bind the template to Excel => RenderView Excel cells based on the datasource currently injected. 
        /// </summary>
        internal void RenderViewDataOnly()
        {
            lock (syncRoot)
            {
                if (!IsDisposed && Renderer != null)
                {
                    if (ETKExcel.ExcelApplication.IsInEditMode())
                        throw new COMException("Excel is on Edit mode");

                    try
                    {
                        if (Renderer.RenderedRange == null)
                            RenderView();
                        else
                        {
                            using (var freezeExcel = new FreezeExcel(ETKExcel.ExcelApplication.KeepStatusVisible))
                            {
                                if (BindingContext != null && BindingContext.Body.ElementsToRender != null)
                                {
                                    BeforeRendering?.Invoke(true);

                                    Renderer.RenderDataOnly();
                                    if (log.GetLogLevel() == LogType.Debug)
                                        log.LogFormat(LogType.Debug, "Sheet '{0}', View '{1}' from '{2}' render data only failed.", ViewSheet.Name, this.Ident, TemplateDefinition.Name);

                                    AfterRendering?.Invoke(true);

                                    CurrentSelectedCell?.Select();
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        var message = $"Sheet '{ViewSheet.Name}', View '{this.Ident}' from '{TemplateDefinition.Name}' render data only failed.";
                        throw new EtkException(message, ex, false);
                    }
                }
            }
        }

        internal bool OnSheetChange(ExcelApplication excelApplication, ExcelInterop.Range target)
        {
            if (!IsDisposed && Renderer != null && Renderer.RenderedRange != null)
            {
                ExcelInterop.Range intersect = excelApplication.Application.Intersect(Renderer.RenderedRange, target);
                if (intersect != null)
                {
                    using (var freeze = new FreezeExcel(ETKExcel.ExcelApplication.KeepStatusVisible))
                    {
                        if (Renderer.OnDataChanged(intersect) && DataChanged != null)
                            DataChanged();// null, null);
                    }
                    intersect = null;
                    return true;
                }
            }
            return false;
        }

        internal IBindingContextItem GetConcernedContextItem(ExcelInterop.Range target)
        {
            IBindingContextItem ret = null;
            if (IsRendered)
                ret = Renderer.GetConcernedContextItem(target);
            return ret;
        }

        internal bool OnSelectionChange(ExcelInterop.Range target)
        {
            try
            {
                CurrentSelectedCell = null;
                UnhighlightSelection();

                if (IsRendered)
                {
                    ExcelInterop.Range intersect = ETKExcel.ExcelApplication.Application.Intersect(RenderedRange, target);
                    if (intersect != null)
                    {
                        CurrentSelectedCell = target.Cells[1, 1];

                        IBindingContextItem currentContextItem = GetConcernedContextItem(target);
                        if (currentContextItem?.BindingDefinition != null)
                        {
                            // If the binding excelBindingDefinition contains a selection callback: invoke it !
                            if (currentContextItem.BindingDefinition.OnSelection != null)
                            {
                                ((ExcelTemplateManager)ETKExcel.TemplateManager).CallbacksManager.Invoke(currentContextItem.BindingDefinition.OnSelection, 
                                                                                                         target, currentContextItem.ParentElement, currentContextItem);
                            }
                            else
                            {
                                // Ask the containing template (and its owner and the owner of its owner etc.... => bubble up the event)) if they contain a selection callback
                                // Invoke the first found 
                                IBindingContextElement catchingContextElement = currentContextItem.ParentElement;
                                bool isResolved = false;
                                do
                                {
                                    ExcelTemplateDefinitionPart currentTemplateDefinition = catchingContextElement.ParentPart.TemplateDefinitionPart as ExcelTemplateDefinitionPart;
                                    if (currentTemplateDefinition.PartType == TemplateDefinitionPartType.Body)
                                    {
                                        EventCallback callback = (currentTemplateDefinition.Parent as ExcelTemplateDefinition).SelectionChanged;
                                        if (callback != null)
                                        {
                                            ((ExcelTemplateManager)ETKExcel.TemplateManager).CallbacksManager.Invoke(callback, target, catchingContextElement, currentContextItem);
                                            isResolved = true;
                                        }
                                    }
                                    if (!isResolved)
                                        catchingContextElement = catchingContextElement.ParentPart.ParentContext?.Parent;
                                }
                                while (!isResolved && catchingContextElement != null);
                            }
                        }
                        intersect = null;
                        HighlightSelection(target);
                    }
                }
            }
            catch (Exception ex)
            {
                string message = $"Sheet '{target.Worksheet.Name}', Template '{TemplateDefinition.Name}' 'OnSelectionChange' failed: '{ex.Message}'";
                log.LogException(LogType.Error, ex, message);
            }
            return CurrentSelectedCell != null;
        }

        internal bool OnBeforeBoubleClick(ExcelInterop.Range target, ref bool cancel)
        {
            ExcelInterop.Range intersect = ETKExcel.ExcelApplication.Application.Intersect(RenderedRange, target);
            if (intersect == null)
                return false;

            IBindingContextItem currentContextItem = GetConcernedContextItem(target);
            if (currentContextItem?.BindingDefinition != null)
            {
                if (currentContextItem.BindingDefinition.IsReadOnly)
                    cancel = true;

                // If the bound excelBindingDefinition contains a left double click callback: invoke it !
                if (currentContextItem.BindingDefinition.OnClick != null)
                {
                    ((ExcelTemplateManager)ETKExcel.TemplateManager).CallbacksManager.Invoke(currentContextItem.BindingDefinition.OnClick, 
                                                                                             target, currentContextItem.ParentElement, currentContextItem);
                    cancel = true;
                }
                else
                {
                    IBindingContextElement currentContextElement = currentContextItem.ParentElement;
                    if (currentContextElement != null && currentContextElement.ParentPart != null && currentContextElement.ParentPart.PartType == BindingContextPartType.Header 
                        && ((TemplateDefinition)currentContextElement.ParentPart.TemplateDefinitionPart.Parent).TemplateOption.HeaderAsExpander != HeaderAsExpander.None)
                    {
                        if(CheckHeaderAsExpander(Renderer, target))
                            cancel = true;
                    }
                }
            }

            intersect = null;
            return true;
        }
        #endregion

        #region private methods
        private bool CheckHeaderAsExpander(ExcelRenderer renderer,  ExcelInterop.Range target)
        {
            if (renderer.HeaderPartRenderer?.RenderedRange != null && ETKExcel.ExcelApplication.Application.Intersect(renderer.HeaderPartRenderer.RenderedRange, target) != null)
            {
                renderer.IsExpanded = ! renderer.IsExpanded;
                ManageExpander(renderer);
                return true;
            }
            else
            {
                foreach (ExcelRenderer nestedRenderer in renderer.NestedRenderer)
                {
                    if (CheckHeaderAsExpander(nestedRenderer, target))
                        return true;
                }
            }
            return false;
        }

        private void ManageExpander(ExcelRenderer renderer)
        {
            using (var freezeExcel = new FreezeExcel(ETKExcel.ExcelApplication.KeepStatusVisible))
            {
                if(renderer.BodyPartRenderer != null && renderer.BodyPartRenderer.RenderedRange != null 
                   || renderer.FooterPartRenderer != null && renderer.FooterPartRenderer.RenderedRange != null)
                {
                    bool carryOn = true;
                    if(renderer.HeaderPartRenderer != null && renderer.HasExpander)
                    {
                        carryOn = renderer.IsExpanded;

                        int toShowHideSize = renderer.RenderedArea.Height - renderer.HeaderPartRenderer.RenderedArea.Height;
                        if (toShowHideSize > 0)
                        {
                            ExcelInterop.Range toShowHide = renderer.RenderedRange.Offset[renderer.HeaderPartRenderer.RenderedArea.Height, Type.Missing];
                            toShowHide = toShowHide.Resize[toShowHideSize, Type.Missing];
                            toShowHide.EntireRow.Hidden = !renderer.IsExpanded;

                            // ExcelApplication.ReleaseComObject(toShowHide);
                            toShowHide = null;
                        }
                    }

                    if (carryOn)
                    {
                        foreach (ExcelRenderer nestedRenderer in renderer.NestedRenderer)
                            ManageExpander(nestedRenderer);
                    }
                }
            }
        }

        private void HighlightSelection(ExcelInterop.Range selectedCell)
        {
            ExcelInterop.Range viewSelectedRange = null;
            ExcelInterop.Worksheet sheet = (ExcelInterop.Worksheet) RenderedRange.Parent;

            if (TemplateDefinition.Orientation == Orientation.Vertical)
            {
                viewSelectedRange = sheet.Cells[selectedCell.Row, RenderedRange.Column];
                viewSelectedRange = viewSelectedRange.Resize[1, RenderedRange.Columns.Count];

                currentSelectedRange = viewSelectedRange;
            }
            else
            {
                viewSelectedRange = sheet.Cells[RenderedRange.Row, selectedCell.Column];
                viewSelectedRange = viewSelectedRange.Resize[RenderedRange.Rows.Count, 1];

                currentSelectedRange = viewSelectedRange;
            }

            for (int i = 1; i <= currentSelectedRange.Cells.Count; i++)
            {
                ExcelInterop.Range cell = currentSelectedRange.Cells[1, i];
                if (CurrentSelectedCell.Column != cell.Column || CurrentSelectedCell.Row != cell.Row)
                {
                    ExcelInterop.Interior interior = cell.Interior;
                    try
                    {
                        if (interior.Gradient != null)
                            currentSelectedRangePattern.Add(null);
                        else
                        {
                            currentSelectedRangePattern.Add(new SelectionPattern(ref interior));
                            interior.Pattern = ExcelInterop.XlPattern.xlPatternGray8;
                            interior.PatternColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DimGray);
                        }
                    }
                    catch
                    { }
                    ExcelApplication.ReleaseComObject(interior);
                    interior = null;
                }
                else
                    currentSelectedRangePattern.Add(null);
            }

            // Redraw the borders of the current selection
            if (((TemplateDefinition) TemplateDefinition).AddBorder)
                Renderer.BorderAround(currentSelectedRange, ExcelInterop.XlLineStyle.xlContinuous, ExcelInterop.XlBorderWeight.xlThin, 1);

            viewSelectedRange = null;
            ExcelApplication.ReleaseComObject(sheet);
            sheet = null;
        }

        private void UnhighlightSelection()
        {
            // If not the first selection, redraw the borders of the previously selected range
            if (currentSelectedRange != null)
            {
                int cpt = 0;
                foreach (ExcelInterop.Range cell in currentSelectedRange.Cells)
                {
                    try
                    {
                        SelectionPattern selectionPattern = currentSelectedRangePattern[cpt++];
                        if (selectionPattern != null)
                        {
                            ExcelInterop.Interior interior = cell.Interior;

                            cell.Interior.Pattern = selectionPattern.Pattern;
                            if (selectionPattern.PatternColorIndex >= 0)
                                cell.Interior.PatternColorIndex = selectionPattern.PatternColorIndex;
                            if (selectionPattern.PatternColor != 0)
                                cell.Interior.PatternColor = selectionPattern.PatternColor;
                            if (selectionPattern.PatternThemeColor != 0)
                                cell.Interior.PatternThemeColor = selectionPattern.PatternThemeColor;
                            cell.Interior.PatternTintAndShade = selectionPattern.PatternTintAndShade;

                            ExcelApplication.ReleaseComObject(interior);
                            interior = null;
                        }
                    }
                    catch
                    { }
                }
                currentSelectedRangePattern.Clear();
                currentSelectedRange = null;
            }
        }
        #endregion
    }
}
