namespace Etk.Excel.BindingTemplates.Views
{
    using Etk.BindingTemplates.Context;
    using Etk.BindingTemplates.Definitions.EventCallBacks;
    using Etk.BindingTemplates.Definitions.Templates;
    using Etk.BindingTemplates.Views;
    using Etk.Excel.Application;
    using Etk.Excel.BindingTemplates.Definitions;
    using Etk.Excel.BindingTemplates.Renderer;
    using Etk.Excel.UI.Log;
    using Microsoft.Office.Interop.Excel;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;

    struct SelectionPatternColor
    {
        public int ThemeColor;
        public object TintAndShade;
        public object Color;
        public double Position;

        public SelectionPatternColor(ColorStop cs)
        {
            ThemeColor = cs.ThemeColor;
            TintAndShade = cs.TintAndShade;
            Color = cs.Color;
            Position = cs.Position;
        }
    }

    class SelectionPattern
    {
        public XlPattern Pattern;
        public int PatternColor;
        public int PatternColorIndex;
        public int PatternThemeColor;
        public double PatternTintAndShade;

        public object GradientDegree;
        public SelectionPatternColor[] SelectionPatternColors;

        public SelectionPattern(ref Interior interior)
        {
            try
            {
                Pattern = (XlPattern)interior.Pattern;
                PatternColor = interior.PatternColor;
                PatternColorIndex = interior.PatternColorIndex;
                PatternThemeColor = interior.PatternThemeColor;
                PatternTintAndShade = interior.PatternTintAndShade;

                if (interior.Gradient != null)
                {
                    GradientDegree = interior.Gradient.Degree;
                    SelectionPatternColors = new SelectionPatternColor[interior.Gradient.ColorStops.Count];
                    for (int cpt = 0; cpt < interior.Gradient.ColorStops.Count; cpt++)
                    {
                        ColorStop cs = interior.Gradient.ColorStops[cpt + 1];
                        SelectionPatternColors[cpt] = new SelectionPatternColor(cs);
                    }
                }
            }
            catch
            { }
        }
    }

    class ExcelTemplateView : TemplateView, IExcelTemplateView
    {
        #region attributes and properties
        private ILogger log = Logger.Instance;
        private Range currentSelectedRange;
        private List<SelectionPattern> currentSelectedRangePattern = new List<SelectionPattern>();

        internal Range CurrentSelectedCell
        { get; private set; }

        public event Action<object, object> DataChanged;
        public event Action<bool> BeforeRendering;
        public event Action<bool> AfterRendering;
        public event Action<IExcelTemplateView> SheetActivation;

        public bool AutoFit
        { get; set; }

        public Worksheet SheetDestination
        { get; private set; }

        public Range FirstOutputCell
        { get; set; }

        public Range ClearingCell
        { get; set; }

        public ExcelRootRenderer Renderer
        { get; private set; }

        public bool IsRendered
        { get { return Renderer != null && Renderer.RenderedRange != null; } }

        public Range RenderedRange
        { get { return Renderer != null ? Renderer.RenderedRange : null; } }

        public RenderedArea RenderedArea
        { get { return Renderer != null ? Renderer.RenderedArea : null; } }

        ////public RenderingArea RenderedArea
        ////{ get { return renderer == null ? null : RenderingArea.CreateInstances(renderer.RenderedRange); } }

        public AccessorParametersManager AccessorParametersManager
        { get; private set; }

        public ExcelPartRenderer Expander
        { get; set; }

        private bool isExpanded = true;
        public bool IsExpanded
        {
            get
            {
                //if ((FilterOwner as ExcelTemplateDefinition).ExpanderBindingDefinition != null)
                //    isExpanded = (bool)(FilterOwner as ExcelTemplateDefinition).ExpanderBindingDefinition.ResolveBinding(GetDataSource());
                return isExpanded;
            }
            set
            {
                //&&if ((FilterOwner as ExcelTemplateDefinition).ExpanderBindingDefinition != null)
                //    isExpanded = (bool)(FilterOwner as ExcelTemplateDefinition).ExpanderBindingDefinition.UpdateDataSource(GetDataSource(), !isExpanded);
                //else
                    isExpanded = !isExpanded;
            }
        }

        public List<Range> CellsThatContainSearchValue
        { get; set; }
        #endregion

        #region .ctors
        public ExcelTemplateView(ITemplateDefinition templateDefinition, Worksheet sheetDestination, Range firstOutputCell, Range clearingCell)
            : base(templateDefinition)
        {
            SheetDestination = sheetDestination;
            FirstOutputCell = firstOutputCell;
            ClearingCell = clearingCell;
            AutoFit = true;
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

                if (CellsThatContainSearchValue != null)
                    CellsThatContainSearchValue.Clear();

                base.Clear();
                if (!IsDisposed && Renderer != null)
                {
                    if (ETKExcel.ExcelApplication.IsInEditMode())
                        throw new COMException("Excel is on Edit mode");
                    Renderer.Clear();
                    if (log.GetLogLevel() == LogType.Debug)
                        log.LogFormat(LogType.Debug, "Sheet '{0}', View '{1}' from '{2}' cleared.", SheetDestination.Name, this.Ident, TemplateDefinition.Name);
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
                        string message = string.Format("Sheet '{0}', View '{1}' from '{2}' Set data source failed.", SheetDestination.Name, this.Ident, TemplateDefinition.Name);
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
                    Expander = null;
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

                    SheetDestination = null;
                    FirstOutputCell = null;
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
        #endregion

        #region internal methods
        internal void ResolveExpander()
        {
            //if (!FilterOwner.HeaderAsExpander)
            //    return;

            //if (FilterOwner.ExpanderMode == ExpanderMode.Hide)
            //{
            //    Worksheet worksheet = Expander.OutputRange.Worksheet;
            //    Range toShowHide;
            //    try
            //    {
            //        if (FilterOwner.Orientation == Orientation.Vertical)
            //        {
            //            int headerHeight = Expander.OutputRange.Rows.Count;
            //            toShowHide = worksheet.Cells[RenderedArea.YFirstCell + headerHeight, 1];
            //            toShowHide = toShowHide.Resize[RenderedArea.Height - headerHeight, 1];
            //        }
            //        else
            //        {
            //            int headerWidth = Expander.OutputRange.Columns.Count;
            //            toShowHide = worksheet.Cells[1, RenderedArea.XFirstCell + headerWidth];
            //            toShowHide = toShowHide.Resize[1, RenderedArea.Width - headerWidth];
            //        }
            //        toShowHide.EntireRow.Hidden = IsExpanded;
            //        IsExpanded = !IsExpanded;
            //    }
            //    finally
            //    {
            //        Marshal.ReleaseComObject(worksheet);
            //        worksheet = null;
            //        toShowHide = null;
            //    }
            //}
            //else
            //{
            //    IsExpanded = !IsExpanded;
            //    ITemplateView viewToRender = this;
            //    while (viewToRender.ParentElement != null)
            //    {
            //        viewToRender = viewToRender.ParentElement;
            //    }
            //    ETKExcel.TemplateManager.Render((IExcelTemplateView)viewToRender);
            //}
        }

        internal void OnSheetActivation()
        {
            if (SheetActivation != null)
                SheetActivation(this);
        }

        /// <summary>
        /// Bind the template to Excel => Refresh Excel cells from the datasource currently injected. 
        /// </summary>
        internal void Render()
        {
            lock (syncRoot)
            {
                if (!IsDisposed && Renderer != null)
                {
                    if (ETKExcel.ExcelApplication.IsInEditMode())
                        throw new COMException("Excel is on Edit mode");

                    try
                    {
                        using (FreezeExcel freezeExcel = new FreezeExcel())
                        {
                            if (BeforeRendering != null)
                                BeforeRendering(false);

                            // Clear the previous rendering.
                            ////////////////////////////////                            
                            Renderer.Clear();

                            Renderer.Render();
                            ExecuteAutoFit();

                            if (log.GetLogLevel() == LogType.Debug)
                                log.LogFormat(LogType.Debug, "Sheet '{0}', View '{1}' from '{2}' rendered.", SheetDestination.Name, this.Ident, TemplateDefinition.Name);

                            if (AfterRendering != null)
                                AfterRendering(false);
                        }
                    }
                    catch (Exception ex)
                    {
                        string message = string.Format("Sheet '{0}', View '{1}' from '{2}' render failed.", SheetDestination.Name, this.Ident, TemplateDefinition.Name);
                        throw new EtkException(message, ex, false);
                    }
                }
            }
        }

        /// <summary>
        /// Bind the template to Excel => Render Excel cells based on the datasource currently injected. 
        /// </summary>
        internal void RenderDataOnly()
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
                            Render();
                        else
                        {
                            using (FreezeExcel freezeExcel = new FreezeExcel())
                            {
                                if (BindingContext != null && BindingContext.Body.ElementsToRender != null)
                                {
                                    if (BeforeRendering != null)
                                        BeforeRendering(true);

                                    Renderer.RenderDataOnly();
                                    if (log.GetLogLevel() == LogType.Debug)
                                        log.LogFormat(LogType.Debug, "Sheet '{0}', View '{1}' from '{2}' render data only failed.", SheetDestination.Name, this.Ident, TemplateDefinition.Name);
                                    if (AutoFit)
                                        ExecuteAutoFit();

                                    if (AfterRendering != null)
                                        AfterRendering(true);

                                    if (CurrentSelectedCell != null)
                                        CurrentSelectedCell.Select();
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        string message = string.Format("Sheet '{0}', View '{1}' from '{2}' render data only failed.", SheetDestination.Name, this.Ident, TemplateDefinition.Name);
                        throw new EtkException(message, ex, false);
                    }
                }
            }
        }

        internal bool OnSheetChange(ExcelApplication excelApplication, Range target)
        {
            if (!IsDisposed && Renderer != null && Renderer.RenderedRange != null)
            {
                Range intersect = excelApplication.Application.Intersect(Renderer.RenderedRange, target);
                if (intersect != null)
                {
                    using (FreezeExcel freeze = new FreezeExcel())
                    {
                        if (Renderer.OnDataChanged(intersect) && DataChanged != null)
                            DataChanged(null, null);
                        intersect = null;
                    }
                    return true;
                }
            }
            return false;
        }

        internal IBindingContextItem GetConcernedContextItem(Range target)
        {
            IBindingContextItem ret = null;
            if (IsRendered)
                ret = Renderer.GetConcernedContextItem(target);
            return ret;
        }

        internal bool OnSelectionChange(ExcelApplication excelApplication, Range realTarget)
        {
            try
            {
                CurrentSelectedCell = null;
                UnhighlightSelection();

                if (IsRendered)
                {
                    Range intersect = excelApplication.Application.Intersect(RenderedRange, realTarget);
                    if (intersect != null)
                    {
                        CurrentSelectedCell = realTarget.Cells[1, 1];

                        IBindingContextItem contextItem = GetConcernedContextItem(realTarget);
                        if (contextItem != null && contextItem.BindingDefinition != null)
                        {
                            // If the binding excelTemplateDefinition contains a selection callback: invoke it !
                            if (contextItem.BindingDefinition.OnSelection != null)
                                contextItem.BindingDefinition.OnSelection.Invoke(realTarget, contextItem.ParentElement, contextItem.ParentElement);
                            else
                            {
                                // Ask the containing template (and its owner and the owner of its owner etc.... => bubble up the event)) if they contain a selection callback
                                // Invoke the first found 
                                IBindingContextElement respondingContextElement = contextItem.ParentElement;
                                IBindingContextElement selectedContextElement = respondingContextElement;
                                bool isResolved = false;
                                do
                                {
                                    ExcelTemplateDefinitionPart currentTemplateDefinition = respondingContextElement.ParentPart.TemplateDefinitionPart as ExcelTemplateDefinitionPart;
                                    EventCallback callback = (currentTemplateDefinition.Parent as ExcelTemplateDefinition).SelectionChanged;
                                    if (callback != null)
                                    {
                                        callback.Invoke(realTarget, respondingContextElement, selectedContextElement);
                                        isResolved = true;
                                    }
                                    if (!isResolved)
                                        respondingContextElement = respondingContextElement.ParentPart.ParentContext == null ? null : respondingContextElement.ParentPart.ParentContext.Parent;
                                }
                                while (!isResolved && respondingContextElement != null);
                            }
                        }
                        intersect = null;

                        HighlightSelection(realTarget);
                    }
                }
            }
            catch (Exception ex)
            {
                string message = string.Format("Sheet '{0}', Template '{1}' 'OnSelectionChange' failed: '{2}'", realTarget.Worksheet.Name, TemplateDefinition.Name, ex.Message);
                log.LogException(LogType.Error, ex, message);
            }
            return CurrentSelectedCell != null;
        }

        internal bool OnBeforeBoubleClick(Range target, ref bool cancel)
        {
            bool ret = false; ;
            Range intersect = ETKExcel.ExcelApplication.Application.Intersect(RenderedRange, target);
            if (intersect != null)
            {
                IBindingContextItem contextItem = GetConcernedContextItem(target);
                if (contextItem != null && contextItem.BindingDefinition != null)
                {
                    if (contextItem.BindingDefinition.IsReadOnly)
                        cancel = true;

                    // If the binding excelTemplateDefinition contains a left double click callback: invoke it !
                    if (contextItem.BindingDefinition.OnClick != null)
                        contextItem.BindingDefinition.OnClick.Invoke(target, contextItem.ParentElement, contextItem.ParentElement);
                    else
                    {
                        // If not, ask the containing template (and its owner and the owner of its owner etc.... => bubble up the event)) if they contain a left double click callback
                        // Invoke the first found 
                        IBindingContextElement respondingContextElement = contextItem.ParentElement;
                        IBindingContextElement selectedContextElement = respondingContextElement;
                        do
                        {
                            ExcelTemplateDefinitionPart currentTemplateDefinition = respondingContextElement.ParentPart.TemplateDefinitionPart as ExcelTemplateDefinitionPart;
                            EventCallback callback = (currentTemplateDefinition.Parent as ExcelTemplateDefinition).OnLeftDoubleClick;
                            if (callback != null)
                            {
                                callback.Invoke(target, respondingContextElement, selectedContextElement);
                                ret = true;
                            }
                            if (!ret)
                                respondingContextElement = respondingContextElement.ParentPart.ParentContext == null ? null : respondingContextElement.ParentPart.ParentContext.Parent;
                        }
                        while (!ret && respondingContextElement != null);
                    }
                }
            }

            // Manage the expander (=> the header capability to expand)
            if (!ret && Expander != null && Expander.RenderedArea != null)
            {
                //Range intersectExpander = ETKExcel.ExcelApplication.Application.Intersect(Expander.RenderedRange, target);
                //if (intersectExpander != null)
                //{
                //    ResolveExpander();
                //    ret = true;
                //}
            }
            intersect = null;

            if (ret)
                cancel = true;
            return ret;
        }
        #endregion

        #region private methods
        private void HighlightSelection(Range selectedCell)
        {
            Range viewSelectedRange = null;
            Worksheet sheet = RenderedRange.Parent as Worksheet;

            if (this.TemplateDefinition.Orientation == Orientation.Vertical)
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
                Range cell = currentSelectedRange.Cells[1, i];
                if (CurrentSelectedCell.Column != cell.Column || CurrentSelectedCell.Row != cell.Row)
                {
                    Interior interior = cell.Interior;
                    try
                    {
                        currentSelectedRangePattern.Add(new SelectionPattern(ref interior));

                        SelectionPattern selectionPattern = null;
                        if (interior.Gradient != null)
                            selectionPattern = new SelectionPattern(ref interior);

                        interior.Pattern = XlPattern.xlPatternGray8;
                        if (selectionPattern == null)
                            interior.PatternColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DimGray);
                        else
                        {
                            if (selectionPattern.SelectionPatternColors != null)
                            {
                                interior.Color = selectionPattern.SelectionPatternColors[0].Color;
                                if (selectionPattern.SelectionPatternColors.Count() > 1)
                                    interior.PatternColor = selectionPattern.SelectionPatternColors[1].Color;
                            }
                        }
                    }
                    catch
                    { }
                    Marshal.ReleaseComObject(interior);
                }
                else
                    currentSelectedRangePattern.Add(null);
            }

            // Redraw the borders of the current selection

            if (((TemplateDefinition) TemplateDefinition).AddBorder)
                Renderer.BorderAround(currentSelectedRange, XlLineStyle.xlContinuous, XlBorderWeight.xlThin, 3);
            Marshal.ReleaseComObject(sheet);
        }

        private void UnhighlightSelection()
        {
            // If not the first selection, redraw the borders of the previously selected range
            if (currentSelectedRange != null)
            {
                int cpt = 0;
                foreach (Range cell in currentSelectedRange.Cells)
                {
                    try
                    {
                        SelectionPattern selectionPattern = currentSelectedRangePattern[cpt++];
                        if (selectionPattern != null)
                        {
                            Interior interior = cell.Interior;

                            cell.Interior.Pattern = selectionPattern.Pattern;
                            if (selectionPattern.PatternColor != 0)
                                cell.Interior.PatternColor = selectionPattern.PatternColor;
                            if (selectionPattern.PatternColorIndex >= 0)
                                cell.Interior.PatternColorIndex = selectionPattern.PatternColorIndex;
                            if (selectionPattern.PatternThemeColor != 0)
                                cell.Interior.PatternThemeColor = selectionPattern.PatternThemeColor;
                            cell.Interior.PatternTintAndShade = selectionPattern.PatternTintAndShade;

                            if (cell.Interior.Gradient != null)
                            {
                                if (selectionPattern.GradientDegree != null)
                                 cell.Interior.Gradient.Degree = selectionPattern.GradientDegree;
                                cell.Interior.Gradient.ColorStops.Clear();

                                if (selectionPattern.SelectionPatternColors != null)
                                {
                                    int i = 0;
                                    foreach (SelectionPatternColor selectionPatternColor in selectionPattern.SelectionPatternColors)
                                    {
                                        ColorStop cs = cell.Interior.Gradient.ColorStops.Add(i++);
                                        if (selectionPatternColor.ThemeColor != 0)
                                            cs.ThemeColor = selectionPatternColor.ThemeColor;
                                        cs.TintAndShade = selectionPatternColor.TintAndShade;
                                        cs.Color = selectionPatternColor.Color;
                                        cs.Position = selectionPatternColor.Position;
                                    }
                                }
                            }
                            Marshal.ReleaseComObject(interior);
                        }
                    }
                    catch
                    { }
                }
                currentSelectedRangePattern.Clear();
                currentSelectedRange = null;
            }
        }

        private void ExecuteAutoFit()
        {
            if (AutoFit && Renderer.RenderedRange != null)
            {
                if (TemplateDefinition.Orientation == Orientation.Horizontal)
                    Renderer.RenderedRange.Rows.AutoFit();
                else
                    Renderer.RenderedRange.Columns.AutoFit();
            }
        }
        #endregion
    }
}
