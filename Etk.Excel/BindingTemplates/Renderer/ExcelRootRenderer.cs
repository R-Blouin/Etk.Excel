using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Context.SortSearchAndFilter;
using Etk.BindingTemplates.Definitions.EventCallBacks;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.Excel.Application;
using Etk.Excel.BindingTemplates.Controls.WithFormula;
using Etk.Excel.BindingTemplates.Decorators;
using Etk.Excel.BindingTemplates.Views;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.BindingTemplates.Renderer
{
    class ExcelRootRenderer : ExcelRenderer
    {
        #region properties
        //@@ private ExcelSortAndFilterButton sortAndFilterButton;

        public bool IsClearing
        { get; private set; }

        public  ExcelTemplateView View
        { get; private set; }

        public List<ExcelElementDecorator> RowDecorators
        { get; private set; }

        public IEnumerable<IFormulaCalculation> toOperateOnSheetCalculation;

        public List<SpecificEventCallback> AfterRenderingActions;
        #endregion

        #region .ctors 
        public ExcelRootRenderer(ExcelTemplateView view) : base(null, view.TemplateDefinition, view.BindingContext, view.FirstOutputCell, null)
        {
            View = view;
            RowDecorators = new List<ExcelElementDecorator>();
            //@@ sortAndFilterButton = new ExcelSortAndFilterButton(View);
        }
        #endregion

        #region public methods
        public override void Render()
        {
            if (IsDisposed)
                return;

            base.Render();

            toOperateOnSheetCalculation = ContextItems?.SelectMany(r => r.Where(ci => ci is IFormulaCalculation))
                                                   .Select(c =>(IFormulaCalculation) c).ToArray();
            RenderDataOnly();
        }

        public void RenderDataOnly()
        {
            if (IsDisposed || RenderedRange == null)
                return;

            contextItems = new IBindingContextItem[RenderedArea.Height, RenderedArea.Width];
            cells = new object[RenderedArea.Height, RenderedArea.Width];
            ConcurrentStack<KeyValuePair<IBindingContextItem, System.Drawing.Point>> decorators = new ConcurrentStack<KeyValuePair<IBindingContextItem, System.Drawing.Point>>();

            //Parallel.For(0, DataRows.Count, i => // Parrallel problem with Com object
            for (int i = 0; i <ContextItems.Count; i++)
            {
                int colId = 0;
                List<IBindingContextItem> itemsInRow = ContextItems[i];
                foreach (IBindingContextItem item in itemsInRow)
                {
                    if (item != null)
                    {
                        if (item.BindingDefinition != null && item.BindingDefinition.DecoratorDefinition != null)
                            decorators.Push(new KeyValuePair<IBindingContextItem, System.Drawing.Point>(item, new System.Drawing.Point(colId + 1, i + 1)));

                        contextItems[i, colId] = item;
                        if (item.CanNotify)
                        {
                            ((IBindingContextItemCanNotify)item).OnPropertyChangedAction = OnNotifyPropertyChanged;
                            ((IBindingContextItemCanNotify)item).OnPropertyChangedActionArgs = new KeyValuePair<int, int>(i, colId);
                        }
                        object value = item.ResolveBinding();
                        cells[i, colId++] = (value as Enum)?.ToString() ?? value;
                    }
                    else
                        cells[i, colId++] = null;
                }
            }
            //);
            RenderedRange.Value2 = cells;

            // Element decorators managements
            foreach(ExcelElementDecorator rowDecorator in RowDecorators)
                rowDecorator.Resolve();

            // Decorators managements
            foreach (KeyValuePair<IBindingContextItem, System.Drawing.Point> kvp in decorators)
            {
                ExcelInterop.Range range = RenderedRange[kvp.Value.Y, kvp.Value.X];
                kvp.Key.BindingDefinition.DecoratorDefinition.Resolve(range, kvp.Key);
                range = null;
            }

            // Redraw the borders of the current selection
            if (((TemplateDefinition) View.TemplateDefinition).AddBorder)
                BorderAround(RenderedRange, ExcelInterop.XlLineStyle.xlContinuous, ExcelInterop.XlBorderWeight.xlMedium, 1);
        }

        public void Clear()
        {
            if (!IsDisposed && RenderedRange != null)
            {
                using (FreezeExcel freezeExcel = new FreezeExcel(ETKExcel.ExcelApplication.KeepStatusVisible))
                {
                    try
                    {
                        IsClearing = true;

                        RenderedRange.Clear();
                        if (View.TemplateDefinition.Orientation == Orientation.Horizontal)
                            RenderedRange.EntireColumn.Hidden = false;
                        else
                            RenderedRange.EntireRow.Hidden = false;

                        View.ClearingCell?.Copy(RenderedRange);

                        RowDecorators.Clear();
                        ClearRenderingData();
                    }
                    finally
                    {
                        IsClearing = false;
                    }
                }
            }
        }

        public bool OnDataChanged(ExcelInterop.Range target)
        {
            bool ret = false;
            if (!IsDisposed && !IsClearing && contextItems != null)
            {
                FreezeExcel freezeExcel = null;
                try
                {
                    foreach (ExcelInterop.Range cell in target.Cells)
                    {
                        IBindingContextItem contextItem = null;
                        // Because of the merge cells ...
                        try
                        { contextItem = contextItems[cell.Row - View.FirstOutputCell.Row, cell.Column - View.FirstOutputCell.Column]; }
                        catch
                        { }

                        if (contextItem != null)
                        {
                            object retValue;
                            bool mustUpdate = contextItem.UpdateDataSource(cell.Value2, out retValue);
                            if (mustUpdate)
                            {
                                if (freezeExcel == null)
                                    freezeExcel = new FreezeExcel(ETKExcel.ExcelApplication.KeepStatusVisible);

                                //if (!object.Equals(cell.Value2, retValue))
                                cell.Value2 = retValue;
                            }

                            if (! (contextItem is BindingFilterContextItem))
                                ret = true;
                        }
                    }
                }
                finally
                {
                    if(freezeExcel != null)
                        freezeExcel.Dispose();
                }
            }
            return ret;
        }

        public IBindingContextItem GetConcernedContextItem(ExcelInterop.Range target)
        {
            IBindingContextItem ret = null;
            if (!IsDisposed && !IsClearing)
            {
                if (contextItems != null)
                    ret = contextItems[target.Row - View.FirstOutputCell.Row, target.Column - View.FirstOutputCell.Column];
            }
            return ret;
        }

        public void OnCalculate()
        {
            if (!IsDisposed && toOperateOnSheetCalculation != null)
            {
                foreach (IFormulaCalculation item in toOperateOnSheetCalculation)
                    item.OnSheetCalculate();
            }
        }

        public override void AddAfterRenderingAction(SpecificEventCallback callBack)
        {
            if(AfterRenderingActions == null)
                AfterRenderingActions = new List<SpecificEventCallback>();
            AfterRenderingActions.Add(callBack);
        }

        public void AfterRendering()
        {
            if (AfterRenderingActions != null)
            {
                foreach(SpecificEventCallback callback in AfterRenderingActions)
                    callback.Invoke();
            }
        }

        public new void Dispose()
        {
            if (!IsDisposed)
            {
                Clear();
                //@@if (sortAndFilterButton != null)
                //@@    sortAndFilterButton.Dispose();
                IsDisposed = true;
            }
        }
        #endregion

        #region internal, private and protected methods
        private void OnNotifyPropertyChanged(IBindingContextItem contextItem, object param)
        {
            if (!IsDisposed && !IsClearing)
                ((ExcelTemplateManager) ETKExcel.TemplateManager).ExcelNotifyPropertyManager.NotifyPropertyChanged(new ExcelNotityPropertyContext(contextItem, View, (KeyValuePair<int, int>)param));
        }

        internal void BorderAround(ExcelInterop.Range range, ExcelInterop.XlLineStyle lineStyle, ExcelInterop.XlBorderWeight Weight, int colorIndex)
        {
            ExcelInterop.Borders borders = range.Borders;
            borders[ExcelInterop.XlBordersIndex.xlEdgeLeft].ColorIndex = colorIndex;
            borders[ExcelInterop.XlBordersIndex.xlEdgeLeft].LineStyle = lineStyle;
            borders[ExcelInterop.XlBordersIndex.xlEdgeLeft].Weight = Weight;

            borders[ExcelInterop.XlBordersIndex.xlEdgeTop].ColorIndex = colorIndex;
            borders[ExcelInterop.XlBordersIndex.xlEdgeTop].LineStyle = lineStyle;
            borders[ExcelInterop.XlBordersIndex.xlEdgeTop].Weight = Weight;

            borders[ExcelInterop.XlBordersIndex.xlEdgeBottom].ColorIndex = colorIndex;
            borders[ExcelInterop.XlBordersIndex.xlEdgeBottom].LineStyle = lineStyle;
            borders[ExcelInterop.XlBordersIndex.xlEdgeBottom].Weight = Weight;

            borders[ExcelInterop.XlBordersIndex.xlEdgeRight].ColorIndex = colorIndex;
            borders[ExcelInterop.XlBordersIndex.xlEdgeRight].LineStyle = lineStyle;
            borders[ExcelInterop.XlBordersIndex.xlEdgeRight].Weight = Weight;

            ////borders.Color = color;
            ExcelApplication.ReleaseComObject(borders);
            borders = null;
        }
        #endregion
    }
}
