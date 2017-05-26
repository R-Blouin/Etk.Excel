using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Context.SortSearchAndFilter;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.Excel.Application;
using Etk.Excel.BindingTemplates.Controls.FormulaResult;
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
        #endregion

        #region .ctors 
        public ExcelRootRenderer(ExcelTemplateView view) : base(null, view.TemplateDefinition, view.BindingContext, view.FirstOutputCell, null)
        {
            this.View = view;
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
            RenderDataOnly();
        }

        public void RenderDataOnly()
        {
            if (IsDisposed || RenderedRange == null)
                return;

            contextItems = new IBindingContextItem[RenderedArea.Height, RenderedArea.Width];
            cells = new object[RenderedArea.Height, RenderedArea.Width];
            ConcurrentStack<KeyValuePair<IBindingContextItem, System.Drawing.Point>> decorators = new ConcurrentStack<KeyValuePair<IBindingContextItem, System.Drawing.Point>>();

            Parallel.For(0, DataRows.Count, i =>
            {
                List<IBindingContextItem> itemsInRow = DataRows[i];
                int colId = 0;
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
                        cells[i, colId++] = value != null && value is Enum ? ((Enum) value).ToString() : value;
                    }
                    else
                        cells[i, colId++] = null;
                }
            });
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
            if (((TemplateDefinition) this.View.TemplateDefinition).AddBorder)
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
                        if (this.View.TemplateDefinition.Orientation == Orientation.Horizontal)
                            RenderedRange.EntireColumn.Hidden = false;
                        else
                            RenderedRange.EntireRow.Hidden = false;

                        if (View.ClearingCell != null)
                            View.ClearingCell.Copy(RenderedRange);

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
                        bool update = contextItem.UpdateDataSource(cell.Value2, out retValue);
                        if (update)
                        {
                            if (!object.Equals(cell.Value2, retValue))
                                cell.Value2 = retValue;
                            if (!(contextItem is BindingFilterContextItem))
                                ret = true;
                        }
                    }
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
            if (!IsDisposed && DataRows != null)
            {
                IEnumerable<IBindingContextItem> items = DataRows.SelectMany(r => r.Where(ci => ci is ISheetCalculate));
                foreach (IBindingContextItem item in items)
                    ((ISheetCalculate) item).OnSheetCalculate();
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
            Marshal.ReleaseComObject(borders);
            borders = null;
        }
        #endregion
    }
}
