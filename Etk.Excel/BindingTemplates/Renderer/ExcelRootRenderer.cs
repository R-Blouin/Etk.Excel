namespace Etk.Excel.BindingTemplates.Renderer
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using System.Threading.Tasks;
    using Etk.BindingTemplates.Context;
    using Etk.BindingTemplates.Context.SortSearchAndFilter;
    using Etk.Excel.Application;
    using Etk.Excel.BindingTemplates.Views;
    using Microsoft.Office.Interop.Excel;
    using Etk.BindingTemplates.Definitions.Templates;

    class ExcelRootRenderer : ExcelRenderer
    {
        #region properties
        //@@ private ExcelSortAndFilterButton sortAndFilterButton;

        public bool IsClearing
        { get; private set; }

        public  ExcelTemplateView View
        { get; private set; }

        public bool IsDisposed
        { get; private set; }
        #endregion

        #region .ctors 
        public ExcelRootRenderer(ExcelTemplateView view) 
               : base(view.TemplateDefinition, view.BindingContext, view.FirstOutputCell)
        {
            this.View = view;
            //@@ sortAndFilterButton = new ExcelSortAndFilterButton(View);
        }
        #endregion

        #region public methods
        override public void Render()
        {
            base.Render();
            RenderDataOnly();
        }

        public void RenderDataOnly()
        {
            if (RenderedArea == null)
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
                        cells[i, colId++] = item.ResolveBinding();
                    }
                    else
                        cells[i, colId++] = null;
                }
            });
            RenderedRange.Value2 = cells;

            // Decorators managements
            foreach (KeyValuePair<IBindingContextItem, System.Drawing.Point> kvp in decorators)
            {
                Range range = RenderedRange[kvp.Value.Y, kvp.Value.X];
                kvp.Key.BindingDefinition.DecoratorDefinition.Resolve(range, kvp.Key);
                range = null;
            }

            // Redraw the borders of the current selection
            if (((TemplateDefinition) this.View.TemplateDefinition).AddBorder)
                BorderAround(RenderedRange, XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, 1);
        }

        public void Clear()
        {
            if (!IsDisposed && RenderedRange != null)
            {
                using (FreezeExcel freezeExcel = new FreezeExcel())
                {
                    try
                    {
                        IsClearing = true;
                        if (View.ClearingCell != null)
                            View.ClearingCell.Copy(RenderedRange);
                        else
                            RenderedRange.Clear();
                        RenderedRange = null;

                        if (HeaderPartRenderer != null)
                        {
                            HeaderPartRenderer.Dispose();
                            HeaderPartRenderer = null;
                        }
                        if (BodyPartRenderer != null)
                        {
                            BodyPartRenderer.Dispose();
                            BodyPartRenderer = null;
                        }
                        if (FooterPartRenderer != null)
                        {
                            FooterPartRenderer.Dispose();
                            FooterPartRenderer = null;
                        }
                    }
                    finally
                    {
                        IsClearing = false;
                    }
                }
            }
        }

        public bool OnDataChanged(Range target)
        {
            bool ret = false;
            foreach (Range cell in target.Cells)
            {
                IBindingContextItem contextItem = null;
                // Be&cause of the merge cells ...
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
                        if (! (contextItem is BindingFilterContextItem))
                            ret = true;
                    }
                }
            }
            return ret;
        }

        public IBindingContextItem GetConcernedContextItem(Range target)
        {
            IBindingContextItem ret = null;
            if (!IsDisposed && !IsClearing)
            {
                if (contextItems != null)
                    ret = contextItems[target.Row - View.FirstOutputCell.Row, target.Column - View.FirstOutputCell.Column];
            }
            return ret;
        }

        new public void Dispose()
        {
            if (!IsDisposed)
            {
                Clear();
                base.Dispose();
                //@@if (sortAndFilterButton != null)
                //@@    sortAndFilterButton.Dispose();
                IsDisposed = true;
            }
        }
        #endregion

        #region private methods
        private void OnNotifyPropertyChanged(IBindingContextItem contextItem, object param)
        {
            if (! IsDisposed && ! IsClearing)
                ((ExcelTemplateManager)ETKExcel.TemplateManager).ExcelNotifyPropertyManager.NotifyPropertyChanged(new ExcelNotityPropertyContext(contextItem, View, (KeyValuePair<int, int>)param));
        }

        internal void BorderAround(Range range, XlLineStyle lineStyle, XlBorderWeight Weight, int colorIndex)
        {
            Borders borders = range.Borders;
            borders[XlBordersIndex.xlEdgeLeft].ColorIndex = colorIndex;
            borders[XlBordersIndex.xlEdgeLeft].LineStyle = lineStyle;
            borders[XlBordersIndex.xlEdgeLeft].Weight = Weight;

            borders[XlBordersIndex.xlEdgeTop].ColorIndex = colorIndex;
            borders[XlBordersIndex.xlEdgeTop].LineStyle = lineStyle;
            borders[XlBordersIndex.xlEdgeTop].Weight = Weight;

            borders[XlBordersIndex.xlEdgeBottom].ColorIndex = colorIndex;
            borders[XlBordersIndex.xlEdgeBottom].LineStyle = lineStyle;
            borders[XlBordersIndex.xlEdgeBottom].Weight = Weight;

            borders[XlBordersIndex.xlEdgeRight].ColorIndex = colorIndex;
            borders[XlBordersIndex.xlEdgeRight].LineStyle = lineStyle;
            borders[XlBordersIndex.xlEdgeRight].Weight = Weight;

            ////borders.Color = color;
            Marshal.ReleaseComObject(borders);
            borders = null;
        }
        #endregion
    }
}
