using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.EventCallBacks;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.Excel.BindingTemplates.Definitions;
using Etk.Excel.BindingTemplates.Views;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.BindingTemplates.Renderer
{
    class ExcelRenderer : IDisposable
    {
        protected ITemplateDefinition templateDefinition;
        protected IBindingContext bindingContext;
        protected IBindingContextItem[,] contextItems;
        protected object[,] cells;

        public List<ExcelRenderer> NestedRenderer
        { get; private set; }

        public bool IsDisposed
        { get; protected set; }

        public ExcelRootRenderer RootRenderer
        { get; private set; }

        public ExcelRenderer Parent
        { get; private set; }

        public ExcelPartRenderer HeaderPartRenderer
        { get; protected set; }

        public ExcelPartRenderer BodyPartRenderer
        { get; protected set; }

        public ExcelPartRenderer FooterPartRenderer
        { get; protected set; }

        public MethodInfo MinOccurencesMethod
        { get; private set; }

        public List<List<IBindingContextItem>> ContextItems
        { get; }

        public ExcelInterop.Range FirstOutputCell
        { get; private set; }

        public ExcelInterop.Range RenderedRange
        { get; protected set; }

        public RenderedArea RenderedArea
        { get; private set; }

        public int Width
        { get; private set; }

        public int Height
        { get; private set; }

        public bool HasExpander => templateDefinition.TemplateOption.HeaderAsExpander != HeaderAsExpander.None;

        public bool IsExpanded
        {
            get
            {
                //if ((FilterOwner as ExcelTemplateDefinition).ExpanderBindingDefinition != null)
                //    isExpanded = (bool)(FilterOwner as ExcelTemplateDefinition).ExpanderBindingDefinition.ResolveBinding(GetDataSource());
                return bindingContext.IsExpanded;
            }
            set
            {
                //&&if ((FilterOwner as ExcelTemplateDefinition).ExpanderBindingDefinition != null)
                //    isExpanded = (bool)(FilterOwner as ExcelTemplateDefinition).ExpanderBindingDefinition.UpdateDataSource(GetDataSource(), !isExpanded);
                //else
                bindingContext.IsExpanded = value;
            }
        }

        #region .ctors
        public ExcelRenderer(ExcelRenderer parent, ITemplateDefinition templateDefinition, IBindingContext bindingContext, ExcelInterop.Range firstOutputCell, MethodInfo minOccurencesMethod)
        {
            NestedRenderer = new List<ExcelRenderer>();
            if (parent == null)
                RootRenderer = this as ExcelRootRenderer;
            else
            {
                RootRenderer = parent.RootRenderer;
                Parent = parent;
            }

            this.templateDefinition = templateDefinition;
            this.bindingContext = bindingContext;
            this.FirstOutputCell = firstOutputCell;
            MinOccurencesMethod = minOccurencesMethod;
            ContextItems = new List<List<IBindingContextItem>>();
        }
        #endregion

        #region public methods
        //public virtual void Render()
        //{
        //    int[] xs = new int[3];
        //    int[] ys = new int[3];

        //    ExcelInterop.Range nextFirstOutputCell = null;
        //    if (templateDefinition.Header != null)
        //    {
        //        HeaderPartRenderer = ExcelPartRenderer.CreateInstance(this, (ExcelTemplateDefinitionPart)templateDefinition.Header, bindingContext.Header, FirstOutputCell, false);
        //        HeaderPartRenderer.Render();
        //        if (HeaderPartRenderer.RenderedArea != null && HeaderPartRenderer.RenderedArea.Width != 0)
        //        {
        //            xs[0] = HeaderPartRenderer.RenderedArea.Width;
        //            ys[0] = HeaderPartRenderer.RenderedArea.Height;

        //            int xOffset = templateDefinition.Orientation == Orientation.Horizontal ? xs[0] : 0;
        //            int yOffset = templateDefinition.Orientation == Orientation.Horizontal ? 0 : ys[0];
        //            nextFirstOutputCell = FirstOutputCell.Offset[yOffset, xOffset];
        //        }
        //    }

        //    if (templateDefinition.Body != null)
        //    {
        //        BodyPartRenderer = ExcelPartRenderer.CreateInstance(this, (ExcelTemplateDefinitionPart)templateDefinition.Body, bindingContext.Body, nextFirstOutputCell ?? FirstOutputCell, true);
        //        BodyPartRenderer.Render();
        //        if (BodyPartRenderer.RenderedArea != null && BodyPartRenderer.RenderedArea.Width != 0)
        //        {
        //            xs[1] = BodyPartRenderer.RenderedArea.Width;
        //            ys[1] = BodyPartRenderer.RenderedArea.Height;

        //            int xOffset = templateDefinition.Orientation == Orientation.Horizontal ? xs[1] : 0;
        //            int yOffset = templateDefinition.Orientation == Orientation.Horizontal ? 0 : ys[1];
        //            nextFirstOutputCell = (nextFirstOutputCell ?? FirstOutputCell).Offset[yOffset, xOffset];
        //        }
        //    }

        //    if (templateDefinition.Footer != null)
        //    {
        //        FooterPartRenderer = ExcelPartRenderer.CreateInstance(this, (ExcelTemplateDefinitionPart)templateDefinition.Footer, bindingContext.Footer, nextFirstOutputCell ?? FirstOutputCell, false);
        //        FooterPartRenderer.Render();
        //        if (FooterPartRenderer.RenderedArea != null && FooterPartRenderer.RenderedArea.Width != 0)
        //        {
        //            xs[2] = FooterPartRenderer.RenderedArea.Width;
        //            ys[2] = FooterPartRenderer.RenderedArea.Height;
        //        }
        //    }

        //    int width = templateDefinition.Orientation == Orientation.Vertical ? xs.Max() : xs.Sum();
        //    int height = templateDefinition.Orientation == Orientation.Vertical ? ys.Sum() : ys.Max();
        //    if (width > 0 && height > 0)
        //    {
        //        RenderedArea = new RenderedArea(FirstOutputCell.Column, FirstOutputCell.Row, width, height);
        //        RenderedRange = FirstOutputCell.Resize[height, width];
        //        Width = width;
        //        Height = height;
        //    }
        //}

        public virtual void Render()
        {
            Width = Height = 0;
            ExcelInterop.Range nextFirstOutputCell = null;
            if (templateDefinition.Header != null)
            {
                HeaderPartRenderer = ExcelPartRenderer.CreateInstance(this, (ExcelTemplateDefinitionPart) templateDefinition.Header, bindingContext.Header, FirstOutputCell, false);
                HeaderPartRenderer.Render();
                if (HeaderPartRenderer.RenderedArea != null && HeaderPartRenderer.RenderedArea.Width != 0)
                {
                    Width = HeaderPartRenderer.RenderedArea.Width;
                    Height = HeaderPartRenderer.RenderedArea.Height;
                    nextFirstOutputCell = FirstOutputCell.Offset[templateDefinition.Orientation == Orientation.Horizontal ? 0 : HeaderPartRenderer.RenderedArea.Height,
                                                                 templateDefinition.Orientation == Orientation.Horizontal ? HeaderPartRenderer.RenderedArea.Width : 0];
                }
            }

            if (templateDefinition.Body != null)
            {
                BodyPartRenderer = ExcelPartRenderer.CreateInstance(this, (ExcelTemplateDefinitionPart) templateDefinition.Body, bindingContext.Body, nextFirstOutputCell ?? FirstOutputCell, true);
                BodyPartRenderer.Render();
                if (BodyPartRenderer.RenderedArea != null && BodyPartRenderer.RenderedArea.Width != 0)
                {
                    Width = templateDefinition.Orientation == Orientation.Vertical ? Width > BodyPartRenderer.RenderedArea.Width ? Width : BodyPartRenderer.RenderedArea.Width
                                                                                   : Width + BodyPartRenderer.RenderedArea.Width;
                    Height = templateDefinition.Orientation == Orientation.Vertical ? Height + BodyPartRenderer.RenderedArea.Height
                                                                                    : Height > BodyPartRenderer.RenderedArea.Height ? Height : BodyPartRenderer.RenderedArea.Height;

                    nextFirstOutputCell = (nextFirstOutputCell ?? FirstOutputCell).Offset[templateDefinition.Orientation == Orientation.Horizontal ? 0 : BodyPartRenderer.RenderedArea.Height,
                                                                                          templateDefinition.Orientation == Orientation.Horizontal ? BodyPartRenderer.RenderedArea.Width : 0];
                }
            }

            if (templateDefinition.Footer != null)
            {
                FooterPartRenderer = ExcelPartRenderer.CreateInstance(this, (ExcelTemplateDefinitionPart) templateDefinition.Footer, bindingContext.Footer, nextFirstOutputCell ?? FirstOutputCell, false);
                FooterPartRenderer.Render();
                if (FooterPartRenderer.RenderedArea != null && FooterPartRenderer.RenderedArea.Width != 0)
                {
                    Width = templateDefinition.Orientation == Orientation.Vertical ? Width > FooterPartRenderer.RenderedArea.Width ? Width : FooterPartRenderer.RenderedArea.Width
                                                                                   : Width + FooterPartRenderer.RenderedArea.Width;
                    Height = templateDefinition.Orientation == Orientation.Vertical ? Height + FooterPartRenderer.RenderedArea.Height
                                                                                    : Height > FooterPartRenderer.RenderedArea.Height ? Height : FooterPartRenderer.RenderedArea.Height;
                }
            }

            if (Width > 0 && Height > 0)
            {
                RenderedArea = new RenderedArea(FirstOutputCell.Column, FirstOutputCell.Row, Width, Height);
                RenderedRange = FirstOutputCell.Resize[Height, Width];
            }
        }

        public void RegisterNestedRenderer(ExcelRenderer nestedRenderer)
        {
            NestedRenderer.Add(nestedRenderer);
        }

        public virtual void AddAfterRenderingAction(SpecificEventCallback callBack)
        {
            RootRenderer.AddAfterRenderingAction(callBack);
        }

        public void Dispose()
        {
            if (!IsDisposed)
            {
                ClearRenderingData();
                FirstOutputCell = null;
                IsDisposed = true;
            }
        }
        #endregion

        #region protected methods
        protected virtual void ClearRenderingData()
        {
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

            foreach (ExcelRenderer nestedRenderer in NestedRenderer)
                nestedRenderer.ClearRenderingData();

            NestedRenderer.Clear();
            ContextItems.Clear();
            RenderedRange = null;
            contextItems = null;
            cells = null;
        }
        #endregion
    }
}
