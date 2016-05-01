using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Etk.BindingTemplates.Context;
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
        protected ExcelInterop.Range firstOutputCell;
        protected IBindingContextItem[,] contextItems;
        protected object[,] cells;

        public ExcelRootRenderer RootRenderer
        { get; private set; }

        public ExcelPartRenderer HeaderPartRenderer
        { get; protected set; }

        public ExcelPartRenderer BodyPartRenderer
        { get; protected set; }

        public ExcelPartRenderer FooterPartRenderer
        { get; protected set; }

        public MethodInfo MinOccurencesMethod
        { get; private set; }

        public List<List<IBindingContextItem>> DataRows
        { get; private set; }

        public ExcelInterop.Range RenderedRange
        { get; protected set; }

        public RenderedArea RenderedArea
        { get; private set; }

        public int Width
        { get; private set; }

        public int Height
        { get; private set; }

        #region .ctors
        public ExcelRenderer(ExcelRootRenderer rootRenderer, ITemplateDefinition templateDefinition, IBindingContext bindingContext, ExcelInterop.Range firstOutputCell, MethodInfo minOccurencesMethod)
        {
            RootRenderer = rootRenderer ?? this as ExcelRootRenderer;
            this.templateDefinition = templateDefinition;
            this.bindingContext = bindingContext;
            this.firstOutputCell = firstOutputCell;
            MinOccurencesMethod = minOccurencesMethod;
            DataRows = new List<List<IBindingContextItem>>();
        }
        #endregion

        #region public methods
        public virtual void Render()
        {
            int[] xs = new int[3];
            int[] ys = new int[3];

            ExcelInterop.Range nextFirstOutputCell = null;
            if (templateDefinition.Header != null)
            {
                HeaderPartRenderer = ExcelPartRenderer.CreateInstance(this, (ExcelTemplateDefinitionPart)templateDefinition.Header, bindingContext.Header, firstOutputCell, false);
                HeaderPartRenderer.Render();
                if (HeaderPartRenderer.RenderedArea != null && HeaderPartRenderer.RenderedArea.Width != 0)
                {
                    xs[0] = HeaderPartRenderer.RenderedArea.Width;
                    ys[0] = HeaderPartRenderer.RenderedArea.Height;

                    int xOffset = templateDefinition.Orientation == Orientation.Horizontal ? xs[0] : 0;
                    int yOffset = templateDefinition.Orientation == Orientation.Horizontal ? 0 : ys[0];
                    nextFirstOutputCell = firstOutputCell.get_Offset(yOffset, xOffset);
                }
            }

            if (templateDefinition.Body != null)
            {
                BodyPartRenderer = ExcelPartRenderer.CreateInstance(this, (ExcelTemplateDefinitionPart)templateDefinition.Body, bindingContext.Body, nextFirstOutputCell ?? firstOutputCell, true);
                BodyPartRenderer.Render();
                if (BodyPartRenderer.RenderedArea != null && BodyPartRenderer.RenderedArea.Width != 0)
                {
                    xs[1] = BodyPartRenderer.RenderedArea.Width;
                    ys[1] = BodyPartRenderer.RenderedArea.Height;

                    int xOffset = templateDefinition.Orientation == Orientation.Horizontal ? xs[1] : 0;
                    int yOffset = templateDefinition.Orientation == Orientation.Horizontal ? 0 : ys[1];
                    nextFirstOutputCell = (nextFirstOutputCell ?? firstOutputCell).get_Offset(yOffset, xOffset);
                }
            }

            if (templateDefinition.Footer != null)
            {
                FooterPartRenderer = ExcelPartRenderer.CreateInstance(this, (ExcelTemplateDefinitionPart)templateDefinition.Footer, bindingContext.Footer, nextFirstOutputCell ?? firstOutputCell, false);
                FooterPartRenderer.Render();
                if (FooterPartRenderer.RenderedArea != null && FooterPartRenderer.RenderedArea.Width != 0)
                {
                    xs[2] = FooterPartRenderer.RenderedArea.Width;
                    ys[2] = FooterPartRenderer.RenderedArea.Height;
                }
            }

            int width = templateDefinition.Orientation == Orientation.Vertical ? xs.Max() : xs.Sum();
            int height = templateDefinition.Orientation == Orientation.Vertical ? ys.Sum() : ys.Max();

            if (width > 0 && height > 0)
            {
                RenderedArea = new RenderedArea(firstOutputCell.Column, firstOutputCell.Row, width, height);
                RenderedRange = firstOutputCell.Resize[height, width];
                Width = width;
                Height = height;
            }
        }

        public void Dispose()
        {
            if (HeaderPartRenderer != null)
                HeaderPartRenderer.Dispose();
            if (BodyPartRenderer != null)
                BodyPartRenderer.Dispose();
            if (FooterPartRenderer != null)
                FooterPartRenderer.Dispose();

            firstOutputCell = null;
            RenderedRange = null;

            contextItems = null;
            cells = null;
        }
        #endregion
    }
}
