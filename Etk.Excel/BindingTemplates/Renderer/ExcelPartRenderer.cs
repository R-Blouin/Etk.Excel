using System;
using System.Linq;
using System.Runtime.InteropServices;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.Excel.BindingTemplates.Controls;
using Etk.Excel.BindingTemplates.Definitions;
using Etk.Excel.BindingTemplates.Views;
using Microsoft.Office.Interop.Excel;

namespace Etk.Excel.BindingTemplates.Renderer
{
    abstract class ExcelPartRenderer : IDisposable
    {
        #region attributes and properties
        protected static EnumManager enumManager = new EnumManager();
        protected static MultiLineManager multiLineManager = new MultiLineManager();

        protected bool useDecorator;

        protected ExcelRenderer Parent;
        protected ExcelTemplateDefinitionPart partToRenderDefinition;
        protected IBindingContextPart bindingContextPart;

        protected Range firstRangeTo;
        protected Range elementFirstRangeTo;

        protected Range currentRenderingFrom;
        protected Range currentRenderingTo;

        public int Height
        { get; protected set; }

        public int Width
        { get; protected set; }

        //public Range RenderedRange
        //{ get; private set; }

        public RenderedArea RenderedArea
        { get; protected set; }

        public bool isExpander = false;
        #endregion

        #region .ctors and factories
        protected ExcelPartRenderer(ExcelRenderer parent, ExcelTemplateDefinitionPart part, IBindingContextPart bindingContextPart, Range firstOutputCell, bool useDecorator)
        {
            this.Parent = parent;
            this.partToRenderDefinition = part;
            this.bindingContextPart = bindingContextPart;
            this.useDecorator = useDecorator;

            currentRenderingFrom = partToRenderDefinition.DefinitionFirstCell;
            firstRangeTo = elementFirstRangeTo = currentRenderingTo = firstOutputCell;

            Height = Width = 0;
        }

        public static ExcelPartRenderer CreateInstance(ExcelRenderer parent, ExcelTemplateDefinitionPart part, IBindingContextPart bindingContextPart, Range firstOutputCell, bool useDecorator)
        {
            if (part.Parent.Orientation == Orientation.Vertical)
                return new ExcelPartVerticalRenderer(parent, part, bindingContextPart, firstOutputCell, useDecorator);

            return new ExcelPartHorozontalRenderer(parent, part, bindingContextPart, firstOutputCell, useDecorator);
        }
        #endregion

        #region public methods
        public void Render()
        {
            Worksheet worksheetTo = currentRenderingTo.Worksheet;
            if (bindingContextPart != null && bindingContextPart.ElementsToRender != null && bindingContextPart.ElementsToRender.Any())
            {
                if (partToRenderDefinition.HasLinkedTemplates || partToRenderDefinition.ContainMultiLinesCells)
                    ManageTemplateWithLinkedTemplates();
                else
                    ManageTemplateWithoutLinkedTemplates();
            }
            if (Width > 0 && Height > 0)
                RenderedArea = new RenderedArea(firstRangeTo.Column, firstRangeTo.Row, Width, Height);
            Marshal.ReleaseComObject(worksheetTo);
        }

        public void Dispose()
        {
            //Marshal.ReleaseComObject(firstRangeTo);
            //Marshal.ReleaseComObject(elementFirstRangeTo);
            //Marshal.ReleaseComObject(currentRenderingFrom);
            //Marshal.ReleaseComObject(currentRenderingTo);

            firstRangeTo = null;
            elementFirstRangeTo = null;
            currentRenderingFrom = null;
            currentRenderingTo = null;
        }
        #endregion

        protected abstract void ManageTemplateWithoutLinkedTemplates();
        protected abstract void ManageTemplateWithLinkedTemplates();

        protected void ManageControls(IBindingContextItem item, ref Range range)
        {
            if (item is IExcelControl)
                ((IExcelControl)item).CreateControl(range);
        }
    }
}
