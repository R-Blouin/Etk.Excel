using System;
using System.Linq;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.Excel.BindingTemplates.Controls;
using Etk.Excel.BindingTemplates.Definitions;
using Etk.Excel.BindingTemplates.Views;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Etk.BindingTemplates.Definitions.EventCallBacks;
using Etk.Excel.Application;

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

        protected ExcelInterop.Range firstRangeTo;
        protected ExcelInterop.Range elementFirstRangeTo;

        protected ExcelInterop.Range currentRenderingFrom;
        protected ExcelInterop.Range currentRenderingTo;

        internal ExcelInterop.Range RenderedRange
        { get; private set; }

        public int Height
        { get; protected set; }

        public int Width
        { get; protected set; }

        public RenderedArea RenderedArea
        { get; protected set; }

        //public bool isExpander = false;
        #endregion

        #region .ctors and factories
        protected ExcelPartRenderer(ExcelRenderer parent, ExcelTemplateDefinitionPart part, IBindingContextPart bindingContextPart, ExcelInterop.Range firstOutputCell, bool useDecorator)
        {
            Parent = parent;
            partToRenderDefinition = part;
            this.bindingContextPart = bindingContextPart;
            this.useDecorator = useDecorator;

            currentRenderingFrom = partToRenderDefinition.DefinitionFirstCell[1, 1];
            firstRangeTo = firstOutputCell[1, 1];
            elementFirstRangeTo = firstOutputCell[1, 1];
            currentRenderingTo = firstOutputCell[1, 1];

            Height = Width = 0;
        }

        public static ExcelPartRenderer CreateInstance(ExcelRenderer parent, ExcelTemplateDefinitionPart part, IBindingContextPart bindingContextPart, ExcelInterop.Range firstOutputCell, bool useDecorator)
        {
            if (part.Parent.Orientation == Orientation.Vertical)
                return new ExcelPartVerticalRenderer(parent, part, bindingContextPart, firstOutputCell, useDecorator);
            return new ExcelPartHorizontalRenderer(parent, part, bindingContextPart, firstOutputCell, useDecorator);
        }
        #endregion

        #region public methods
        public void Render()
        {
            if (bindingContextPart != null )
//                && ((bindingContextPart is LinkedTemplateDefinition && ((LinkedTemplateDefinition) bindingContextPart).MinOccurencesMethod != null || bindingContextPart.ElementsToRender.ElementsToRender != null && bindingContextPart.ElementsToRender.ElementsToRender.Any())
            {
                if (partToRenderDefinition.HasLinkedTemplates || partToRenderDefinition.ContainMultiLinesCells)
                    ManageTemplateWithLinkedTemplates();
                else
                    ManageTemplateWithoutLinkedTemplates();
            }
            if (Width > 0 && Height > 0)
            {
                RenderedArea = new RenderedArea(firstRangeTo.Column, firstRangeTo.Row, Width, Height);
                RenderedRange = firstRangeTo.Resize[Height, Width];
            }
        }

        public void Dispose()
        {
            ExcelApplication.ReleaseComObject(currentRenderingFrom);
            ExcelApplication.ReleaseComObject(firstRangeTo);
            ExcelApplication.ReleaseComObject(elementFirstRangeTo);
            ExcelApplication.ReleaseComObject(currentRenderingTo);

            if(RenderedRange != null)
                ExcelApplication.ReleaseComObject(RenderedRange);
        }
        #endregion

        #region protected method
        protected abstract void ManageTemplateWithoutLinkedTemplates();
        protected abstract void ManageTemplateWithLinkedTemplates();

        protected void AddAfterRenderingAction(IBindingDefinition bindingDefinition, ExcelInterop.Range concernedRange)
        {
            if (bindingDefinition.OnAfterRendering?.Parameters != null)
            {
                foreach (SpecificEventCallbackParameter param in bindingDefinition.OnAfterRendering.Parameters.Where(p => p.IsSender))
                    param.ParameterValue = concernedRange[1, 1];
            }
            Parent.AddAfterRenderingAction(bindingDefinition.OnAfterRendering);
        }
        #endregion
    }
}
