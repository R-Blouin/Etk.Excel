using Etk.Excel.BindingTemplates.Decorators;

namespace Etk.Excel.BindingTemplates.Renderer
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;
    using Etk.BindingTemplates.Context;
    using Etk.BindingTemplates.Definitions.Templates;
    using Etk.Excel.BindingTemplates.Controls;
    using Etk.Excel.BindingTemplates.Definitions;
    using Microsoft.Office.Interop.Excel;

    class RenderingContext
    {
        public int RowId { get; private set; }
        public List<IBindingContextItem> DataRow { get; private set; }
        public IBindingContextElement ContextElement { get; private set; }

        public LinkedTemplateDefinition LinkedTemplateDefinition { get; set; }

        public int CurrentRowHeight { get; set; }
        public int CurrentRowWidth { get; set; }
        public int PosCurrentLink { get; set; }
        public int PosPreviousLink { get; set; }
        public int LinkedViewRenderedWidth { get; set; }
        public int RefRow { get; set; }
        public bool RowAdded { get; set; }

        public RenderingContext(IBindingContextElement contextElement, int rowId)
        {
            ContextElement = contextElement;
            RowId = rowId;
            DataRow = new List<IBindingContextItem>();
        }
    }

    class ExcelPartVerticalRenderer : ExcelPartRenderer
    {
        #region .ctors and factories
        public ExcelPartVerticalRenderer(ExcelRenderer parent, ExcelTemplateDefinitionPart part, IBindingContextPart bindingContextPart, Range firstOutputCell, bool useDecorator)
                                        : base(parent, part, bindingContextPart, firstOutputCell, useDecorator)
        {}
        #endregion

        #region private methods
        protected override void ManageTemplateWithoutLinkedTemplates()
        { 
            Range firstCell = currentRenderingTo;
            Worksheet worksheetTo = currentRenderingTo.Worksheet;
            int cptElements = 0;

            int nbrOfElement = bindingContextPart.ElementsToRender.Count();
            int localWidth = partToRenderDefinition.Width;
            int localHeight = partToRenderDefinition.Height * nbrOfElement;
            Range workingRange = currentRenderingTo.Resize[localHeight, localWidth];

            partToRenderDefinition.DefinitionCells.Copy(workingRange);
            currentRenderingTo = worksheetTo.Cells[currentRenderingTo.Row + localHeight, currentRenderingTo.Column + localWidth];

            foreach (IBindingContextElement contextElement in bindingContextPart.ElementsToRender)
            {
                int cptItems = 0;
                for (int rowId = 0; rowId < partToRenderDefinition.Height; rowId++)
                {
                    List<IBindingContextItem> row = new List<IBindingContextItem>();
                    Parent.DataRows.Add(row);
                    for (int colId = 0; colId < partToRenderDefinition.Width; colId++)
                    {
                        IBindingContextItem item = partToRenderDefinition.DefinitionParts[rowId, colId] == null ? null : contextElement.BindingContextItems[cptItems++];
                        if (item != null && ((item.BindingDefinition != null && item.BindingDefinition.IsEnum && !item.BindingDefinition.IsReadOnly) || item is IExcelControl))
                        {
                            Range range = worksheetTo.Cells[firstCell.Row + rowId + cptElements * partToRenderDefinition.Height, firstCell.Column + colId];
                            if (item.BindingDefinition != null && item.BindingDefinition.IsEnum)
                                enumManager.CreateControl(item, ref range);
                            else
                                ManageControls(item, ref range);
                            range = null;
                        }
                        row.Add(item);
                    }
                }
                if (useDecorator && ((ExcelTemplateDefinition) partToRenderDefinition.Parent).Decorator != null)
                {
                    Range elementRange = firstCell.Offset[cptElements, 0];
                    elementRange = elementRange.Resize[1, localWidth];

                    Parent.RootRenderer.RowDecorators.Add(new ExcelElementDecorator(elementRange, ((ExcelTemplateDefinition)partToRenderDefinition.Parent).Decorator, contextElement));
                }
                cptElements++;
            }
            Height += localHeight;
            if (Width < localWidth)
                Width = localWidth;

            Marshal.ReleaseComObject(worksheetTo);
            Marshal.ReleaseComObject(workingRange);
            firstCell = null;
        }

        protected override void ManageTemplateWithLinkedTemplates()
        {
            Worksheet worksheetTo = currentRenderingTo.Worksheet;
            Width = partToRenderDefinition.Width;
            foreach (IBindingContextElement contextElement in bindingContextPart.ElementsToRender)
            {
                Range firstElementCell = currentRenderingTo;
                int bindingContextItemsCpt = 0;
                int cptLinkedDefinition = 0;
                int elementHeight = 0;
                int elementWidth = partToRenderDefinition.Width;
                for (int rowId = 0; rowId < partToRenderDefinition.Height; rowId++)
                {
                    RenderingContext renderingContext = new RenderingContext(contextElement, rowId);

                    List<int> posLinks = partToRenderDefinition.PositionLinkedTemplates[rowId];
                    if (posLinks == null)
                    {
                        Parent.DataRows.Add(renderingContext.DataRow);
                        int vOffset = 1;
                        ManageTemplatePart(renderingContext, ref bindingContextItemsCpt, ref vOffset, 0, partToRenderDefinition.Width);
                        currentRenderingTo = worksheetTo.Cells[currentRenderingTo.Row + vOffset, currentRenderingTo.Column];
                        Height += vOffset;
                    }
                    else
                    {
                        renderingContext.RefRow = Parent.DataRows.Count > 0 ? Parent.DataRows.Count : 0;
                        renderingContext.PosPreviousLink = 0;
                        int lastPosLink = posLinks.Count - 1;
                        for (int linkCpt = 0; linkCpt < posLinks.Count; linkCpt++)
                        {
                            renderingContext.LinkedViewRenderedWidth = 0;
                            renderingContext.PosCurrentLink = posLinks[linkCpt];
                            renderingContext.LinkedTemplateDefinition = partToRenderDefinition.DefinitionParts[rowId, renderingContext.PosCurrentLink] as LinkedTemplateDefinition;
                            // Render before link
                            if (renderingContext.PosCurrentLink > 0)
                                bindingContextItemsCpt = RenderBeforeLink(renderingContext, linkCpt, worksheetTo, bindingContextItemsCpt);

                            // Render link
                            IBindingContext linkedBindingContext = contextElement.LinkedBindingContexts[cptLinkedDefinition++];
                            if (linkedBindingContext.Body != null && linkedBindingContext.Body.ElementsToRender != null && linkedBindingContext.Body.ElementsToRender.Any())
                            {
                                RenderLink(renderingContext, linkedBindingContext, worksheetTo);
                                renderingContext.CurrentRowWidth += renderingContext.LinkedViewRenderedWidth;
                            }

                            // Render after link
                            if (linkCpt == lastPosLink && renderingContext.PosCurrentLink != partToRenderDefinition.Width)
                                bindingContextItemsCpt = RenderAfterLink(renderingContext, bindingContextItemsCpt);

                            if (renderingContext.CurrentRowWidth > elementWidth)
                                elementWidth = renderingContext.CurrentRowWidth;
                            if (renderingContext.CurrentRowHeight > elementHeight)
                                elementHeight = renderingContext.CurrentRowHeight;
                            renderingContext.PosPreviousLink = renderingContext.PosCurrentLink;
                        }
                        if (elementWidth > Width)
                            Width = elementWidth;
                        Height += elementHeight;
                        currentRenderingTo = worksheetTo.Cells[firstRangeTo.Row + Height, firstRangeTo.Column];
                    }
                }

                if (useDecorator && ((ExcelTemplateDefinition) partToRenderDefinition.Parent).Decorator != null)
                {
                    Range elementRange = firstElementCell.Resize[elementHeight, elementWidth];
                    Parent.RootRenderer.RowDecorators.Add(new ExcelElementDecorator(elementRange, ((ExcelTemplateDefinition)partToRenderDefinition.Parent).Decorator, contextElement));
                }
            }
            Marshal.ReleaseComObject(worksheetTo);
            worksheetTo = null;
        }

        private int RenderBeforeLink(RenderingContext renderingContext, int linkCpt, Worksheet worksheetTo, int bindingContextItemsCpt)
        {
            int firstCol, gap;
            if (linkCpt == 0)
            {
                firstCol = 0;
                gap = renderingContext.PosCurrentLink;
            }
            else
            {
                if (renderingContext.LinkedTemplateDefinition.Positioning == LinkedTemplatePositioning.Absolute)
                {
                    firstCol = renderingContext.CurrentRowWidth;
                    gap = renderingContext.PosCurrentLink - renderingContext.CurrentRowWidth;
                }
                else
                {
                    firstCol = renderingContext.PosPreviousLink + 1;
                    gap = renderingContext.PosCurrentLink - firstCol;
                }
            }
            if (gap > 0)
            {
                if (!renderingContext.RowAdded)
                {
                    AddRow(renderingContext);
                    renderingContext.RowAdded = true;
                }
                int vOffset = 1;
                ManageTemplatePart(renderingContext, ref bindingContextItemsCpt, ref vOffset, firstCol, firstCol + gap);
                currentRenderingTo = worksheetTo.Cells[currentRenderingTo.Row, currentRenderingTo.Column + gap];
                renderingContext.CurrentRowWidth += gap;
                if (vOffset > renderingContext.CurrentRowHeight)
                    renderingContext.CurrentRowHeight = vOffset;
            }
            return bindingContextItemsCpt;
        }

        private void RenderLink(RenderingContext renderingContext, IBindingContext linkedBindingContext, Worksheet worksheetTo)
        {
            using (ExcelRenderer linkedRenderer = new ExcelRenderer(Parent.RootRenderer, renderingContext.LinkedTemplateDefinition.TemplateDefinition, linkedBindingContext, currentRenderingTo))
            {
                linkedRenderer.Render();

                if (linkedRenderer.RenderedArea != null)
                {
                    renderingContext.LinkedViewRenderedWidth = linkedRenderer.Width;
                    if (!renderingContext.RowAdded)
                    {
                        AddRow(renderingContext);
                        renderingContext.RowAdded = true;
                    }

                    renderingContext.DataRow.AddRange(linkedRenderer.DataRows[0]);

                    for (int i = 1; i < linkedRenderer.Height; i++) //for (int i = 1; i < linkedRenderer.DataRow.Count; i++)
                    {
                        List<IBindingContextItem> rowToUse;
                        if (i >= renderingContext.CurrentRowHeight)
                        {
                            rowToUse = renderingContext.CurrentRowWidth > 0 ? new List<IBindingContextItem>(new IBindingContextItem[renderingContext.CurrentRowWidth])
                                                                            : new List<IBindingContextItem>();
                            Parent.DataRows.Add(rowToUse);
                        }
                        else
                        {
                            rowToUse = Parent.DataRows[i + renderingContext.RefRow];
                            if (rowToUse.Count < renderingContext.CurrentRowWidth)
                                rowToUse.AddRange(new IBindingContextItem[renderingContext.CurrentRowWidth - rowToUse.Count]);
                        }
                        rowToUse.AddRange(linkedRenderer.DataRows[i]);
                    }

                    // To take the multilines into account
                    if (linkedRenderer.Height > linkedRenderer.DataRows.Count)
                    {
                        for (int cpt = linkedRenderer.DataRows.Count + 1; cpt <= linkedRenderer.Height; cpt++)
                        {
                            //Parent.DataRow.Add(new List<IBindingContextItem>(new IBindingContextItem[0]));
                            List<IBindingContextItem> rowToUse;
                            if (cpt >= renderingContext.CurrentRowHeight)
                            {
                                rowToUse = renderingContext.CurrentRowWidth > 0 ? new List<IBindingContextItem>(new IBindingContextItem[renderingContext.CurrentRowWidth])
                                                                                : new List<IBindingContextItem>();
                                Parent.DataRows.Add(rowToUse);
                            }
                            else
                            {
                                rowToUse = Parent.DataRows[cpt + renderingContext.RefRow];
                                if (rowToUse.Count < renderingContext.CurrentRowWidth)
                                    rowToUse.AddRange(new IBindingContextItem[renderingContext.CurrentRowWidth - rowToUse.Count]);
                            }
                            rowToUse.AddRange(new IBindingContextItem[linkedRenderer.Width]);
                        }
                    }

                    if (renderingContext.CurrentRowHeight < linkedRenderer.Height)
                        renderingContext.CurrentRowHeight = linkedRenderer.Height;

                    currentRenderingTo = worksheetTo.Cells[currentRenderingTo.Row, currentRenderingTo.Column + linkedRenderer.Width];
                }
            }
        }

        private int RenderAfterLink(RenderingContext renderingContext, int bindingContextItemsCpt)
        {
            int vOffset = 1;
            int startPosition = renderingContext.PosCurrentLink + 1;
            if (renderingContext.LinkedTemplateDefinition.Positioning == LinkedTemplatePositioning.Absolute)
            {
                int afterWrittenPosition = renderingContext.PosCurrentLink + renderingContext.LinkedViewRenderedWidth;
                for (int i = startPosition; i < afterWrittenPosition; i++)
                {
                    if (partToRenderDefinition.DefinitionParts[renderingContext.RowId, i] != null)
                        bindingContextItemsCpt++;
                }
                startPosition = afterWrittenPosition;
            }
            if (startPosition < partToRenderDefinition.Width)
            {
                int realEnd = partToRenderDefinition.Width;
                for (int i = partToRenderDefinition.Width - 1; i >= startPosition; i--)
                {
                    if (partToRenderDefinition.DefinitionParts[renderingContext.RowId, i] != null)
                        break;
                    realEnd--;
                }
                if (realEnd > startPosition)
                {
                    ManageTemplatePart(renderingContext, ref bindingContextItemsCpt, ref vOffset, startPosition, realEnd);
                    renderingContext.CurrentRowWidth += realEnd - startPosition;
                }
            }
            if (vOffset > renderingContext.CurrentRowHeight)
                renderingContext.CurrentRowHeight = vOffset;
            return bindingContextItemsCpt;
        }

        private void AddRow(RenderingContext renderingContext)
        {
            Parent.DataRows.Add(renderingContext.DataRow);
            renderingContext.CurrentRowHeight = 1;
        }

        private void ManageTemplatePart(RenderingContext renderingContext, ref int currentBindingContextItemId, ref int vOffset, int startPos, int endPos)
        {
            Worksheet worksheetFrom = partToRenderDefinition.DefinitionFirstCell.Worksheet;
            Worksheet worksheetTo = currentRenderingTo.Worksheet;

            int gap = endPos - startPos;
            Range source = worksheetFrom.Cells[partToRenderDefinition.DefinitionFirstCell.Row + renderingContext.RowId, partToRenderDefinition.DefinitionFirstCell.Column + startPos];
            source = source.Resize[1, gap];
            Range workingRange = currentRenderingTo.Resize[1, gap];
            source.Copy(workingRange);

            for (int colId = startPos; colId < endPos; colId++)
            {
                IBindingContextItem item = partToRenderDefinition.DefinitionParts[renderingContext.RowId, colId] == null ? null : renderingContext.ContextElement.BindingContextItems[currentBindingContextItemId++];
                if (item != null && ((item.BindingDefinition != null && (item.BindingDefinition.IsEnum || item.BindingDefinition.IsMultiLine)) || item is IExcelControl))
                {
                    Range range = worksheetTo.Cells[currentRenderingTo.Row, currentRenderingTo.Column + colId];
                    if(item.BindingDefinition != null)
                    {
                        if (item.BindingDefinition.IsEnum && !item.BindingDefinition.IsReadOnly)
                            enumManager.CreateControl(item, ref range);
                        if (item.BindingDefinition.IsMultiLine)
                            multiLineManager.CreateControl(item, ref range, ref vOffset);
                    }
                    if (item is IExcelControl)
                        ManageControls(item, ref range);
                    range = null;
                }
                renderingContext.DataRow.Add(item);
            }

            Marshal.ReleaseComObject(worksheetFrom);
            Marshal.ReleaseComObject(worksheetTo);
            Marshal.ReleaseComObject(source);
            Marshal.ReleaseComObject(workingRange);
            worksheetFrom = null;
            worksheetTo = null;
            source = null;
            workingRange = null;
        }
        #endregion
    }
}
