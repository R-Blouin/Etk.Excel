using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.Excel.BindingTemplates.Controls;
using Etk.Excel.BindingTemplates.Decorators;
using Etk.Excel.BindingTemplates.Definitions;
using Etk.Excel.BindingTemplates.SortSearchAndFilter;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Etk.BindingTemplates.Context.SortSearchAndFilter;

namespace Etk.Excel.BindingTemplates.Renderer
{
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
        public ExcelPartVerticalRenderer(ExcelRenderer parent, ExcelTemplateDefinitionPart part, IBindingContextPart bindingContextPart, ExcelInterop.Range firstOutputCell, bool useDecorator)
                                        : base(parent, part, bindingContextPart, firstOutputCell, useDecorator)
        {}
        #endregion

        #region private methods
        protected override void ManageTemplateWithoutLinkedTemplates()
        {
            ExcelInterop.Range firstCell = currentRenderingTo;
            ExcelInterop.Worksheet worksheetTo = currentRenderingTo.Worksheet;
            int cptElements = 0;

            int nbrOfElement = bindingContextPart.ElementsToRender.Count();
            int localWidth = partToRenderDefinition.Width;
            int localHeight = partToRenderDefinition.Height * nbrOfElement;
            if (localHeight > 0)
            {
                ExcelInterop.Range workingRange = currentRenderingTo.Resize[localHeight, localWidth];

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
                            if (item != null && ((item.BindingDefinition != null && (item.BindingDefinition.IsEnum)) || item is IExcelControl || item is ExcelBindingSearchContextItem))
                            {
                                ExcelInterop.Range range = worksheetTo.Cells[firstCell.Row + rowId + cptElements * partToRenderDefinition.Height, firstCell.Column + colId];

                                if (item is ExcelBindingSearchContextItem)
                                    ((ExcelBindingSearchContextItem)item).SetRange(ref range);

                                this.ManageControls(item, ref range);
                                range = null;
                            }
                            row.Add(item);
                        }
                    }
                    if (useDecorator && ((ExcelTemplateDefinition) partToRenderDefinition.Parent).Decorator != null)
                    {
                        ExcelInterop.Range elementRange = firstCell.Offset[cptElements, 0];
                        elementRange = elementRange.Resize[1, localWidth];

                        Parent.RootRenderer.RowDecorators.Add(new ExcelElementDecorator(elementRange, ((ExcelTemplateDefinition)partToRenderDefinition.Parent).Decorator, contextElement));
                    }
                    cptElements++;
                }
                Marshal.ReleaseComObject(workingRange);
            }

            // To take into account the min number of elements to render.
            if( Parent.MinOccurencesMethod != null)
            {
                IBindingContextElement parentElement = null;
                if (bindingContextPart.ParentContext != null)
                    parentElement = bindingContextPart.ParentContext.Parent;

                int minElementsToRender = LinkedTemplateDefinition.ResolveMinOccurences(Parent.MinOccurencesMethod, parentElement);
                if (minElementsToRender > nbrOfElement)
                    localHeight = partToRenderDefinition.Height * minElementsToRender;
            }

            Height += localHeight;
            if (Width < localWidth)
                Width = localWidth;

            Marshal.ReleaseComObject(worksheetTo);
            firstCell = null;
        }

        protected override void ManageTemplateWithLinkedTemplates()
        {
            ExcelInterop.Worksheet worksheetTo = currentRenderingTo.Worksheet;
            Width = partToRenderDefinition.Width;
            foreach (IBindingContextElement contextElement in bindingContextPart.ElementsToRender)
            {
                ExcelInterop.Range firstElementCell = currentRenderingTo;
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
                            // RenderView before link
                            if (renderingContext.PosCurrentLink > 0)
                                bindingContextItemsCpt = RenderBeforeLink(renderingContext, linkCpt, worksheetTo, bindingContextItemsCpt);

                            // RenderView link
                            IBindingContext linkedBindingContext = contextElement.LinkedBindingContexts[cptLinkedDefinition++];
                            if (linkedBindingContext.Body != null && (renderingContext.LinkedTemplateDefinition.MinOccurencesMethod != null || linkedBindingContext.Body.ElementsToRender != null && linkedBindingContext.Body.ElementsToRender.Any()))
                            {
                                RenderLink(renderingContext, linkedBindingContext, worksheetTo);
                                renderingContext.CurrentRowWidth += renderingContext.LinkedViewRenderedWidth;
                            }

                            // RenderView after last link
                            if (linkCpt == lastPosLink && renderingContext.PosCurrentLink + 1 < partToRenderDefinition.Width)
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
                    ExcelInterop.Range elementRange = firstElementCell.Resize[elementHeight, elementWidth];
                    Parent.RootRenderer.RowDecorators.Add(new ExcelElementDecorator(elementRange, ((ExcelTemplateDefinition)partToRenderDefinition.Parent).Decorator, contextElement));
                }
            }
            Marshal.ReleaseComObject(worksheetTo);
            worksheetTo = null;
        }

        private int RenderBeforeLink(RenderingContext renderingContext, int linkCpt, ExcelInterop.Worksheet worksheetTo, int bindingContextItemsCpt)
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
                {
                    for (int i = renderingContext.CurrentRowHeight; i < vOffset; i++)
                        Parent.DataRows.Add(new List<IBindingContextItem>(new IBindingContextItem[gap]));
                    renderingContext.CurrentRowHeight = vOffset;
                }
            }
            return bindingContextItemsCpt;
        }

        private void RenderLink(RenderingContext renderingContext, IBindingContext linkedBindingContext, ExcelInterop.Worksheet worksheetTo)
        {
            ExcelRenderer linkedRenderer = new ExcelRenderer(Parent, renderingContext.LinkedTemplateDefinition.TemplateDefinition, linkedBindingContext, currentRenderingTo,
                                                                renderingContext.LinkedTemplateDefinition.MinOccurencesMethod);
            Parent.RegisterNestedRenderer(linkedRenderer);

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

        private int RenderAfterLink(RenderingContext renderingContext, int bindingContextItemsCpt)
        {
            int vOffset = 1;
            int startPosition = renderingContext.PosCurrentLink + 1;
            if (renderingContext.LinkedTemplateDefinition.Positioning == LinkedTemplatePositioning.Absolute 
                )//&& renderingContext.ContextElement != null && renderingContext.ContextElement.BindingContextItems.Count > renderingContext.PosCurrentLink)
            {
                int afterWrittenPosition = renderingContext.PosCurrentLink + renderingContext.LinkedViewRenderedWidth;
                for (int i = startPosition; i < afterWrittenPosition; i++)
                {
                    if (partToRenderDefinition.DefinitionParts[renderingContext.RowId, i] != null)
                        bindingContextItemsCpt++;
                }
                startPosition = afterWrittenPosition;
                //if (partToRenderDefinition.Width > startPosition + partToRenderDefinition.DefinitionFirstCell.Column)
                //{
                //    ManageTemplatePart(renderingContext, ref bindingContextItemsCpt, ref vOffset, startPosition, realEnd);
                //    renderingContext.CurrentRowWidth += realEnd - startPosition;
                //}
            }
            if (startPosition < partToRenderDefinition.Width)
            {
                int realEnd = partToRenderDefinition.Width;
                for (int i = partToRenderDefinition.Width - 1; i >= startPosition; i--)
                {
                    ExcelInterop.Range current = partToRenderDefinition.DefinitionFirstCell[1, i];
                    if (current.MergeCells || partToRenderDefinition.DefinitionParts[renderingContext.RowId, i] != null)
                        break;
                    realEnd--;
                }
                if (realEnd > startPosition)
                {
                    if (renderingContext.LinkedViewRenderedWidth + startPosition - 1 > renderingContext.DataRow.Count)
                        renderingContext.DataRow.AddRange(new IBindingContextItem[renderingContext.LinkedViewRenderedWidth + startPosition - 1 - renderingContext.DataRow.Count]);

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
            ExcelInterop.Worksheet worksheetFrom = partToRenderDefinition.DefinitionFirstCell.Worksheet;
            ExcelInterop.Worksheet worksheetTo = currentRenderingTo.Worksheet;

            int gap = endPos - startPos;
            ExcelInterop.Range source = worksheetFrom.Cells[partToRenderDefinition.DefinitionFirstCell.Row + renderingContext.RowId, partToRenderDefinition.DefinitionFirstCell.Column + startPos];
            source = source.Resize[1, gap];
            ExcelInterop.Range workingRange = currentRenderingTo.Resize[1, gap];
            source.Copy(workingRange);

            int bindingContextItemsCount =  renderingContext.ContextElement.BindingContextItems.Count;
            for (int colId = startPos; colId < endPos; colId++)
            {
                IBindingContextItem item = partToRenderDefinition.DefinitionParts[renderingContext.RowId, colId] == null || bindingContextItemsCount <= currentBindingContextItemId 
                                           ? null 
                                           : renderingContext.ContextElement.BindingContextItems[currentBindingContextItemId++];
                if (item != null && ((item.BindingDefinition != null && (item.BindingDefinition.IsEnum || item.BindingDefinition.IsMultiLine)) || item is IExcelControl || item is ExcelBindingSearchContextItem))
                {                  
                    ExcelInterop.Range range = worksheetTo.Cells[currentRenderingTo.Row, currentRenderingTo.Column + colId - startPos];

                    if (item is ExcelBindingSearchContextItem)
                        ((ExcelBindingSearchContextItem)item).SetRange(ref range);

                    if(item.BindingDefinition != null)
                    {
                        if (item.BindingDefinition.IsEnum && !item.BindingDefinition.IsReadOnly)
                            enumManager.CreateControl(item, ref range);
                        if (item.BindingDefinition.IsMultiLine)
                        {
                            ExcelInterop.Range localSource = source[1, 1 + colId - startPos];
                            multiLineManager.CreateControl(item, ref range, ref localSource, ref vOffset);
                        }
                    }
                    this.ManageControls(item, ref range);
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
