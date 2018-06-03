using System.Collections.Generic;
using System.Linq;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.Excel.Application;
using Etk.Excel.BindingTemplates.Controls;
using Etk.Excel.BindingTemplates.Decorators;
using Etk.Excel.BindingTemplates.Definitions;
using Etk.Excel.BindingTemplates.SortSearchAndFilter;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.BindingTemplates.Renderer
{
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
            ExcelInterop.Range firstCell = currentRenderingTo[1, 1];
            int cptElements = 0;

            int nbrOfElement;
            // To take into account the number of elements to render.
            if (Parent.NumberOfOccurencesMethod != null)
            {
                IBindingContextElement parentElement = null;
                if (bindingContextPart.ParentContext != null)
                    parentElement = bindingContextPart.ParentContext.Parent;

                nbrOfElement = LinkedTemplateDefinition.ResolveNumberOfOccurences(Parent.NumberOfOccurencesMethod, parentElement);
            }
            else
                nbrOfElement = bindingContextPart.ElementsToRender.Count();


            int localWidth = partToRenderDefinition.Width;
            int localHeight = partToRenderDefinition.Height * nbrOfElement;
            if (localHeight > 0)
            {
                ExcelInterop.Range workingRange = currentRenderingTo.Resize[localHeight, localWidth];

                partToRenderDefinition.DefinitionCells.Copy(workingRange);
                currentRenderingTo = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row + localHeight, currentRenderingTo.Column + localWidth];
                foreach (IBindingContextElement contextElement in bindingContextPart.ElementsToRender)
                {
                    int cptItems = 0;
                    for (int rowId = 0; rowId < partToRenderDefinition.Height; rowId++)
                    {
                        List<IBindingContextItem> row = new List<IBindingContextItem>();
                        Parent.ContextItems.Add(row);
                        for (int colId = 0; colId < partToRenderDefinition.Width; colId++)
                        {
                            IBindingContextItem item = partToRenderDefinition.DefinitionParts[rowId, colId] == null ? null : contextElement.BindingContextItems[cptItems++];
                            if (item != null)
                            {
                                if (item is ExcelBindingSearchContextItem)
                                {
                                    ExcelInterop.Range range = Parent.RootRenderer.View.ViewSheet.Cells[firstCell.Row + rowId + cptElements * partToRenderDefinition.Height, firstCell.Column + colId];
                                    ((ExcelBindingSearchContextItem)item).SetRange(range);
                                    ExcelApplication.ReleaseComObject(range);
                                }
                                else if (item is IExcelControl)
                                {
                                    ExcelInterop.Range range = Parent.RootRenderer.View.ViewSheet.Cells[firstCell.Row + rowId + cptElements * partToRenderDefinition.Height, firstCell.Column + colId];
                                    ((IExcelControl)item).CreateControl(range);
                                    ExcelApplication.ReleaseComObject(range);
                                }
                                else if (item.BindingDefinition != null)
                                {
                                    if (item.BindingDefinition.IsEnum && !item.BindingDefinition.IsReadOnly)
                                    {
                                        ExcelInterop.Range range = Parent.RootRenderer.View.ViewSheet.Cells[firstCell.Row + rowId + cptElements * partToRenderDefinition.Height, firstCell.Column + colId];
                                        enumManager.CreateControl(item, range);
                                        ExcelApplication.ReleaseComObject(range);
                                    }
                                    if (item.BindingDefinition.OnAfterRendering != null)
                                    {
                                        ExcelInterop.Range range = Parent.RootRenderer.View.ViewSheet.Cells[firstCell.Row + rowId + cptElements * partToRenderDefinition.Height, firstCell.Column + colId];
                                        AddAfterRenderingAction(item.BindingDefinition, range);
                                        ExcelApplication.ReleaseComObject(range);
                                    }
                                }
                            }
                            row.Add(item);
                        }
                    }
                    if (useDecorator && ((ExcelTemplateDefinition) partToRenderDefinition.Parent).Decorator != null)
                    {
                        ExcelInterop.Range elementRange = firstCell.Offset[cptElements, 0];
                        Parent.RootRenderer.RowDecorators.Add(new ExcelElementDecorator(elementRange, 1, localWidth, ((ExcelTemplateDefinition)partToRenderDefinition.Parent).Decorator, contextElement));
                        ExcelApplication.ReleaseComObject(elementRange);
                    }
                    cptElements++;
                }
                ExcelApplication.ReleaseComObject(workingRange);
            }

            // To take into account the min number of elements to render.
            if ( Parent.MinOccurencesMethod != null)
            {
                IBindingContextElement parentElement = null;
                if (bindingContextPart.ParentContext != null)
                    parentElement = bindingContextPart.ParentContext.Parent;

                int minElementsToRender = LinkedTemplateDefinition.ResolveNumberOfOccurences(Parent.MinOccurencesMethod, parentElement);
                if (minElementsToRender > nbrOfElement)
                    localHeight = partToRenderDefinition.Height * minElementsToRender;
            }

            Height += localHeight;
            if (Width < localWidth)
                Width = localWidth;

            ExcelApplication.ReleaseComObject(firstCell);
        }

        protected override void ManageTemplateWithLinkedTemplates()
        {
            Width = partToRenderDefinition.Width;
            foreach (IBindingContextElement contextElement in bindingContextPart.ElementsToRender)
            {
                ExcelInterop.Range firstElementCell = currentRenderingTo[1, 1];
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
                        Parent.ContextItems.Add(renderingContext.ContextItems);
                        int vOffset = 1;
                        ManageTemplatePart(renderingContext, ref bindingContextItemsCpt, ref vOffset, 0, partToRenderDefinition.Width);
                        currentRenderingTo = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row + vOffset, currentRenderingTo.Column];
                        elementHeight += vOffset;
                    }
                    else
                    {
                        int currentHeight = 0;
                        renderingContext.RefPos = Parent.ContextItems.Count > 0 ? Parent.ContextItems.Count : 0;
                        renderingContext.PosPreviousLink = 0;
                        int lastPosLink = posLinks.Count - 1;
                        for (int linkCpt = 0; linkCpt < posLinks.Count; linkCpt++)
                        {
                            renderingContext.LinkedViewRenderedOffset = 0;
                            renderingContext.PosCurrentLink = posLinks[linkCpt];
                            renderingContext.LinkedTemplateDefinition = partToRenderDefinition.DefinitionParts[rowId, renderingContext.PosCurrentLink] as LinkedTemplateDefinition;
                            // RenderView before link
                            if (renderingContext.PosCurrentLink > 0)
                                bindingContextItemsCpt = RenderBeforeLink(renderingContext, linkCpt, bindingContextItemsCpt);

                            // RenderView link
                            IBindingContext linkedBindingContext = contextElement.LinkedBindingContexts[cptLinkedDefinition++];
                            if (linkedBindingContext.Body != null && (renderingContext.LinkedTemplateDefinition.MinOccurencesMethod != null || linkedBindingContext.Body.ElementsToRender != null && linkedBindingContext.Body.ElementsToRender.Any()))
                            {
                                RenderLink(renderingContext, linkedBindingContext);
                                renderingContext.CurrentWidth += renderingContext.LinkedViewRenderedOffset;
                            }

                            // RenderView after last link
                            if (linkCpt == lastPosLink && partToRenderDefinition.Width > renderingContext.PosCurrentLink + 1)
                                bindingContextItemsCpt = RenderAfterLink(renderingContext, bindingContextItemsCpt);

                            if (renderingContext.CurrentWidth > elementWidth)
                                elementWidth = renderingContext.CurrentWidth;
                            if (renderingContext.CurrentHeight > currentHeight)
                                currentHeight = renderingContext.CurrentHeight;
                            renderingContext.PosPreviousLink = renderingContext.PosCurrentLink;
                        }
                        elementHeight += currentHeight;
                        if (elementWidth > Width)
                            Width = elementWidth;
                        currentRenderingTo = Parent.RootRenderer.View.ViewSheet.Cells[firstElementCell.Row + elementHeight, firstRangeTo.Column];
                    }
                }

                if (useDecorator && ((ExcelTemplateDefinition) partToRenderDefinition.Parent).Decorator != null)
                {
                    ExcelInterop.Range elementRange = firstElementCell.Resize[elementHeight, elementWidth];
                    Parent.RootRenderer.RowDecorators.Add(new ExcelElementDecorator(elementRange, elementHeight, elementWidth, ((ExcelTemplateDefinition)partToRenderDefinition.Parent).Decorator, contextElement));
                    ExcelApplication.ReleaseComObject(elementRange);
                }
                Height += elementHeight;
                ExcelApplication.ReleaseComObject(firstElementCell);
            }
        }

        private int RenderBeforeLink(RenderingContext renderingContext, int linkCpt, int bindingContextItemsCpt)
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
                    firstCol = renderingContext.CurrentWidth;
                    gap = renderingContext.PosCurrentLink - renderingContext.CurrentWidth;
                }
                else
                {
                    firstCol = renderingContext.PosPreviousLink + 1;
                    gap = renderingContext.PosCurrentLink - firstCol;
                }
            }
            if (gap > 0)
            {
                if (!renderingContext.RowColAdded)
                {
                    AddRow(renderingContext);
                    renderingContext.RowColAdded = true;
                }
                int vOffset = 1;
                ManageTemplatePart(renderingContext, ref bindingContextItemsCpt, ref vOffset, firstCol, firstCol + gap);
                currentRenderingTo = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row, currentRenderingTo.Column + gap];
                renderingContext.CurrentWidth += gap;
                if (vOffset > renderingContext.CurrentHeight)
                {
                    for (int i = renderingContext.CurrentHeight; i < vOffset; i++)
                        Parent.ContextItems.Add(new List<IBindingContextItem>(new IBindingContextItem[gap]));
                    renderingContext.CurrentHeight = vOffset;
                }
            }
            return bindingContextItemsCpt;
        }

        private void RenderLink(RenderingContext renderingContext, IBindingContext linkedBindingContext)
        {
            ExcelRenderer linkedRenderer = new ExcelRenderer(Parent, renderingContext.LinkedTemplateDefinition.TemplateDefinition, linkedBindingContext, currentRenderingTo,
                                                             renderingContext.LinkedTemplateDefinition.MinOccurencesMethod, renderingContext.LinkedTemplateDefinition.NumberOfOccurencesMethod);
            Parent.RegisterNestedRenderer(linkedRenderer);

            linkedRenderer.Render();

            if (linkedRenderer.RenderedArea != null)
            {
                renderingContext.LinkedViewRenderedOffset = linkedRenderer.Width;
                if (!renderingContext.RowColAdded)
                {
                    AddRow(renderingContext);
                    renderingContext.RowColAdded = true;
                }

                renderingContext.ContextItems.AddRange(linkedRenderer.ContextItems[0]);
                for (int i = 1; i < linkedRenderer.Height; i++)
                {
                    List<IBindingContextItem> toUse;
                    if (i >= renderingContext.CurrentHeight)
                    {
                        toUse = renderingContext.CurrentWidth > 0 ? new List<IBindingContextItem>(new IBindingContextItem[renderingContext.CurrentWidth])
                                                                  : new List<IBindingContextItem>();
                        Parent.ContextItems.Add(toUse);
                    }
                    else
                    {
                        toUse = Parent.ContextItems[i + renderingContext.RefPos];
                        if (toUse.Count < renderingContext.CurrentWidth)
                            toUse.AddRange(new IBindingContextItem[renderingContext.CurrentWidth - toUse.Count]);
                    }
                    toUse.AddRange(linkedRenderer.ContextItems[i]);
                }

                // To take the multilines into account
                if (linkedRenderer.Height > linkedRenderer.ContextItems.Count)
                {
                    for (int cpt = linkedRenderer.ContextItems.Count + 1; cpt <= linkedRenderer.Height; cpt++)
                    {
                        //Parent.DataRow.Add(new List<IBindingContextItem>(new IBindingContextItem[0]));
                        List<IBindingContextItem> toUse;
                        if (cpt >= renderingContext.CurrentHeight)
                        {
                            toUse = renderingContext.CurrentWidth > 0 ? new List<IBindingContextItem>(new IBindingContextItem[renderingContext.CurrentWidth])
                                                                      : new List<IBindingContextItem>();
                            Parent.ContextItems.Add(toUse);
                        }
                        else
                        {
                            toUse = Parent.ContextItems[cpt + renderingContext.RefPos];
                            if (toUse.Count < renderingContext.CurrentWidth)
                                toUse.AddRange(new IBindingContextItem[renderingContext.CurrentWidth - toUse.Count]);
                        }
                        toUse.AddRange(new IBindingContextItem[linkedRenderer.Width]);
                    }
                }

                if (renderingContext.CurrentHeight < linkedRenderer.Height)
                    renderingContext.CurrentHeight = linkedRenderer.Height;

                currentRenderingTo = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row, currentRenderingTo.Column + linkedRenderer.Width];
            }
        }

        private int RenderAfterLink(RenderingContext renderingContext, int bindingContextItemsCpt)
        {
            int vOffset = 1;
            int startPosition = renderingContext.PosCurrentLink + 1;
            if (renderingContext.LinkedTemplateDefinition.Positioning == LinkedTemplatePositioning.Absolute 
                )//&& renderingContext.ContextElement != null && renderingContext.ContextElement.BindingContextItems.Count > renderingContext.PosCurrentLink)
            {
                int afterWrittenPosition = renderingContext.PosCurrentLink + renderingContext.LinkedViewRenderedOffset;
                for (int i = startPosition; i < afterWrittenPosition; i++)
                {
                    if (partToRenderDefinition.DefinitionParts[renderingContext.InitPos, i] != null)
                        bindingContextItemsCpt++;
                }
                startPosition = afterWrittenPosition;
                //if (partToRenderDefinition.Width > startPosition + partToRenderDefinition.DefinitionFirstCell.Column)
                //{
                //    ManageTemplatePart(renderingContext, ref bindingContextItemsCpt, ref vOffset, startPosition, realEnd);
                //    renderingContext.CurrentWidth += realEnd - startPosition;
                //}
            }
            if (partToRenderDefinition.Width > startPosition)
            {
                int realEnd = partToRenderDefinition.Width;
                for (int i = partToRenderDefinition.Width - 1; i >= startPosition; i--)
                {
                    ExcelInterop.Range current = partToRenderDefinition.DefinitionFirstCell.Offset[0, 1];
                    if (current.MergeCells || partToRenderDefinition.DefinitionParts[renderingContext.InitPos, i] != null)
                    {
                        ExcelApplication.ReleaseComObject(current);
                        break;
                    }
                    ExcelApplication.ReleaseComObject(current);
                    realEnd--;
                }
                if (realEnd > startPosition)
                {
                    if (renderingContext.LinkedViewRenderedOffset + startPosition - 1 > renderingContext.ContextItems.Count)
                        renderingContext.ContextItems.AddRange(new IBindingContextItem[renderingContext.LinkedViewRenderedOffset + startPosition - 1 - renderingContext.ContextItems.Count]);

                    ManageTemplatePart(renderingContext, ref bindingContextItemsCpt, ref vOffset, startPosition, realEnd);
                    renderingContext.CurrentWidth += realEnd - startPosition;
                }
            }
            if (vOffset > renderingContext.CurrentHeight)
                renderingContext.CurrentHeight = vOffset;
            return bindingContextItemsCpt;
        }

        private void AddRow(RenderingContext renderingContext)
        {
            Parent.ContextItems.Add(renderingContext.ContextItems);
            renderingContext.CurrentHeight = 1;
        }

        private void ManageTemplatePart(RenderingContext renderingContext, ref int currentBindingContextItemId, ref int vOffset, int startPos, int endPos)
        {
            int gap = endPos - startPos;
            ExcelInterop.Range source = Parent.RootRenderer.View.TemplateSheet.Cells[partToRenderDefinition.DefinitionFirstCell.Row + renderingContext.InitPos, partToRenderDefinition.DefinitionFirstCell.Column + startPos];
            source = source.Resize[1, gap];
            ExcelInterop.Range workingRange = currentRenderingTo.Resize[1, gap];
            source.Copy(workingRange);

            int bindingContextItemsCount =  renderingContext.ContextElement.BindingContextItems.Count;
            for (int colId = startPos; colId < endPos; colId++)
            {
                IBindingContextItem item = partToRenderDefinition.DefinitionParts[renderingContext.InitPos, colId] == null || bindingContextItemsCount <= currentBindingContextItemId 
                                           ? null 
                                           : renderingContext.ContextElement.BindingContextItems[currentBindingContextItemId++];
                if (item != null)
                {
                    if (item is ExcelBindingSearchContextItem)
                    {
                        ExcelInterop.Range range = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row, currentRenderingTo.Column + colId - startPos];
                        ((ExcelBindingSearchContextItem)item).SetRange(range);
                        ExcelApplication.ReleaseComObject(range);
                    }
                    else if (item is IExcelControl)
                    {
                        ExcelInterop.Range range = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row, currentRenderingTo.Column + colId - startPos];
                        ((IExcelControl) item).CreateControl(range);
                        ExcelApplication.ReleaseComObject(range);
                    }
                    if (item.BindingDefinition != null)
                    {
                        if (item.BindingDefinition.IsEnum && !item.BindingDefinition.IsReadOnly)
                        {
                            ExcelInterop.Range range = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row, currentRenderingTo.Column + colId - startPos];
                            enumManager.CreateControl(item, range);
                            ExcelApplication.ReleaseComObject(range);
                        }
                        if (item.BindingDefinition.IsMultiLine)
                        {
                            ExcelInterop.Range range = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row, currentRenderingTo.Column + colId - startPos];
                            ExcelInterop.Range localSource = source[1, 1 + colId - startPos];
                            multiLineManager.CreateControl(item, range, localSource, ref vOffset);
                            ExcelApplication.ReleaseComObject(range);
                            ExcelApplication.ReleaseComObject(localSource);
                        }
                        if (item.BindingDefinition.OnAfterRendering != null)
                        {
                            ExcelInterop.Range range = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row, currentRenderingTo.Column + colId - startPos];
                            AddAfterRenderingAction(item.BindingDefinition, range);
                            ExcelApplication.ReleaseComObject(range);
                        }
                    }
                }
                renderingContext.ContextItems.Add(item);
            }
            ExcelApplication.ReleaseComObject(source);
            ExcelApplication.ReleaseComObject(workingRange);
        }
        #endregion
    }
}
