using System.Collections.Generic;
using System.Linq;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.Excel.BindingTemplates.Controls;
using Etk.Excel.BindingTemplates.Decorators;
using Etk.Excel.BindingTemplates.Definitions;
using Etk.Excel.BindingTemplates.SortSearchAndFilter;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.BindingTemplates.Renderer
{
    class ExcelPartHorizontalRenderer : ExcelPartRenderer
    {
        #region .ctors and factories
        public ExcelPartHorizontalRenderer(ExcelRenderer parent, ExcelTemplateDefinitionPart part, IBindingContextPart bindingContextPart, ExcelInterop.Range firstOutputCell, bool useDecorator)
                                          : base(parent, part, bindingContextPart, firstOutputCell, useDecorator)
        { }
        #endregion

        #region protected methods
        protected override void ManageTemplateWithoutLinkedTemplates()
        {
            ExcelInterop.Range firstCell = currentRenderingTo;
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


            int localWidth = partToRenderDefinition.Width * nbrOfElement;
            int localHeight = partToRenderDefinition.Height;
            if (nbrOfElement > 0)
            {
                ExcelInterop.Range workingRange = currentRenderingTo.Resize[localHeight, localWidth];

                partToRenderDefinition.DefinitionCells.Copy(workingRange);
                currentRenderingTo = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row + localHeight, currentRenderingTo.Column + localWidth];

                for (int rowId = 0; rowId < partToRenderDefinition.Height; rowId++)
                    Parent.ContextItems.Add(new List<IBindingContextItem>());

                foreach(IBindingContextElement contextElement in bindingContextPart.ElementsToRender)
                {
                    int cptItems = 0;
                    for (int colId = 0; colId < partToRenderDefinition.Width; colId++)
                    {
                        for (int rowId = 0; rowId < partToRenderDefinition.Height; rowId++)
                        {
                            IBindingContextItem item = partToRenderDefinition.DefinitionParts[rowId, colId] == null ? null : contextElement.BindingContextItems[cptItems++];
                            if (item != null && ((item.BindingDefinition != null && item.BindingDefinition.IsEnum && !item.BindingDefinition.IsReadOnly) || item is IExcelControl))
                            {
                                // Live Cycle of range is managed by Control
                                ExcelInterop.Range range = Parent.RootRenderer.View.ViewSheet.Cells[firstCell.Row + rowId, firstCell.Column + colId + cptElements * partToRenderDefinition.Width];
                                if (item.BindingDefinition.IsEnum )
                                    enumManager.CreateControl(item, range);
                                else
                                    ((IExcelControl)item) .CreateControl(range);
                            }
                            Parent.ContextItems[rowId].Add(item);
                        }
                    }
                    if (useDecorator && ((ExcelTemplateDefinition)this.partToRenderDefinition.Parent).Decorator != null)
                    {
                        // Live cycle of elementRange managed by ExcelElementDecorator
                        ExcelInterop.Range elementRange = firstCell.Offset[0, cptElements];
                        elementRange = elementRange.Resize[localHeight, 1];
                        Parent.RootRenderer.RowDecorators.Add(new ExcelElementDecorator(elementRange, ((ExcelTemplateDefinition)partToRenderDefinition.Parent).Decorator, contextElement));
                    }
                    cptElements++;
                }
                workingRange = null;
                //ExcelApplication.ReleaseComObject(workingRange);
            }

            // To take into account the min number of elements to render.
            if (Parent.MinOccurencesMethod != null)
            {
                IBindingContextElement parentElement = null;
                if (bindingContextPart.ParentContext != null)
                    parentElement = bindingContextPart.ParentContext.Parent;

                int minElementsToRender = LinkedTemplateDefinition.ResolveNumberOfOccurences(Parent.MinOccurencesMethod, parentElement);
                if (minElementsToRender > nbrOfElement)
                {
                    localWidth = partToRenderDefinition.Width * minElementsToRender;
                    IBindingContextItem[] toAdd = new IBindingContextItem[minElementsToRender - nbrOfElement];

                    if (Parent.ContextItems.Count == 0)
                    {
                        for (int rowId = 0; rowId < partToRenderDefinition.Height; rowId++)
                            Parent.ContextItems.Add(new List<IBindingContextItem>());
                    }
                    for (int rowId = 0; rowId < partToRenderDefinition.Height; rowId++)
                        Parent.ContextItems[rowId].AddRange(toAdd);
                }
            }

            Width += localWidth;
            if (Height < localHeight)
                Height = localHeight;
            firstCell = null;
        }

        protected override void ManageTemplateWithLinkedTemplates()
        {
            Height = partToRenderDefinition.Height;
            foreach (IBindingContextElement contextElement in bindingContextPart.ElementsToRender)
            {
                ExcelInterop.Range firstElementCell = currentRenderingTo;
                int bindingContextItemsCpt = 0;
                int cptLinkedDefinition = 0;
                int elementHeight = partToRenderDefinition.Height;
                int elementWidth = 0;
                for (int colId = 0; colId < partToRenderDefinition.Width; colId++)
                {
                    RenderingContext renderingContext = new RenderingContext(contextElement, colId);
                    List<int> posLinks = partToRenderDefinition.PositionLinkedTemplates[colId];
                    if (posLinks == null)
                    {
                        int hOffset = 1;
                        ManageTemplatePart(renderingContext, ref bindingContextItemsCpt, ref hOffset, 0, partToRenderDefinition.Height);
                        currentRenderingTo = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row, currentRenderingTo.Column + hOffset];
                        elementWidth += hOffset;

                        int newRows = Height - Parent.ContextItems.Count;
                        if (newRows > 0)
                        {
                            for (int i = 0; i < newRows; i++)
                                Parent.ContextItems.Add(new List<IBindingContextItem>(new IBindingContextItem[currentRenderingTo.Column - Parent.FirstOutputCell.Column]));
                        }

                        for (int cpt = 0; cpt < renderingContext.ContextItems.Count; cpt++)
                            Parent.ContextItems[cpt].Add(renderingContext.ContextItems[cpt]);
                    }
                    else
                    {
                        int currentWidth = 0;
                        renderingContext.PosPreviousLink = 0;
                        int lastPosLink = posLinks.Count - 1;
                        for (int linkCpt = 0; linkCpt < posLinks.Count; linkCpt++)
                        {
                            renderingContext.LinkedViewRenderedOffset = 0;
                            renderingContext.RefPos = posLinks[linkCpt];
                            renderingContext.PosCurrentLink = posLinks[linkCpt];
                            renderingContext.LinkedTemplateDefinition = partToRenderDefinition.DefinitionParts[renderingContext.PosCurrentLink, colId] as LinkedTemplateDefinition;
                            // RenderView before link
                            if (renderingContext.PosCurrentLink > 0)
                                bindingContextItemsCpt = RenderBeforeLink(renderingContext, linkCpt, bindingContextItemsCpt);

                            // RenderView link
                            IBindingContext linkedBindingContext = contextElement.LinkedBindingContexts[cptLinkedDefinition++];
                            if (linkedBindingContext.Body != null && (renderingContext.LinkedTemplateDefinition.MinOccurencesMethod != null || linkedBindingContext.Body.ElementsToRender != null && linkedBindingContext.Body.ElementsToRender.Any()))
                            {
                                RenderLink(renderingContext, linkedBindingContext);
                                renderingContext.CurrentHeight += renderingContext.LinkedViewRenderedOffset;
                            }

                            // RenderView after last link
                            if (linkCpt == lastPosLink && partToRenderDefinition.Height > renderingContext.PosCurrentLink + 1)
                                bindingContextItemsCpt = RenderAfterLink(renderingContext, bindingContextItemsCpt);

                            if (renderingContext.CurrentWidth > currentWidth)
                                currentWidth = renderingContext.CurrentWidth;
                            if (renderingContext.CurrentHeight > elementHeight)
                                elementHeight = renderingContext.CurrentHeight;
                            renderingContext.PosPreviousLink = renderingContext.PosCurrentLink;
                        }
                        elementWidth += currentWidth;
                        if (elementHeight > Height)
                            Height = elementHeight;
                        currentRenderingTo = Parent.RootRenderer.View.ViewSheet.Cells[firstRangeTo.Row, firstElementCell.Column + elementWidth];
                    }

                    int nbrOfItemToHave = Parent.Width + Width + elementWidth;
                    Parent.ContextItems.ForEach(r =>
                    {
                        if (r.Count < nbrOfItemToHave)
                            r.AddRange(new IBindingContextItem[nbrOfItemToHave - r.Count]);
                    });
                }

                if (useDecorator && ((ExcelTemplateDefinition)partToRenderDefinition.Parent).Decorator != null)
                {
                    // Live cycle of elementRange managed by ExcelElementDecorator
                    ExcelInterop.Range elementRange = firstElementCell.Resize[elementHeight, elementWidth];
                    Parent.RootRenderer.RowDecorators.Add(new ExcelElementDecorator(elementRange, ((ExcelTemplateDefinition)partToRenderDefinition.Parent).Decorator, contextElement));
                }
                Width += elementWidth;
            }
        }
        #endregion

        #region private methods
        private int RenderBeforeLink(RenderingContext renderingContext, int linkCpt, int bindingContextItemsCpt)
        {
            int firstRow, gap;
            if (linkCpt == 0)
            {
                firstRow = 0;
                gap = renderingContext.PosCurrentLink;
            }
            else
            {
                if (renderingContext.LinkedTemplateDefinition.Positioning == LinkedTemplatePositioning.Absolute)
                {
                    firstRow = renderingContext.CurrentHeight;
                    gap = renderingContext.PosCurrentLink - renderingContext.CurrentHeight;
                }
                else
                {
                    firstRow = renderingContext.PosPreviousLink + 1;
                    gap = renderingContext.PosCurrentLink - firstRow;
                }
            }
            if (gap > 0)
            {
                int newRows = renderingContext.CurrentHeight + gap - Parent.ContextItems.Count;
                if (newRows > 0)
                {
                    for (int i = 0; i < newRows; i++)
                        Parent.ContextItems.Add(new List<IBindingContextItem>(new IBindingContextItem[currentRenderingTo.Column - Parent.FirstOutputCell.Column]));
                }

                int hOffset = 1;
                ManageTemplatePart(renderingContext, ref bindingContextItemsCpt, ref hOffset, firstRow, firstRow + gap);
                currentRenderingTo = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row + gap, currentRenderingTo.Column];

                for (int cpt = 0 ; cpt < renderingContext.ContextItems.Count; cpt++)
                    Parent.ContextItems[renderingContext.CurrentHeight + cpt].Add(renderingContext.ContextItems[cpt]);

                renderingContext.CurrentHeight += gap;
                if (hOffset > renderingContext.CurrentWidth)
                    renderingContext.CurrentWidth = hOffset;
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
                renderingContext.LinkedViewRenderedOffset = linkedRenderer.Height;
                int newRows = linkedRenderer.Height + renderingContext.CurrentHeight - Parent.ContextItems.Count;
                if (newRows > 0)
                {
                    for (int i = 0; i < newRows; i++)
                        Parent.ContextItems.Add(new List<IBindingContextItem>(new IBindingContextItem[currentRenderingTo.Column - Parent.FirstOutputCell.Column]));
                }

                int rowPos = renderingContext.CurrentHeight;
                int min = linkedRenderer.ContextItems.Count < linkedRenderer.Height ? linkedRenderer.ContextItems.Count : linkedRenderer.Height;
                for (int i = 0; i < min; i++)
                    Parent.ContextItems[rowPos + i].AddRange(linkedRenderer.ContextItems[i]);

                // To take the multilines into account
                //if (linkedRenderer.Height > linkedRenderer.ContextItems.Count)
                //{
                //    for (int cpt = linkedRenderer.ContextItems.Count + 1; cpt <= linkedRenderer.Width; cpt++)
                //    {
                //        //Parent.DataRow.Add(new List<IBindingContextItem>(new IBindingContextItem[0]));
                //        List<IBindingContextItem> toUse;
                //        if (cpt >= renderingContext.CurrentWidth)
                //        {
                //            toUse = renderingContext.CurrentHeight > 0 ? new List<IBindingContextItem>(new IBindingContextItem[renderingContext.CurrentHeight])
                //                                                       : new List<IBindingContextItem>();
                //            Parent.ContextItems.Add(toUse);
                //        }
                //        else
                //        {
                //            toUse = Parent.ContextItems[cpt + renderingContext.RefPos];
                //            if (toUse.Count < renderingContext.CurrentHeight)
                //                toUse.AddRange(new IBindingContextItem[renderingContext.CurrentHeight - toUse.Count]);
                //        }
                //        toUse.AddRange(new IBindingContextItem[linkedRenderer.Height]);
                //    }
                //}

                if (linkedRenderer.Width > renderingContext.CurrentWidth)
                    renderingContext.CurrentWidth = linkedRenderer.Width;

                currentRenderingTo = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row + linkedRenderer.Height, currentRenderingTo.Column];
            }
        }

        private int RenderAfterLink(RenderingContext renderingContext, int bindingContextItemsCpt)
        {
            int hOffset = 1;
            int startPosition = renderingContext.PosCurrentLink + 1;
            if (renderingContext.LinkedTemplateDefinition.Positioning == LinkedTemplatePositioning.Absolute )//&& renderingContext.ContextElement != null && renderingContext.ContextElement.BindingContextItems.Count > renderingContext.PosCurrentLink)
            {
                int afterWrittenPosition = renderingContext.PosCurrentLink + renderingContext.LinkedViewRenderedOffset;
                for (int i = startPosition; i < afterWrittenPosition; i++)
                {
                    if (partToRenderDefinition.DefinitionParts[renderingContext.InitPos, i] != null)
                        bindingContextItemsCpt++;
                }
                startPosition = afterWrittenPosition;
            }
            if (partToRenderDefinition.Height > startPosition)
            {
                int realEnd = partToRenderDefinition.Height;
                for (int i = partToRenderDefinition.Height; i > startPosition; i--)
                {
                    ExcelInterop.Range current = partToRenderDefinition.DefinitionFirstCell.Offset[1, 0];
                    if (current.MergeCells || partToRenderDefinition.DefinitionParts[i - 1, renderingContext.InitPos] != null)
                        break;
                    realEnd--;
                }
                if (realEnd > startPosition)
                {
                    int parentDataRows = Parent.ContextItems.Count;
                    int newRows = currentRenderingTo.Row + realEnd - Parent.FirstOutputCell.Row - parentDataRows - startPosition;
                    if (newRows > 0)
                    {
                        for (int i = 0; i < newRows; i++)
                            Parent.ContextItems.Add(new List<IBindingContextItem>(new IBindingContextItem[currentRenderingTo.Column - Parent.FirstOutputCell.Column]));
                    }

                    ManageTemplatePart(renderingContext, ref bindingContextItemsCpt, ref hOffset, startPosition, realEnd);
                    for (int cpt = 0; cpt < renderingContext.ContextItems.Count; cpt++)
                        //Parent.ContextItems[parentDataRows + cpt].Add(renderingContext.ContextItems[cpt]);
                        Parent.ContextItems[renderingContext.CurrentHeight + cpt].Add(renderingContext.ContextItems[cpt]);

                    renderingContext.CurrentHeight += realEnd - startPosition;
                }
            }
            if (hOffset > renderingContext.CurrentWidth)
                renderingContext.CurrentWidth = hOffset;
            return bindingContextItemsCpt;
        }

        private void ManageTemplatePart(RenderingContext renderingContext, ref int currentBindingContextItemId, ref int hOffset, int startPos, int endPos)
        {
            renderingContext.ContextItems.Clear();

            int gap = endPos - startPos;
            ExcelInterop.Range source = Parent.RootRenderer.View.TemplateSheet.Cells[partToRenderDefinition.DefinitionFirstCell.Row + startPos, partToRenderDefinition.DefinitionFirstCell.Column + renderingContext.InitPos];

            source = source.Resize[gap, 1];
            ExcelInterop.Range workingRange = currentRenderingTo.Resize[gap, 1];
            source.Copy(workingRange);

            int bindingContextItemsCount = renderingContext.ContextElement.BindingContextItems.Count;
            for (int rowId = startPos; rowId < endPos; rowId++)
            {
                IBindingContextItem item = partToRenderDefinition.DefinitionParts[rowId, renderingContext.InitPos] == null || bindingContextItemsCount <= currentBindingContextItemId
                                           ? null
                                           : renderingContext.ContextElement.BindingContextItems[currentBindingContextItemId++];
                if (item != null)
                {
                    //if (item is ExcelBindingSearchContextItem)
                    //{
                    //    ExcelInterop.Range range = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row, currentRenderingTo.Column + colId - startPos];
                    //    ((ExcelBindingSearchContextItem)item).SetRange(ref range);
                    //    range = null;
                    //}
                    //else 
                    if (item is IExcelControl)
                    {
                        // Live Cycle of range is managed by Control
                        ExcelInterop.Range range = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row + rowId - startPos, currentRenderingTo.Column];
                        ((IExcelControl)item).CreateControl(range);
                    }
                    if (item.BindingDefinition != null)
                    {
                        if (item.BindingDefinition.IsEnum && !item.BindingDefinition.IsReadOnly)
                        {
                            // Live Cycle of range is managed by Control
                            ExcelInterop.Range range = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row + rowId - startPos, currentRenderingTo.Column];
                            enumManager.CreateControl(item, range);
                        }
                        //if (item.BindingDefinition.IsMultiLine)
                        //{
                        //    ExcelInterop.Range range = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row + rowId - startPos, currentRenderingTo.Column];
                        //    ExcelInterop.Range localSource = source[1 + rowId - startPos, 1];
                        //    multiLineManager.CreateControl(item, ref range, ref localSource, ref hOffset);
                        //    range = null;
                        //}
                        if (item.BindingDefinition.OnAfterRendering != null)
                        {
                            ExcelInterop.Range range = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row + rowId - startPos, currentRenderingTo.Column];
                            AddAfterRenderingAction(item.BindingDefinition, range);
                        }
                    }
                }
                renderingContext.ContextItems.Add(item);
            }
            source = null;
            workingRange = null;
        }
        #endregion
    }
}
