using System;
using System.Collections.Generic;
using System.Linq;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.Excel.BindingTemplates.Controls;
using Etk.Excel.BindingTemplates.Decorators;
using Etk.Excel.BindingTemplates.Definitions;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Etk.Excel.Application;

namespace Etk.Excel.BindingTemplates.Renderer
{
    class ExcelPartHorozontalRenderer : ExcelPartRenderer
    {
        #region .ctors and factories
        public ExcelPartHorozontalRenderer(ExcelRenderer parent, ExcelTemplateDefinitionPart part, IBindingContextPart bindingContextPart, ExcelInterop.Range firstOutputCell, bool useDecorator)
                                          : base(parent, part, bindingContextPart, firstOutputCell, useDecorator)
        { }
        #endregion

        #region protected methods
        protected override void ManageTemplateWithoutLinkedTemplates()
        {
            ExcelInterop.Range firstCell = currentRenderingTo;
            int cptElements = 0;

            int nbrOfElement = bindingContextPart.ElementsToRender.Count();
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
                    for (int rowId = 0; rowId < partToRenderDefinition.Height; rowId++)
                    {
                        for (int colId = 0; colId < partToRenderDefinition.Width; colId++)
                        {
                            IBindingContextItem item = partToRenderDefinition.DefinitionParts[rowId, colId] == null ? null : contextElement.BindingContextItems[cptItems++];
                            if (item != null && ((item.BindingDefinition != null && item.BindingDefinition.IsEnum && !item.BindingDefinition.IsReadOnly) || item is IExcelControl))
                            {
                                ExcelInterop.Range range = Parent.RootRenderer.View.ViewSheet.Cells[firstCell.Row + rowId, firstCell.Column + colId + cptElements * partToRenderDefinition.Width];
                                if (item.BindingDefinition.IsEnum )
                                    enumManager.CreateControl(item, ref range);
                                else
                                    ((IExcelControl)item).CreateControl(range);
                                range = null;
                            }
                            Parent.ContextItems[rowId].Add(item);
                        }
                    }
                    if (useDecorator && ((ExcelTemplateDefinition)this.partToRenderDefinition.Parent).Decorator != null)
                    {
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

                int minElementsToRender = LinkedTemplateDefinition.ResolveMinOccurences(Parent.MinOccurencesMethod, parentElement);
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
            Width = partToRenderDefinition.Width;
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
                        Parent.ContextItems.Add(renderingContext.ContextItems);
                        int hOffset = 1;
                        ManageTemplatePart(renderingContext, ref bindingContextItemsCpt, ref hOffset, 0, partToRenderDefinition.Width);
                        currentRenderingTo = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row, currentRenderingTo.Column + hOffset];
                        elementWidth += hOffset;
                    }
                    else
                    {
                        int currentWidth = 0;
                        renderingContext.RefPos = Parent.ContextItems.Count > 0 ? Parent.ContextItems.Count : 0;
                        renderingContext.PosPreviousLink = 0;
                        int lastPosLink = posLinks.Count - 1;
                        for (int linkCpt = 0; linkCpt < posLinks.Count; linkCpt++)
                        {
                            renderingContext.LinkedViewRenderedOffset = 0;
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
                            if (linkCpt == lastPosLink && renderingContext.PosCurrentLink + 1 < partToRenderDefinition.Width)
                                bindingContextItemsCpt = RenderAfterLink(renderingContext, bindingContextItemsCpt);

                            if (renderingContext.CurrentWidth > elementWidth)
                                elementWidth = renderingContext.CurrentWidth;
                            if (renderingContext.CurrentHeight > currentWidth)
                                currentWidth = renderingContext.CurrentHeight;
                            renderingContext.PosPreviousLink = renderingContext.PosCurrentLink;
                        }
                        elementWidth += currentWidth;
                        if (elementHeight > Height)
                            Height = elementHeight;
                        currentRenderingTo = Parent.RootRenderer.View.ViewSheet.Cells[firstRangeTo.Row, firstElementCell.Column + elementWidth];
                    }
                }

                if (useDecorator && ((ExcelTemplateDefinition)partToRenderDefinition.Parent).Decorator != null)
                {
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
                if (!renderingContext.RowColAdded)
                {
                    AddCol(renderingContext);
                    renderingContext.RowColAdded = true;
                }
                int hOffset = 1;
                ManageTemplatePart(renderingContext, ref bindingContextItemsCpt, ref hOffset, firstRow, firstRow + gap);
                currentRenderingTo = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row + gap, currentRenderingTo.Column];
                renderingContext.CurrentHeight += gap;
                if (hOffset > renderingContext.CurrentWidth)
                {
                    for (int i = renderingContext.CurrentWidth; i < hOffset; i++)
                        Parent.ContextItems.Add(new List<IBindingContextItem>(new IBindingContextItem[gap]));
                    renderingContext.CurrentWidth = hOffset;
                }
            }
            return bindingContextItemsCpt;
        }

        private void RenderLink(RenderingContext renderingContext, IBindingContext linkedBindingContext)
        {
            ExcelRenderer linkedRenderer = new ExcelRenderer(Parent, renderingContext.LinkedTemplateDefinition.TemplateDefinition, linkedBindingContext, currentRenderingTo,
                                                             renderingContext.LinkedTemplateDefinition.MinOccurencesMethod);
            Parent.RegisterNestedRenderer(linkedRenderer);
            linkedRenderer.Render();

            if (linkedRenderer.RenderedArea != null)
            {
                renderingContext.LinkedViewRenderedOffset = linkedRenderer.Height;
                if (!renderingContext.RowColAdded)
                {
                    AddCol(renderingContext);
                    renderingContext.RowColAdded = true;
                }

                renderingContext.ContextItems.AddRange(linkedRenderer.ContextItems[0]);
                //for (int i = 1; i < linkedRenderer.ContextItems.Count; i++)
                for (int i = 1; i < linkedRenderer.Width; i++)
                {
                    List<IBindingContextItem> toUse;
                    if (i >= renderingContext.CurrentWidth)
                    {
                        toUse = renderingContext.CurrentHeight > 0 ? new List<IBindingContextItem>(new IBindingContextItem[renderingContext.CurrentHeight])
                                                                        : new List<IBindingContextItem>();
                        Parent.ContextItems.Add(toUse);
                    }
                    else
                    {
                        toUse = Parent.ContextItems[i + renderingContext.RefPos];
                        if (toUse.Count < renderingContext.CurrentHeight)
                            toUse.AddRange(new IBindingContextItem[renderingContext.CurrentHeight - toUse.Count]);
                    }
                    toUse.AddRange(linkedRenderer.ContextItems[i]);
                }

                // To take the multilines into account
                if (linkedRenderer.Height > linkedRenderer.ContextItems.Count)
                {
                    for (int cpt = linkedRenderer.ContextItems.Count + 1; cpt <= linkedRenderer.Width; cpt++)
                    {
                        //Parent.DataRow.Add(new List<IBindingContextItem>(new IBindingContextItem[0]));
                        List<IBindingContextItem> toUse;
                        if (cpt >= renderingContext.CurrentWidth)
                        {
                            toUse = renderingContext.CurrentHeight > 0 ? new List<IBindingContextItem>(new IBindingContextItem[renderingContext.CurrentHeight])
                                                                            : new List<IBindingContextItem>();
                            Parent.ContextItems.Add(toUse);
                        }
                        else
                        {
                            toUse = Parent.ContextItems[cpt + renderingContext.RefPos];
                            if (toUse.Count < renderingContext.CurrentHeight)
                                toUse.AddRange(new IBindingContextItem[renderingContext.CurrentHeight - toUse.Count]);
                        }
                        toUse.AddRange(new IBindingContextItem[linkedRenderer.Height]);
                    }
                }

                if (renderingContext.CurrentWidth < linkedRenderer.Width)
                    renderingContext.CurrentWidth = linkedRenderer.Width;

                currentRenderingTo = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row + linkedRenderer.Height, currentRenderingTo.Column];
            }
        }

        private int RenderAfterLink(RenderingContext renderingContext, int bindingContextItemsCpt)
        {
            int vOffset = 1;
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
                //if (partToRenderDefinition.Width > startPosition + partToRenderDefinition.DefinitionFirstCell.Column)
                //{
                //    ManageTemplatePart(renderingContext, ref bindingContextItemsCpt, ref vOffset, startPosition, realEnd);
                //    renderingContext.CurrentWidth += realEnd - startPosition;
                //}
            }
            if (startPosition < partToRenderDefinition.Width)
            {
                int realEnd = partToRenderDefinition.Width;
                for (int i = partToRenderDefinition.Width - 1; i >= startPosition; i--)
                {
                    ExcelInterop.Range current = partToRenderDefinition.DefinitionFirstCell.get_Offset(0, 1);
                    if (current.MergeCells || partToRenderDefinition.DefinitionParts[renderingContext.InitPos, i] != null)
                        break;
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


        private void ManageTemplatePart(RenderingContext renderingContext, ref int currentBindingContextItemId, ref int hOffset, int startPos, int endPos)
        {
            int gap = endPos - startPos;
            //ExcelInterop.Range source = Parent.RootRenderer.View.TemplateSheet.Cells[partToRenderDefinition.DefinitionFirstCell.Row + startPos, partToRenderDefinition.DefinitionFirstCell.Column + hOffset];
            //ExcelInterop.Range source = Parent.RootRenderer.View.TemplateSheet.Cells[partToRenderDefinition.DefinitionFirstCell.Row + renderingContext.InitPos, partToRenderDefinition.DefinitionFirstCell.Column + startPos];
            ExcelInterop.Range source = Parent.RootRenderer.View.TemplateSheet.Cells[partToRenderDefinition.DefinitionFirstCell.Row + startPos, partToRenderDefinition.DefinitionFirstCell.Column + renderingContext.InitPos];


            source = source.Resize[gap, 1];
            ExcelInterop.Range workingRange = currentRenderingTo.Resize[gap, 1];
            source.Copy(workingRange);

            //for (int rowId = startPos; rowId < endPos; rowId++)
            //{
            //    IBindingContextItem item = partToRenderDefinition.DefinitionParts[colId, rowId] == null ? null : contextElement.BindingContextItems[cpt++];
            //    if (item != null && ((item.BindingDefinition != null && item.BindingDefinition.IsEnum) || item is IExcelControl))
            //    {
            //        ExcelInterop.Range range = Parent.RootRenderer.View.ViewSheet.Cells[currentRenderingTo.Row + rowId, currentRenderingTo.Column];
            //        if (item.BindingDefinition.IsEnum && !item.BindingDefinition.IsReadOnly)
            //            enumManager.CreateControl(item, ref range);
            //        else
            //            ((IExcelControl) item).CreateControl(range);
            //        range = null;
            //    }
            //    col.Add(item);
            //}

            //ExcelApplication.ReleaseComObject(source);
            //ExcelApplication.ReleaseComObject(workingRange);
            source = null;
            workingRange = null;
        }

        // To redo !!!!
        private void AddCol(RenderingContext renderingContext)
        {
            Parent.ContextItems.Add(renderingContext.ContextItems);
            renderingContext.CurrentWidth = 1;
        }
        #endregion
    }
}
