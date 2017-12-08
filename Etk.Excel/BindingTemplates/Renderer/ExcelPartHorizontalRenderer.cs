using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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

        #region private methods
        protected override void ManageTemplateWithoutLinkedTemplates()
        {
            ExcelInterop.Range firstCell = currentRenderingTo;
            ExcelInterop.Worksheet worksheetTo = currentRenderingTo.Worksheet;
            int cptElements = 0;

            int nbrOfElement = bindingContextPart.ElementsToRender.Count();
            int localWidth = partToRenderDefinition.Width * nbrOfElement;
            int localHeight = partToRenderDefinition.Height;
            if (nbrOfElement > 0)
            {
                ExcelInterop.Range workingRange = currentRenderingTo.Resize[localHeight, localWidth];

                partToRenderDefinition.DefinitionCells.Copy(workingRange);
                currentRenderingTo = worksheetTo.Cells[currentRenderingTo.Row + localHeight, currentRenderingTo.Column + localWidth];

                for (int rowId = 0; rowId < partToRenderDefinition.Height; rowId++)
                    Parent.DataRows.Add(new List<IBindingContextItem>());

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
                                ExcelInterop.Range range = worksheetTo.Cells[firstCell.Row + rowId, firstCell.Column + colId + cptElements * partToRenderDefinition.Width];
                                if (item.BindingDefinition.IsEnum )
                                    enumManager.CreateControl(item, ref range);
                                else
                                    ((IExcelControl)item).CreateControl(range);
                                range = null;
                            }
                            Parent.DataRows[rowId].Add(item);
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

                    if (Parent.DataRows.Count == 0)
                    {
                        for (int rowId = 0; rowId < partToRenderDefinition.Height; rowId++)
                            Parent.DataRows.Add(new List<IBindingContextItem>());
                    }
                    for (int rowId = 0; rowId < partToRenderDefinition.Height; rowId++)
                        Parent.DataRows[rowId].AddRange(toAdd);
                }
            }

            Width += localWidth;
            if (Height < localHeight)
                Height = localHeight;

            ExcelApplication.ReleaseComObject(worksheetTo);
            firstCell = null;
        }

        // To redo !!!!
        protected override void ManageTemplateWithLinkedTemplates()
        {
            throw new NotImplementedException("Manage horizontal templates with linked templates are not supported yet");
        }

        // To redo !!!!
        private void AddCol(ref int currentColWidth, ref bool colAdded, List<IBindingContextItem> dataRow)
        {
            Parent.DataRows.Add(dataRow);
            currentColWidth = 1;
            colAdded = true;
        }

        // To redo !!!!
        private void ManageTemplatePart(ref int cpt, IBindingContextElement contextElement, List<IBindingContextItem> col, int colId, int startPos, int endPos)
        {
            ExcelInterop.Worksheet worksheetFrom = partToRenderDefinition.DefinitionFirstCell.Worksheet;
            ExcelInterop.Worksheet worksheetTo = currentRenderingTo.Worksheet;

            int gap = endPos - startPos;
            ExcelInterop.Range source = worksheetFrom.Cells[partToRenderDefinition.DefinitionFirstCell.Row + startPos, partToRenderDefinition.DefinitionFirstCell.Column + colId];
            source = source.Resize[gap, 1];
            ExcelInterop.Range workingRange = currentRenderingTo.Resize[gap, 1];
            source.Copy(workingRange);

            for (int rowId = startPos; rowId < endPos; rowId++)
            {
                IBindingContextItem item = partToRenderDefinition.DefinitionParts[colId, rowId] == null ? null : contextElement.BindingContextItems[cpt++];
                if (item != null && ((item.BindingDefinition != null && item.BindingDefinition.IsEnum) || item is IExcelControl))
                {
                    ExcelInterop.Range range = worksheetTo.Cells[currentRenderingTo.Row + rowId, currentRenderingTo.Column];
                    if (item.BindingDefinition.IsEnum && !item.BindingDefinition.IsReadOnly)
                        enumManager.CreateControl(item, ref range);
                    else
                        ((IExcelControl) item).CreateControl(range);
                    range = null;
                }
                col.Add(item);
            }

            ExcelApplication.ReleaseComObject(worksheetFrom);
            ExcelApplication.ReleaseComObject(worksheetTo);
            worksheetFrom = null;
            worksheetTo = null;

            //ExcelApplication.ReleaseComObject(source);
            //ExcelApplication.ReleaseComObject(workingRange);
            source = null;
            workingRange = null;
        }
        #endregion
    }
}
