namespace Etk.Excel.BindingTemplates.Renderer
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;
    using Etk.BindingTemplates.Context;
    using Etk.BindingTemplates.Definitions.Templates;
    using Etk.Excel.BindingTemplates.Controls;
    using Etk.Excel.BindingTemplates.Definitions;
    using Microsoft.Office.Interop.Excel;
    using Views;

    class ExcelPartRenderer : IDisposable
    {
        #region attributes and properties
        private static EnumManager enumManager = new EnumManager();
        private static MultiLineManager multiLineManager = new MultiLineManager();

        private bool useDecorator;

        private ExcelRenderer Parent;
        private ExcelTemplateDefinitionPart partToRenderDefinition;
        private IBindingContextPart bindingContextPart;

        private Range firstRangeTo;
        private Range elementFirstRangeTo;

        private Range currentRenderingFrom;
        private Range currentRenderingTo;

        public int Height
        { get; private set; }

        public int Width
        { get; private set; }

        //public Range RenderedRange
        //{ get; private set; }

        public RenderedArea RenderedArea
        { get; private set; }

        public bool isExpander = false;
        #endregion

        #region .ctors and factories
        public ExcelPartRenderer(ExcelRenderer parent, ExcelTemplateDefinitionPart part, IBindingContextPart bindingContextPart, Range firstOutputCell, bool useDecorator)
        {
            this.Parent = parent;
            this.partToRenderDefinition = part;
            this.bindingContextPart = bindingContextPart;
            this.useDecorator = useDecorator;

            currentRenderingFrom = partToRenderDefinition.DefinitionFirstCell;
            firstRangeTo = elementFirstRangeTo = currentRenderingTo = firstOutputCell;

            Init();
        }
        #endregion

        #region public methods
        public void Render()
        {
            Worksheet worksheetTo = currentRenderingTo.Worksheet;
            if (bindingContextPart != null && bindingContextPart.ElementsToRender != null && bindingContextPart.ElementsToRender.Any())
            {
                if (partToRenderDefinition.HasLinkedTemplates || partToRenderDefinition.ContainMultiLinesCells)
                {
                    if (partToRenderDefinition.Parent.Orientation == Orientation.Vertical)
                        ManageVerticalTemplateWithLinkedTemplates();
                    else
                        ManageHorizontalTemplateWithLinkedTemplates();
                }
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

        #region private method
        private void Init()
        {
            Height = Width = 0;
        }
        #endregion

        #region private methods
        private void ManageTemplateWithoutLinkedTemplates()
        { 
            Range workingRange;
            Range firstCell = currentRenderingTo;
            Worksheet worksheetTo = currentRenderingTo.Worksheet;
            int cptItems = 0;
            int cptElements = 0;

            if (partToRenderDefinition.Parent.Orientation == Orientation.Vertical)
            {
                int nbrOfElement = bindingContextPart.ElementsToRender.Count();
                int localWidth = partToRenderDefinition.Width;
                int localHeight = partToRenderDefinition.Height * nbrOfElement;
                workingRange = currentRenderingTo.Resize[localHeight, localWidth];

                partToRenderDefinition.DefinitionCells.Copy(workingRange);
                currentRenderingTo = worksheetTo.Cells[currentRenderingTo.Row + localHeight, currentRenderingTo.Column + localWidth];

                foreach (IBindingContextElement contextElement in bindingContextPart.ElementsToRender)
                {
                    cptItems = 0;
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
                                if (item.BindingDefinition.IsEnum)
                                    enumManager.CreateControl(item, ref range);
                                else
                                    ManageControls(item, ref range);
                                range = null;
                            }
                            row.Add(item);
                        }
                    }
                    if (useDecorator && ((ExcelTemplateDefinition)this.partToRenderDefinition.Parent).Decorator != null)
                    {
                        Range elementRange = firstCell.Offset[cptElements, 0];
                        elementRange = elementRange.Resize[1, localWidth];
                        ((ExcelTemplateDefinition)this.partToRenderDefinition.Parent).Decorator.Resolve(elementRange, contextElement);
                    }
                    cptElements++;
                }
                Height += localHeight;
                if (Width < localWidth)
                    Width = localWidth;
            }
            else
            {
                int localWidth = partToRenderDefinition.Width * bindingContextPart.ElementsToRender.Count();
                int localHeight = partToRenderDefinition.Height;
                workingRange = currentRenderingTo.Resize[localHeight, localWidth];

                partToRenderDefinition.DefinitionCells.Copy(workingRange);
                currentRenderingTo = worksheetTo.Cells[currentRenderingTo.Row + localHeight, currentRenderingTo.Column + localWidth];

                for (int rowId = 0; rowId < partToRenderDefinition.Height; rowId++)
                    Parent.DataRows.Add(new List<IBindingContextItem>());

                foreach(IBindingContextElement contextElement in bindingContextPart.ElementsToRender)
                {
                    cptItems = 0;
                    for (int rowId = 0; rowId < partToRenderDefinition.Height; rowId++)
                    {
                        for (int colId = 0; colId < partToRenderDefinition.Width; colId++)
                        {
                            IBindingContextItem item = partToRenderDefinition.DefinitionParts[rowId, colId] == null ? null : contextElement.BindingContextItems[cptItems++];
                            if (item != null && ((item.BindingDefinition != null && item.BindingDefinition.IsEnum && !item.BindingDefinition.IsReadOnly) || item is IExcelControl))
                            {
                                Range range = worksheetTo.Cells[firstCell.Row + rowId, firstCell.Column + colId + cptElements * partToRenderDefinition.Width];
                                if (item.BindingDefinition.IsEnum )
                                    enumManager.CreateControl(item, ref range);
                                else
                                    ManageControls(item, ref range);

                                range = null;
                            }
                            Parent.DataRows[rowId].Add(item);
                        }
                    }
                    if (useDecorator && ((ExcelTemplateDefinition)this.partToRenderDefinition.Parent).Decorator != null)
                    {
                        Range elementRange = firstCell.Offset[0, cptElements];
                        elementRange = elementRange.Resize[localHeight, 1];
                        ((ExcelTemplateDefinition)this.partToRenderDefinition.Parent).Decorator.Resolve(elementRange, contextElement);
                    }
                    cptElements++;
                }
                Width += localWidth;
                if (Height < localHeight)
                    Height = localHeight;
            }

            Marshal.ReleaseComObject(worksheetTo);
            Marshal.ReleaseComObject(workingRange);
            firstCell = null;
        }

        private void ManageVerticalTemplateWithLinkedTemplates()
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
                    List<IBindingContextItem> dataRow = new List<IBindingContextItem>();
                    List<int> posLinks = partToRenderDefinition.PositionLinkedTemplates[rowId];
                    if (posLinks == null)
                    {
                        Parent.DataRows.Add(dataRow);
                        int vOffset = 1;
                        ManageVerticalTemplatePart(ref bindingContextItemsCpt, ref vOffset, contextElement, dataRow, rowId, 0, partToRenderDefinition.Width);
                        currentRenderingTo = worksheetTo.Cells[currentRenderingTo.Row + vOffset, currentRenderingTo.Column];
                        Height += vOffset;
                    }
                    else
                    {
                        int refRow = Parent.DataRows.Count > 0 ? Parent.DataRows.Count : 0;
                        int currentRowHeight = 0;
                        int currentRowWidth = 0;
                        int posPreviousLink = 0;
                        bool rowAdded = false;
                        int lastPosLink = posLinks.Count - 1;
                        for (int linkCpt = 0; linkCpt < posLinks.Count; linkCpt++)
                        {
                            int linkedViewRenderedWidth = 0;
                            int posCurrentLink = posLinks[linkCpt];
                            LinkedTemplateDefinition linkedTemplateDefinition = partToRenderDefinition.DefinitionParts[rowId, posCurrentLink] as LinkedTemplateDefinition;
                            // Render before link
                            if (posCurrentLink > 0)
                            {
                                int firstCol, gap;
                                if (linkCpt == 0)
                                {
                                    firstCol = 0;
                                    gap = posCurrentLink;
                                }
                                else
                                {
                                    if (linkedTemplateDefinition.Positioning == LinkedTemplatePositioning.Absolute)
                                    {
                                        firstCol = currentRowWidth;
                                        gap = posCurrentLink - currentRowWidth;
                                    }
                                    else
                                    {
                                        firstCol = posPreviousLink + 1;
                                        gap = posCurrentLink - firstCol;
                                    }
                                }
                                if (gap > 0)
                                {
                                    if (!rowAdded)
                                        AddRow(ref currentRowHeight, ref rowAdded, dataRow);
                                    int vOffset = 1;
                                    ManageVerticalTemplatePart(ref bindingContextItemsCpt, ref vOffset, contextElement, dataRow, rowId, firstCol, firstCol + gap);
                                    currentRenderingTo = worksheetTo.Cells[currentRenderingTo.Row, currentRenderingTo.Column + gap];
                                    currentRowWidth += gap;
                                    if (vOffset > currentRowHeight)
                                        currentRowHeight = vOffset;
                                }
                            }
                            // Render link
                            IBindingContext linkedBindingContext = (IBindingContext)contextElement.LinkedBindingContexts[cptLinkedDefinition++];
                            if (linkedBindingContext.Body != null && linkedBindingContext.Body.ElementsToRender != null && linkedBindingContext.Body.ElementsToRender.Any())
                            {
                                using (ExcelRenderer linkedRenderer = new ExcelRenderer(linkedTemplateDefinition.TemplateDefinition, linkedBindingContext, currentRenderingTo))
                                {
                                    linkedRenderer.Render();

                                    if (linkedRenderer.RenderedArea != null)
                                    {
                                        linkedViewRenderedWidth = linkedRenderer.Width;
                                        if (!rowAdded)
                                            AddRow(ref currentRowHeight, ref rowAdded, dataRow);
                                        dataRow.AddRange(linkedRenderer.DataRows[0]);

                                        for (int i = 1; i < linkedRenderer.DataRows.Count; i++)
                                        {
                                            List<IBindingContextItem> rowToUse;
                                            if (i >= currentRowHeight)
                                            {
                                                rowToUse = currentRowWidth > 0 ? new List<IBindingContextItem>(new IBindingContextItem[currentRowWidth]) : new List<IBindingContextItem>();
                                                Parent.DataRows.Add(rowToUse);
                                            }
                                            else
                                            {
                                                rowToUse = Parent.DataRows[i + refRow];
                                                if (rowToUse.Count < currentRowWidth)
                                                    rowToUse.AddRange(new IBindingContextItem[currentRowWidth - rowToUse.Count]);
                                            }
                                            rowToUse.AddRange(linkedRenderer.DataRows[i]);
                                        }

                                        // To take the multilines into account
                                        if (linkedRenderer.Height > linkedRenderer.DataRows.Count)
                                        {
                                            for (int cpt = linkedRenderer.DataRows.Count + 1; cpt <= linkedRenderer.Height; cpt++)
                                            {
                                                //Parent.DataRows.Add(new List<IBindingContextItem>(new IBindingContextItem[0]));
                                                List<IBindingContextItem> rowToUse;
                                                if (cpt >= currentRowHeight)
                                                {
                                                    rowToUse = currentRowWidth > 0 ? new List<IBindingContextItem>(new IBindingContextItem[currentRowWidth]) : new List<IBindingContextItem>();
                                                    Parent.DataRows.Add(rowToUse);
                                                }
                                                else
                                                {
                                                    rowToUse = Parent.DataRows[cpt + refRow];
                                                    if (rowToUse.Count < currentRowWidth)
                                                        rowToUse.AddRange(new IBindingContextItem[currentRowWidth - rowToUse.Count]);
                                                }
                                                rowToUse.AddRange(new IBindingContextItem[linkedRenderer.Width]);
                                            }
                                        }

                                        if (currentRowHeight < linkedRenderer.Height)
                                            currentRowHeight = linkedRenderer.Height;

                                        currentRenderingTo = worksheetTo.Cells[currentRenderingTo.Row, currentRenderingTo.Column + linkedRenderer.Width];
                                    }
                                }
                            }
                            currentRowWidth += linkedViewRenderedWidth;

                            // Render after link
                            if (linkCpt == lastPosLink && posCurrentLink != partToRenderDefinition.Width)
                            {
                                int vOffset = 1;
                                if (linkedTemplateDefinition.Positioning == LinkedTemplatePositioning.Absolute)
                                {
                                    int startPosition = posCurrentLink + linkedViewRenderedWidth;
                                    if (startPosition < partToRenderDefinition.Width)
                                    {
                                        ManageVerticalTemplatePart(ref bindingContextItemsCpt, ref vOffset, contextElement, dataRow, rowId, startPosition, partToRenderDefinition.Width);
                                        currentRowWidth += partToRenderDefinition.Width - startPosition;
                                    }
                                }
                                else
                                {
                                    int startPosition = posCurrentLink + 1;
                                    if (startPosition < partToRenderDefinition.Width)
                                    {
                                        int realEnd = partToRenderDefinition.Width;
                                        for (int i = partToRenderDefinition.Width - 1; i >= startPosition; i--)
                                        {
                                            if (contextElement.BindingContextItems[i] != null)
                                                break;
                                            realEnd--;
                                        }
                                        if (realEnd > 0)
                                        {
                                            ManageVerticalTemplatePart(ref bindingContextItemsCpt, ref vOffset, contextElement, dataRow, rowId, startPosition, realEnd);
                                            currentRowWidth += partToRenderDefinition.Width - startPosition;
                                        }
                                    }
                                }
                                if (vOffset > currentRowHeight)
                                    currentRowHeight = vOffset;
                            }
                            // End Render after link

                            if (currentRowWidth > elementWidth)
                                elementWidth = currentRowWidth;
                            if (currentRowHeight > elementHeight)
                                elementHeight = currentRowHeight;
                            posPreviousLink = posCurrentLink;
                        }
                        if (elementWidth > Width)
                            Width = elementWidth;
                        Height += elementHeight;
                        currentRenderingTo = worksheetTo.Cells[firstRangeTo.Row + Height, firstRangeTo.Column];
                    }
                }

                if (useDecorator && ((ExcelTemplateDefinition)this.partToRenderDefinition.Parent).Decorator != null)
                {
                    Range elementRange = firstElementCell.Resize[elementHeight, elementWidth];
                    ((ExcelTemplateDefinition)this.partToRenderDefinition.Parent).Decorator.Resolve(elementRange, contextElement);
                }
            }
            Marshal.ReleaseComObject(worksheetTo);
            worksheetTo = null;
        }

        private void AddRow(ref int currentRowHeight, ref bool rowAdded, List<IBindingContextItem> dataRow)
        {
            Parent.DataRows.Add(dataRow);
            currentRowHeight = 1;
            rowAdded = true;
        }

        private void ManageVerticalTemplatePart(ref int currentBindingContextItemId, ref int vOffset, IBindingContextElement contextElement, List<IBindingContextItem> row, int rowId, int startPos, int endPos)
        {
            Worksheet worksheetFrom = partToRenderDefinition.DefinitionFirstCell.Worksheet;
            Worksheet worksheetTo = currentRenderingTo.Worksheet;

            int gap = endPos - startPos;
            Range source = worksheetFrom.Cells[partToRenderDefinition.DefinitionFirstCell.Row + rowId, partToRenderDefinition.DefinitionFirstCell.Column + startPos];
            source = source.Resize[1, gap];
            Range workingRange = currentRenderingTo.Resize[1, gap];
            source.Copy(workingRange);

            for (int colId = startPos; colId < endPos; colId++)
            {
                IBindingContextItem item = partToRenderDefinition.DefinitionParts[rowId, colId] == null ? null : contextElement.BindingContextItems[currentBindingContextItemId++];
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
                row.Add(item);
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

        // To redo !!!!
        private void ManageHorizontalTemplateWithLinkedTemplates()
        {
            Worksheet worksheetTo = currentRenderingTo.Worksheet;
            foreach (IBindingContextElement contextElement in bindingContextPart.ElementsToRender)
            {
                int bindingContextItemsCpt = 0;
                int cptLinkedDefinition = 0;
                int elementHeight = 0;
                int elementWidth = partToRenderDefinition.Width;
                for (int rowId = 0; rowId < partToRenderDefinition.Height; rowId++)
                {
                    List<IBindingContextItem> dataRow = new List<IBindingContextItem>();
                    List<int> posLinks = partToRenderDefinition.PositionLinkedTemplates[rowId];
                    if (posLinks == null)
                    {
                        Height++;
                        Parent.DataRows.Add(dataRow);
                        ManageHorizontalTemplatePart(ref bindingContextItemsCpt, contextElement, dataRow, rowId, 0, partToRenderDefinition.Width);
                        currentRenderingTo = worksheetTo.Cells[currentRenderingTo.Row + 1, currentRenderingTo.Column];
                    }
                    else
                    {
                        int refRow = Parent.DataRows.Count > 0 ? Parent.DataRows.Count : 0;
                        int currentRowHeight = 0;
                        int currentRowWidth = 0;
                        int posPreviousLink = 0;
                        bool rowAdded = false;
                        for (int linkCpt = 0; linkCpt < posLinks.Count; linkCpt++)
                        {
                            int linkedViewRenderedWidth = 0;
                            int posCurrentLink = posLinks[linkCpt];
                            LinkedTemplateDefinition linkedTemplateDefinition = partToRenderDefinition.DefinitionParts[rowId, posCurrentLink] as LinkedTemplateDefinition;
                            // Render before link
                            if (posCurrentLink > 0)
                            {
                                int firstCol, gap;
                                if (linkCpt == 0)
                                {
                                    firstCol = 0;
                                    gap = posCurrentLink;
                                }
                                else
                                {
                                    if (linkedTemplateDefinition.Positioning == LinkedTemplatePositioning.Absolute)
                                    {
                                        firstCol = currentRowWidth;
                                        gap = posCurrentLink - currentRowWidth;
                                    }
                                    else
                                    {
                                        firstCol = posPreviousLink + 1;
                                        gap = posCurrentLink - firstCol;
                                    }
                                }
                                if (gap > 0)
                                {
                                    if (!rowAdded)
                                        AddRow(ref currentRowHeight, ref rowAdded, dataRow);
                                    ManageHorizontalTemplatePart(ref bindingContextItemsCpt, contextElement, dataRow, rowId, firstCol, firstCol + gap);
                                    currentRenderingTo = worksheetTo.Cells[currentRenderingTo.Row, currentRenderingTo.Column + gap];
                                    currentRowWidth += gap;
                                }
                            }
                            // Render link
                            IBindingContext linkedBindingContext = (IBindingContext)contextElement.LinkedBindingContexts[cptLinkedDefinition++];
                            if (linkedBindingContext.Body != null && linkedBindingContext.Body.ElementsToRender != null && linkedBindingContext.Body.ElementsToRender.Any())
                            {
                                using (ExcelRenderer linkedRenderer = new ExcelRenderer(linkedTemplateDefinition.TemplateDefinition, linkedBindingContext, currentRenderingTo))
                                {
                                    linkedRenderer.Render();

                                    if (linkedRenderer.RenderedArea != null)
                                    {
                                        linkedViewRenderedWidth = linkedRenderer.Width;
                                        if (!rowAdded)
                                            AddRow(ref currentRowHeight, ref rowAdded, dataRow);
                                        dataRow.AddRange(linkedRenderer.DataRows[0]);

                                        for (int i = 1; i < linkedRenderer.Height; i++)
                                        {
                                            List<IBindingContextItem> rowToUse;
                                            if (i >= currentRowHeight)
                                            {
                                                rowToUse = currentRowWidth > 0 ? new List<IBindingContextItem>(new IBindingContextItem[currentRowWidth]) : new List<IBindingContextItem>();
                                                Parent.DataRows.Add(rowToUse);
                                            }
                                            else
                                            {
                                                rowToUse = Parent.DataRows[i + refRow];
                                                if (rowToUse.Count < currentRowWidth)
                                                    rowToUse.AddRange(new IBindingContextItem[currentRowWidth - rowToUse.Count]);
                                            }
                                            rowToUse.AddRange(linkedRenderer.DataRows[i]);
                                        }

                                        if (currentRowHeight < linkedRenderer.Height)
                                            currentRowHeight = linkedRenderer.Height;

                                        currentRenderingTo = worksheetTo.Cells[currentRenderingTo.Row, currentRenderingTo.Column + linkedRenderer.Width];
                                    }
                                }
                            }
                            currentRowWidth += linkedViewRenderedWidth;

                            // Render after link
                            // ToDo

                            if (currentRowWidth > elementWidth)
                                elementWidth = currentRowWidth;
                            if (currentRowHeight > elementHeight)
                                elementHeight = currentRowHeight;
                            posPreviousLink = posCurrentLink;
                        }
                        if (elementWidth > Width)
                            Width = elementWidth;
                        Height += elementHeight;
                        currentRenderingTo = worksheetTo.Cells[firstRangeTo.Row + Height, firstRangeTo.Column];
                    }
                }
            }
            Marshal.ReleaseComObject(worksheetTo);
            worksheetTo = null;
        }

        // To redo !!!!
        private void AddCol(ref int currentColWidth, ref bool colAdded, List<IBindingContextItem> dataRow)
        {
            Parent.DataRows.Add(dataRow);
            currentColWidth = 1;
            colAdded = true;
        }

        // To redo !!!!
        private void ManageHorizontalTemplatePart(ref int cpt, IBindingContextElement contextElement, List<IBindingContextItem> col, int colId, int startPos, int endPos)
        {
            Worksheet worksheetFrom = partToRenderDefinition.DefinitionFirstCell.Worksheet;
            Worksheet worksheetTo = currentRenderingTo.Worksheet;

            int gap = endPos - startPos;
            Range source = worksheetFrom.Cells[partToRenderDefinition.DefinitionFirstCell.Row + startPos, partToRenderDefinition.DefinitionFirstCell.Column + colId];
            source = source.Resize[gap, 1];
            Range workingRange = currentRenderingTo.Resize[gap, 1];
            source.Copy(workingRange);

            for (int rowId = startPos; rowId < endPos; rowId++)
            {
                IBindingContextItem item = partToRenderDefinition.DefinitionParts[colId, rowId] == null ? null : contextElement.BindingContextItems[cpt++];
                if (item != null && ((item.BindingDefinition != null && item.BindingDefinition.IsEnum) || item is IExcelControl))
                {
                    Range range = worksheetTo.Cells[currentRenderingTo.Row + rowId, currentRenderingTo.Column];
                    if (item.BindingDefinition.IsEnum && !item.BindingDefinition.IsReadOnly)
                        enumManager.CreateControl(item, ref range);
                    else
                        ManageControls(item, ref range);
                    range = null;
                }
                col.Add(item);
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

        private void ManageControls(IBindingContextItem item, ref Range range)
        {
            if (item is IExcelControl)
                ((IExcelControl)item).CreateControl(range);
        }
        #endregion
    }
}
