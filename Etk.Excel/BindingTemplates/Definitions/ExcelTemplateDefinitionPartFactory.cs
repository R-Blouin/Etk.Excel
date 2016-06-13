using System;
using System.Collections.Generic;
using Etk.BindingTemplates.Definitions;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.BindingTemplates.Definitions.Templates.Xml;
using Etk.Excel.BindingTemplates.Controls.Button;
using Etk.Excel.BindingTemplates.Controls.CheckBox;
using Etk.Excel.BindingTemplates.Controls.FormulaResult;
using Etk.Excel.BindingTemplates.Controls.NamedRange;
using Etk.Excel.BindingTemplates.SortSearchAndFilter;
using Etk.Tools.Extensions;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.BindingTemplates.Definitions
{
    /// <summary> ExcelTemplateDefinitionPart factory</summary>
    class ExcelTemplateDefinitionPartFactory
    {
        #region const
        private const string LINKED_TEMPLATE_PREFIX = "<Link";
        #endregion

        #region .ctors
        private ExcelTemplateDefinitionPartFactory()
        {}
        #endregion

        #region public method
        public static ExcelTemplateDefinitionPart CreateInstance(ExcelTemplateDefinition excelTemplateDefinition, ExcelInterop.Range firstRange, ExcelInterop.Range lastRange)
        {
            ExcelTemplateDefinitionPartFactory factory = new ExcelTemplateDefinitionPartFactory();
            return factory.Execute(excelTemplateDefinition, firstRange, lastRange);
        }
        #endregion

        private ExcelTemplateDefinitionPart Execute(ExcelTemplateDefinition excelTemplateDefinition, ExcelInterop.Range firstRange, ExcelInterop.Range lastRange)
        {
            ExcelTemplateDefinitionPart part = new ExcelTemplateDefinitionPart(excelTemplateDefinition, firstRange, lastRange);
            for (int rowId = 0; rowId < part.DefinitionCells.Rows.Count; rowId++)
            {
                List<int> posLinks = null;
                ExcelInterop.Range row = part.DefinitionCells.Rows[rowId + 1];

                for (int cellId = 0; cellId < row.Cells.Count; cellId++)
                {
                    ExcelInterop.Range cell = row.Cells[cellId + 1];
                    IDefinitionPart definitionPart = AnalyzeCell(part, cell);
                    part.DefinitionParts[rowId, cellId] = definitionPart;

                    if (definitionPart is LinkedTemplateDefinition)
                    {
                        if (posLinks == null)
                            posLinks = new List<int>();
                        posLinks.Add(cellId);
                    }

                    if (definitionPart is IBindingDefinition && ((IBindingDefinition)definitionPart).IsMultiLine)
                        part.ContainMultiLinesCells = true;
                }
                part.PositionLinkedTemplates.Add(posLinks);
            }

            return part;
        }

        /// <summary>Analyze a cell of the template part</summary>
        private IDefinitionPart AnalyzeCell(ExcelTemplateDefinitionPart templateDefinitionPart, ExcelInterop.Range cell)
        {
            IDefinitionPart part = null;
            if (cell.Value2 != null)
            {
                string value = cell.Value2.ToString();
                if (! string.IsNullOrEmpty(value))
                {
                    string trimmedValue = value.Trim();
                    if (trimmedValue.StartsWith(LINKED_TEMPLATE_PREFIX))
                    {
                        try
                        {
                            XmlTemplateLink xmlTemplateLink = trimmedValue.Deserialize<XmlTemplateLink>();
                            TemplateLink templateLink = TemplateLink.CreateInstance(xmlTemplateLink);

                            ExcelTemplateDefinition templateDefinition = (ETKExcel.TemplateManager as ExcelTemplateManager).GetTemplateDefinitionFromLink(templateDefinitionPart, templateLink);
                            LinkedTemplateDefinition linkedTemplateDefinition = new LinkedTemplateDefinition(templateDefinitionPart.Parent, templateDefinition, templateLink);
                            templateDefinitionPart.AddLinkedTemplate(linkedTemplateDefinition);
                            part = linkedTemplateDefinition;
                        }
                        catch (Exception ex)
                        {
                            string message = string.Format("Cannot create the linked template dataAccessor '{0}'. {1}", trimmedValue, ex.Message);
                            throw new EtkException(message, false);
                        }
                    }
                    else
                    {
                        if (trimmedValue.StartsWith(ExcelBindingFilterDefinition.Filter_PREFIX))
                        {
                            ExcelBindingFilterDefinition filterdefinition = ExcelBindingFilterDefinition.CreateInstance(templateDefinitionPart, trimmedValue);
                            templateDefinitionPart.AddFilterDefinition(filterdefinition);
                            part = filterdefinition;
                        }
                        else if (trimmedValue.StartsWith(ExcelBindingSearchDefinition.Search_PREFIX))
                        {
                            ExcelBindingSearchDefinition searchdefinition = ExcelBindingSearchDefinition.CreateInstance(trimmedValue);
                            templateDefinitionPart.AddSearchDefinition(searchdefinition);
                            part = searchdefinition;
                        }
                        else
                        {
                            try
                            {
                                IBindingDefinition bindingDefinition = CreateBindingDefinition(templateDefinitionPart, value, trimmedValue);
                                templateDefinitionPart.AddBindingDefinition(bindingDefinition);
                                part = bindingDefinition;
                            }
                            catch (Exception ex)
                            {
                                string message = string.Format("Cannot create the binding definition for '{0}'. {1}", trimmedValue, ex.Message);
                                throw new EtkException(message, false);
                            }
                        }
                    }
                }
            }
            return part;
        }

        /// <summary>Create a binding definition from a cell value</summary>
        private IBindingDefinition CreateBindingDefinition(ExcelTemplateDefinitionPart templateDefinitionPart, string value, string trimmedValue)
        {
            IBindingDefinition ret = null;
            if (trimmedValue.StartsWith(ExcelBindingDefinitionButton.BUTTON_TEMPLATE_PREFIX))
                ret = ExcelBindingDefinitionButton.CreateInstance(templateDefinitionPart.Parent as ExcelTemplateDefinition, trimmedValue);
            else if (trimmedValue.StartsWith(ExcelBindingDefinitionCheckBox.CHECKBOX_TEMPLATE_PREFIX))
                ret = ExcelBindingDefinitionCheckBox.CreateInstance(templateDefinitionPart.Parent as ExcelTemplateDefinition, trimmedValue);
            else if (trimmedValue.StartsWith(ExcelBindingDefinitionFormulaResult.FORMULA_RESULT_PREFIX))
                ret = ExcelBindingDefinitionFormulaResult.CreateInstance(templateDefinitionPart.Parent as ExcelTemplateDefinition, trimmedValue);
            else if (trimmedValue.StartsWith(ExcelBindingDefinitionNamedRange.NAMEDRANGE_TEMPLATE_PREFIX))
            {
                ExcelNamedRangeDefinition excelTemplateDefinition = ExcelBindingDefinitionNamedRange.RetrieveNamedRangeDefinition(trimmedValue);
                if (excelTemplateDefinition != null)
                {
                    BindingDefinition nestedBindingDefinition = null;
                    if (!string.IsNullOrEmpty(excelTemplateDefinition.Value))
                        nestedBindingDefinition = CreateBindingDefinition(templateDefinitionPart, excelTemplateDefinition.Value, excelTemplateDefinition.Value.Trim()) as BindingDefinition;
                    ret = ExcelBindingDefinitionNamedRange.CreateInstance(templateDefinitionPart, excelTemplateDefinition, nestedBindingDefinition);
                }
            }
            else
            {
                BindingDefinitionDescription bindingDefinitionDescription = BindingDefinitionDescription.CreateBindingDescription(value, trimmedValue);
                ret = BindingDefinitionFactory.CreateInstances(templateDefinitionPart.Parent as ExcelTemplateDefinition, bindingDefinitionDescription);
            }
            return ret;
        }
    }
}
