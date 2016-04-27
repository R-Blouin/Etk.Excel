using System;
using System.Collections.Generic;
using System.Reflection;
using Etk.Excel.ContextualMenus.Definition;
using Etk.Tools.Extensions;
using ExcelInterop = Microsoft.Office.Interop.Excel; 

namespace Etk.Excel.ContextualMenus
{
    public delegate void MenuAction(ExcelInterop.Range range);

    static class ContextualMenuFactory
    {
        /// <summary>
        /// Create a contextual menu from an Xml
        /// </summary>
        public static IEnumerable<IContextualMenu> CreateInstances(string xmlValue)
        {
            List<IContextualMenu> ret = null;
            XmlContextualMenuDefinitions definitions = null;
            try
            {
                definitions = XmlContextualMenuDefinitions.CreateInstance(xmlValue);
                if (definitions != null && definitions.ContextualMenus != null)
                {
                    ret = new List<IContextualMenu>();

                    foreach (XmlContextualMenuDefinition definition in definitions.ContextualMenus)
                    {
                        definition.Name = definition.Name.EmptyIfNull().Trim();
                        if (string.IsNullOrEmpty(definition.Name))
                            throw new EtkException("Contextual menu must have a name");

                        if (definition.Items != null && definition.Items.Count > 0)
                        {
                            List<IContextualPart> items = new List<IContextualPart>();
                            foreach (XmlContextualMenuPart xmlPart in definition.Items)
                            {
                                XmlContextualMenuItemDefinition xmlItem = xmlPart as XmlContextualMenuItemDefinition;
                                MethodInfo methodInfo = ConstextualMethodRetriever.RetrieveContextualMethodInfo(xmlItem.Action);
                                IContextualMenuItem menuItem = new ContextualMenuItem(xmlItem.Caption, xmlItem.BeginGroup, methodInfo, xmlItem.FaceId);

                                items.Add((IContextualPart)menuItem);
                            }
                            if (items.Count > 0)
                                ret.Add(new ContextualMenu(definition.Name, definition.Caption, definition.BeginGroup, items));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string message = xmlValue.Length > 350 ? xmlValue.Substring(0, 350) + "..." : xmlValue;
                throw new EtkException(string.Format("Cannot create Event Callbacks from xml '{0}':{1}", message, ex.Message));
            }
            return ret;
        }

        /*
        /// <summary>
        /// Create a contextual menu from a menu excelTemplateDefinition contains in a Excel Cell
        /// </summary>
        static public IContextualMenu CreateInstanceFromCell(Range targetRange, string contextMenuPath, Type mainBindingDefinitionType)
        {
            IContextualMenu menu = null;

            if (!string.IsNullOrEmpty(contextMenuPath))
            {
                string menuName = null;
                Worksheet sheetContainer = null;
                try
                {
                    string[] contextMenuParts = contextMenuPath.Split('.');
                    Worksheet currentWorkSheet = targetRange.Worksheet;
                    if (contextMenuParts.Count() == 1)
                    {
                        sheetContainer = targetRange.Worksheet;
                        menuName = contextMenuParts[0].EmptyIfNull().Trim();
                    }
                    else
                    {
                        string worksheetContainerName = contextMenuParts[0].EmptyIfNull().Trim();
                        menuName = contextMenuParts[1].EmptyIfNull().Trim();

                        Workbook workbook = currentWorkSheet.ParentElement as Workbook;
                        if (workbook == null)
                            throw new EtkException("Cannot retrieve the sheet workbook");

                        List<Worksheet> sheets = new List<Worksheet>(workbook.Worksheets.Cast<Worksheet>());
                        sheetContainer = sheets.FirstOrDefault(s => !string.IsNullOrEmpty(s.Name) && s.Name.Equals(worksheetContainerName));
                        if (sheetContainer == null)
                            throw new EtkException(string.Format("Cannot find the sheet '{0}' in the current workbook", worksheetContainerName), false);

                        Marshal.ReleaseComObject(workbook);
                        workbook = null;
                    }

                    Range menuRange = sheetContainer.Cells.Find(string.Format(ExcelTemplateDefinition.CONTEXTUALMENU_START_FORMAT, menuName), Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false);
                    if (menuRange == null)
                        throw new EtkException("Cannot find the ContextualMenu");

                    Marshal.ReleaseComObject(sheetContainer);
                    sheetContainer = null;
                    Marshal.ReleaseComObject(currentWorkSheet);
                    currentWorkSheet = null;

                    string xmlValue = menuRange.Value2;
                    XmlContextualMenuDefinition xmlDefinition = XmlContextualMenuDefinition.CreateInstances(xmlValue);
                    menu = ContextualMenuFactory.CreateInstanceFromCell(menuRange, xmlDefinition, mainBindingDefinitionType);
                }
                catch (Exception ex)
                {
                    throw new EtkException(string.Format("Cannot create ContextualMenu '{0}'", menuName.EmptyIfNull(), sheetContainer == null ? string.Empty : sheetContainer.Name.EmptyIfNull()), ex);
                }
            }
            return menu;
        }

        /// <summary>
        /// Create a contextual menu from a menu excelTemplateDefinition contains in a Excel Cell
        /// </summary>
        static private IContextualMenu CreateInstanceFromCell(Range targetRange, XmlContextualMenuDefinition excelTemplateDefinition, Type mainBindingDefinitionType)
        {
            ContextualMenu ret = null;
            if (excelTemplateDefinition != null)
            {
                excelTemplateDefinition.Name = excelTemplateDefinition.Name.EmptyIfNull().Trim();
                if (string.IsNullOrEmpty(excelTemplateDefinition.Name))
                    throw new EtkException("Contextual menu must have a name");

                if (excelTemplateDefinition.Items != null && excelTemplateDefinition.Items.Count > 0)
                {
                    List<IContextualPart> items = new List<IContextualPart>();
                    foreach (XmlContextualMenuPart xmlPart in excelTemplateDefinition.Items)
                    {
                        if (xmlPart is XmlContextualMenuDefinition)
                        {
                            XmlContextualMenuDefinition xmlMenu = xmlPart as XmlContextualMenuDefinition;
                            IContextualMenu subMenu = null;
                            // If no item, then we consider that the menu is set elsewhere
                            if (xmlMenu.Items == null || xmlMenu.Items.Count() == 0)
                                subMenu = CreateInstanceFromCell(targetRange, xmlMenu.Name, mainBindingDefinitionType);
                            else
                                subMenu = CreateInstanceFromCell(targetRange, xmlMenu, mainBindingDefinitionType);
                            if (subMenu != null)
                                items.Add((IContextualPart) subMenu);
                        }
                        else
                        {
                            XmlContextualMenuItemDefinition xmlItem = xmlPart as XmlContextualMenuItemDefinition;
                            MethodInfo methodInfo = ConstextualMethodRetriever.RetrieveContextualMethodInfo(mainBindingDefinitionType, xmlItem.Action);
                            IContextualMenuItem menuItem = new ContextualMenuItem(xmlItem.Caption, xmlItem.BeginGroup, methodInfo, xmlItem.FaceId);
                            items.Add((IContextualPart) menuItem);
                        }
                    }
                    if(items.Count > 0)
                        ret = new ContextualMenu(excelTemplateDefinition.Name, excelTemplateDefinition.Caption, excelTemplateDefinition.BeginGroup, items);
                }
            }
            return ret;
        }
        */
    }
}
