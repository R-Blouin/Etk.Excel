using System;
using System.Linq;
using Etk.BindingTemplates.Definitions.Decorators;
using Etk.BindingTemplates.Definitions.EventCallBacks;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.Excel.BindingTemplates.Decorators;
using Etk.Excel.ContextualMenus;
using Etk.Tools.Extensions;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using System.Reflection;
using Etk.Tools.Reflection;
using System.Runtime.InteropServices;

namespace Etk.Excel.BindingTemplates.Definitions
{
    public class ExcelTemplateDefinition : TemplateDefinition
    {
        #region const
        private static DecoratorsManager decoratorsManager;
        private static DecoratorsManager DecoratorsManager
        {
            get
            {
                if (decoratorsManager == null)
                    decoratorsManager = CompositionManager.Instance.GetExportedValue<DecoratorsManager>();
                return decoratorsManager;
            }
        }

        public int Width
        { get; private set; }

        public int Height
        { get; private set; }

        public ExcelInterop.Range DefinitionFirstCell
        { get; private set; }

        public ExcelInterop.Range DefinitionLastCell
        { get; private set; }

        public IContextualMenu ContextualMenu
        { get; internal set; }

        //public bool IsSelectionChangedCom
        //{ get; internal set; }

        public MethodInfo OnLeftDoubleClick
        { get; internal set; }

        public ExcelRangeDecorator Decorator
        { get; private set; }

        public MethodInfo SelectionChanged { get; private set; }

        public string SelectionChangedComFonctionName { get; private set; }
        #endregion

        #region .ctors and factories
        internal ExcelTemplateDefinition(ExcelInterop.Range firstRange, ExcelInterop.Range lastRange, TemplateOption templateOption)
            : base(templateOption)
        {
            DefinitionFirstCell = firstRange;
            DefinitionLastCell = lastRange;

            Width = DefinitionLastCell.Column - DefinitionFirstCell.Column + 1;
            Height = DefinitionLastCell.Row - DefinitionFirstCell.Row + 1;
        }
        #endregion

        #region internal metrhods
        protected internal void ExcelInit(ITemplateDefinitionPart header, ITemplateDefinitionPart body, ITemplateDefinitionPart footer)
        {
            Init(header, body, footer);

            RetrieveSelectionChangeMethod();
            RetrieveContextualMenu();
            RetrieveDecorator();
        }
        #endregion

        #region private method
        private void RetrieveContextualMenu()
        {
            string contextMenuRef = TemplateOption.ContextualMenu;
            if (!string.IsNullOrEmpty(contextMenuRef))
            {
                ContextualMenu = ETKExcel.ContextualMenuManager.GetContextualMenu(contextMenuRef);
                if(ContextualMenu == null)
                {
                    ExcelInterop.Worksheet worksheet = null;
                    ExcelInterop.Range menuRange = null;
                    try
                    {
                        worksheet = DefinitionFirstCell.Worksheet;
                        try
                        {
                            menuRange = worksheet.Range[contextMenuRef];
                        }
                        catch
                        { }
                        if(menuRange != null)
                            ContextualMenu = ETKExcel.ContextualMenuManager.RegisterMenuDefinitionFromXml(menuRange.Value2);
                    }
                    finally
                    {
                        if (worksheet != null)
                        {
                            Marshal.ReleaseComObject(worksheet);
                            worksheet = null;
                        }
                        menuRange = null;
                    }
                }
                if (ContextualMenu == null)
                    throw new Exception($"Cannot find contextual menu '{contextMenuRef ?? string.Empty}'");
            }
        }

        private void RetrieveSelectionChangeMethod()
        {
            string selectionChanged = TemplateOption.SelectionChanged.EmptyIfNull();
            selectionChanged = selectionChanged.Trim();

            if (!string.IsNullOrEmpty(selectionChanged))
            {
                try
                {
                    if(selectionChanged.StartsWith("$"))
                    {
                        SelectionChanged = TypeHelpers.GetMethod(typeof(EventExcelCallbacksManager), "ComInvoke");
                        SelectionChangedComFonctionName = selectionChanged.TrimStart('$');
                    }
                    else
                    {
                        Type type = MainBindingDefinition.BindingTypeIsGeneric ? MainBindingDefinition.BindingGenericType : MainBindingDefinition.BindingType;

                        string[] parts = selectionChanged.Split(',');
                        if (parts.Count() == 1)
                        {
                            EventCallback callback = ((ExcelTemplateManager)ETKExcel.TemplateManager).CallbacksManager.GetCallback(selectionChanged);
                            if (callback != null)
                                SelectionChanged = callback.Callback;
                        }
                        if (parts.Count() == 3)
                            SelectionChanged = TypeHelpers.GetMethod(null, selectionChanged);

                        if (SelectionChanged == null)
                            throw new Exception($"Cannot find the callback '{selectionChanged}'");
                    }
                }
                catch (Exception ex)
                {
                    throw new EtkException($"Retrieve 'SelectionChanged' method information failed:{ex.Message}");
                }
            }
        }

        private void RetrieveDecorator()
        {
            if (!string.IsNullOrEmpty(TemplateOption.DecoratorIdent))
                Decorator = DecoratorsManager.GetDecorator(TemplateOption.DecoratorIdent) as ExcelRangeDecorator;
        }
        #endregion
    }
}
