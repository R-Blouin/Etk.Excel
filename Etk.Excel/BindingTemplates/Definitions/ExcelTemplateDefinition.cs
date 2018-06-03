using System;
using Etk.BindingTemplates.Definitions.Decorators;
using Etk.BindingTemplates.Definitions.EventCallBacks;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.Excel.BindingTemplates.Decorators;
using Etk.Excel.ContextualMenus;
using Etk.Tools.Extensions;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Etk.Excel.Application;

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

        private static EventCallbacksManager eventCallbacksManager;
        private static EventCallbacksManager EventCallbacksManager => eventCallbacksManager ??
                                                                      (eventCallbacksManager = CompositionManager.Instance.GetExportedValue<EventCallbacksManager>());

        public int Width
        { get; private set; }

        public int Height
        { get; private set; }

        public ExcelInterop.Range DefinitionFirstCell
        { get;  }

        public ExcelInterop.Range DefinitionLastCell
        { get; }

        public IContextualMenu ContextualMenu
        { get; internal set; }

        public ExcelRangeDecorator Decorator
        { get; private set; }

        public EventCallback OnLeftDoubleClick
        { get; internal set; }

        public EventCallback SelectionChanged { get; private set; }
        #endregion

        #region .ctors and factories
        internal ExcelTemplateDefinition(ExcelInterop.Range firstRange, ExcelInterop.Range lastRange, TemplateOption templateOption)
            : base(templateOption)
        {
            DefinitionFirstCell = firstRange[1, 1];
            DefinitionLastCell = lastRange[1, 1];

            Width = DefinitionLastCell.Column - DefinitionFirstCell.Column + 1;
            Height = DefinitionLastCell.Row - DefinitionFirstCell.Row + 1;
        }

        ~ExcelTemplateDefinition()
        {
            ExcelApplication.ReleaseComObject(DefinitionFirstCell);
            ExcelApplication.ReleaseComObject(DefinitionLastCell);
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
                            ExcelApplication.ReleaseComObject(worksheet);
                        if (menuRange != null)
                            ExcelApplication.ReleaseComObject(menuRange);
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
                    SelectionChanged = EventCallbacksManager.RetrieveCallback(this, selectionChanged);
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
