namespace Etk.Excel.BindingTemplates.Definitions
{
    using System;
    using Etk.BindingTemplates.Definitions.EventCallBacks;
    using Etk.BindingTemplates.Definitions.Templates;
    using Etk.Excel.ContextualMenus;
    using Etk.Excel.UI.Extensions;
    using Microsoft.Office.Interop.Excel;
    using Etk.BindingTemplates.Definitions.Decorators;
    using Etk.Excel.BindingTemplates.Decorators;

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

        public Range DefinitionFirstCell
        { get; private set; }

        public Range DefinitionLastCell
        { get; private set; }

        public IContextualMenu ContextualMenu
        { get; internal set; }

        public EventCallback SelectionChanged
        { get; internal set; }

        public EventCallback OnLeftDoubleClick
        { get; internal set; }

        public ExcelRangeDecorator Decorator
        { get; private set; }
        #endregion

        #region .ctors and factories
        internal ExcelTemplateDefinition(Range firstRange, Range lastRange, TemplateOption templateOption) : base(templateOption)
        {
            DefinitionFirstCell = firstRange;
            DefinitionLastCell = lastRange;

            Width = DefinitionLastCell.Column - DefinitionFirstCell.Column + 1;
            Height = DefinitionLastCell.Row - DefinitionFirstCell.Row + 1;
        }
        #endregion

        #region internal metrhods
        internal protected void ExcelInit(ITemplateDefinitionPart header, ITemplateDefinitionPart body, ITemplateDefinitionPart footer)
        {
            Init(header, body, footer);

            RetrieveSelectionChangeMethod();
            RetrieveContextualMenuMethod();
            RetrieveDecorator();
        }
        #endregion

        #region private method
        private void RetrieveContextualMenuMethod()
        {
            string contextMenu = TemplateOption.ContextualMenu;
            if (!string.IsNullOrEmpty(contextMenu))
                ContextualMenu = ETKExcel.ContextualMenuManager.GetContextualMenu(contextMenu);
        }

        private void RetrieveSelectionChangeMethod()
        {
            string selectionChanged = TemplateOption.SelectionChanged.EmptyIfNull();
            selectionChanged = selectionChanged.Trim();

            if (!string.IsNullOrEmpty(selectionChanged))
            {
                try
                {
                    Type type = MainBindingDefinition.BindingTypeIsGeneric ? MainBindingDefinition.BindingGenericType : MainBindingDefinition.BindingType;
                    SelectionChanged = EventCallback.CreateInstance(null, null, type, selectionChanged);
                }
                catch (Exception ex)
                {
                    throw new EtkException("Retrieve 'ChangeSelection' method information failed", ex);
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
