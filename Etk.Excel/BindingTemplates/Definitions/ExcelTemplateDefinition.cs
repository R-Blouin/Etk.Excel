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

        public MethodInfo SelectionChanged
        { get; internal set; }

        public MethodInfo OnLeftDoubleClick
        { get; internal set; }

        public ExcelRangeDecorator Decorator
        { get; private set; }
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
                        throw new Exception(string.Format("Cannot find the callback '{0}'", selectionChanged));
                }
                catch (Exception ex)
                {
                    throw new EtkException(string.Format("Retrieve 'SelectionChanged' method information failed:[0}", ex.Message));
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
