using System;
using System.Collections.Generic;
using System.ComponentModel;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.Excel.BindingTemplates.Definitions;

namespace Etk.Excel.BindingTemplates.Controls.WithFormula
{
    class ExcelBindingDefinitionWithFormula : BindingDefinition
    {
        #region attributes and properties
        public const string FORMULA_RESULT_PREFIX = "{>";

        public IBindingDefinition FormulaBindingDefinition
        { get; private set; }

        public IBindingDefinition TargetBindingDefinition
        { get; private set; }

        public override string Name => TargetBindingDefinition != null ? TargetBindingDefinition.Name : string.Empty;

        public override string Description => TargetBindingDefinition != null ? TargetBindingDefinition.Description : string.Empty;

        #endregion

        #region .ctors and factories
        private ExcelBindingDefinitionWithFormula(BindingDefinitionDescription bindingDefinitionDescription, IBindingDefinition targetBindingDefinition, IBindingDefinition formulaBindingDefinition)
                                                 : base(bindingDefinitionDescription) 
        {
            TargetBindingDefinition = targetBindingDefinition;
            FormulaBindingDefinition = formulaBindingDefinition;
            if(TargetBindingDefinition != null)
                CanNotify = TargetBindingDefinition.CanNotify;

            DefinitionDescription = new BindingDefinitionDescription();
        }

        public static ExcelBindingDefinitionWithFormula CreateInstance(ExcelTemplateDefinition templateDefinition, BindingDefinitionDescription definition)
        {
            try
            {
                IBindingDefinition formulaBindingDefinition = null;
                IBindingDefinition targetBindingDefinition = null;

                if (! string.IsNullOrEmpty(definition.Formula))
                {
                    BindingDefinitionDescription formulaBindingDefinitionDescription = BindingDefinitionDescription.CreateBindingDescription(templateDefinition, definition.Formula, definition.Formula);
                    formulaBindingDefinition = BindingDefinitionFactory.CreateInstances(templateDefinition, formulaBindingDefinitionDescription);
                }

                if (!string.IsNullOrEmpty(definition.BindingExpression))
                {
                    string bindingExpression = $"{{{definition.BindingExpression}}}";
                    BindingDefinitionDescription targetBindingDefinitionDescription = BindingDefinitionDescription.CreateBindingDescription(templateDefinition, bindingExpression, bindingExpression);
                    targetBindingDefinition = BindingDefinitionFactory.CreateInstances(templateDefinition, targetBindingDefinitionDescription);
                }

                ExcelBindingDefinitionWithFormula ret = new ExcelBindingDefinitionWithFormula(definition, targetBindingDefinition, formulaBindingDefinition);
                return ret;
            }
            catch (Exception ex)
            {
                string message = $"Cannot create the 'ExcelBindingDefinitionWithFormula' from '{definition.BindingExpression ?? string.Empty}'. {ex.Message}";
                throw new EtkException(message);
            }
        }
        #endregion

        public override IBindingContextItem ContextItemFactory(IBindingContextElement parent)
        {
            //IBindingContextItem nestedContextItem = NestedBindingDefinition.ContextItemFactory(parent);
            return new ExcelContextItemWithFormula(parent, this);
        }

        //  Cannot be reached... Managed in 'ExcelContextItemWithFormula'
        public override object UpdateDataSource(object dataSource, object data)
        {
            return null;
        }

        //  Cannot be reached... Managed in 'ExcelContextItemWithFormula'
        public override object ResolveBinding(object dataSource)
        {
            return null;
        }

        public override bool MustNotify(object dataSource, object source, PropertyChangedEventArgs args)
        {
            return TargetBindingDefinition != null && TargetBindingDefinition.MustNotify(dataSource, source, args);
        }

        public override IEnumerable<INotifyPropertyChanged> GetObjectsToNotify(object dataSource)
        {
            return TargetBindingDefinition == null ? null : TargetBindingDefinition.GetObjectsToNotify(dataSource);
        }

    }
}
