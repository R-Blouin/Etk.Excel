using System.Collections.Generic;
using System.ComponentModel;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.Excel.BindingTemplates.Definitions;

namespace Etk.Excel.BindingTemplates.Controls.FormulaResult
{
    class ExcelBindingDefinitionWithFormula : BindingDefinition
    {
        #region attributes and properties
        public const string FORMULA_RESULT_PREFIX = "{>";

        public IBindingDefinition UseFormulaBindingDefinition
        { get; private set; }

        public IBindingDefinition NestedBindingDefinition
        { get; private set; }

        public override string Name
        { get { return NestedBindingDefinition != null ? NestedBindingDefinition.Name : string.Empty; } }

        public override string Description
        { get { return NestedBindingDefinition != null ? NestedBindingDefinition.Description : string.Empty; } }
        #endregion

        #region .ctors and factories
        private ExcelBindingDefinitionWithFormula(BindingDefinitionDescription bindingDefinitionDescription, IBindingDefinition underlyingBindingDefinition, IBindingDefinition useFormulaBindingDefinition)
                                                    : base(bindingDefinitionDescription) 
        {
            NestedBindingDefinition = underlyingBindingDefinition;
            UseFormulaBindingDefinition = useFormulaBindingDefinition;
            CanNotify = underlyingBindingDefinition.CanNotify;

            DefinitionDescription = new BindingDefinitionDescription();
        }

        public static ExcelBindingDefinitionFormulaResult CreateInstance(ExcelTemplateDefinition templateDefinition, BindingDefinitionDescription definition)
        {
            return null;
            //try
            //{
            //    definition = definition.Replace(FORMULA_RESULT_PREFIX, string.Empty);
            //    definition = definition.TrimEnd('}');

            //    //UseFormulaBindingDefinition
            //    string[] parts = definition.Split(';');
            //    if (parts.Count() > 2)
            //        throw new ArgumentException(string.Format("dataAccessor '{0}' is invalid.", definition));

            //    string useFormulaDefinition = null;
            //    string underlyingDefinition;
            //    if (parts.Count() == 1)
            //        underlyingDefinition = string.Format("{{{0}}}", parts[0].Trim());
            //    else
            //    {
            //        useFormulaDefinition = string.Format("{{{0}}}", parts[0].Trim());
            //        underlyingDefinition = string.Format("{{{0}}}", parts[1].Trim());
            //    }

            //    BindingDefinitionDescription bindingDefinitionDescription = BindingDefinitionDescription.CreateBindingDescription(templateDefinition, underlyingDefinition, underlyingDefinition);
            //    IBindingDefinition underlyingBindingDefinition = BindingDefinitionFactory.CreateInstances(templateDefinition, bindingDefinitionDescription);

            //    IBindingDefinition useFormulaBindingDefinition = null;
            //    if (!string.IsNullOrEmpty(useFormulaDefinition))
            //    {
            //        bindingDefinitionDescription = BindingDefinitionDescription.CreateBindingDescription(templateDefinition, useFormulaDefinition, useFormulaDefinition);
            //        useFormulaBindingDefinition = BindingDefinitionFactory.CreateInstances(templateDefinition, bindingDefinitionDescription);
            //    }
            //    ExcelBindingDefinitionFormulaResult ret = new ExcelBindingDefinitionFormulaResult(bindingDefinitionDescription, underlyingBindingDefinition, useFormulaBindingDefinition);
            //    return ret;
            //}
            //catch (Exception ex)
            //{
            //    string message = string.Format("Cannot retrieve the formula result binding dataAccessor '{0}'. {1}", definition.EmptyIfNull(), ex.Message);
            //    throw new EtkException(message);
            //}
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
            return NestedBindingDefinition != null && NestedBindingDefinition.MustNotify(dataSource, source, args);
        }

        public override IEnumerable<INotifyPropertyChanged> GetObjectsToNotify(object dataSource)
        {
            return NestedBindingDefinition == null ? null : NestedBindingDefinition.GetObjectsToNotify(dataSource);
        }

    }
}
