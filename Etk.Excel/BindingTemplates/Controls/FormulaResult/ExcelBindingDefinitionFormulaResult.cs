namespace Etk.Excel.BindingTemplates.Controls.FormulaResult
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using Etk.BindingTemplates.Context;
    using Etk.BindingTemplates.Definitions.Binding;
    using Etk.Excel.BindingTemplates.Definitions;
    using Etk.Excel.UI.Extensions;

    class ExcelBindingDefinitionFormulaResult : BindingDefinition
    {
        #region attributes and properties
        public const string FORMULA_RESULT_PREFIX = "{>";

        public IBindingDefinition UseFormulaBindingDefinition
        { get; private set; }

        public IBindingDefinition NestedBindingDefinition
        { get; private set; }

        override public string Name
        { get { return NestedBindingDefinition != null ? NestedBindingDefinition.Name : string.Empty; } }

        override public string Description
        { get { return NestedBindingDefinition != null ? NestedBindingDefinition.Description : string.Empty; } }
        #endregion

        #region .ctors and factories
        private ExcelBindingDefinitionFormulaResult(BindingDefinitionDescription bindingDefinitionDescription, IBindingDefinition underlyingBindingDefinition, IBindingDefinition useFormulaBindingDefinition)
                                                    : base(bindingDefinitionDescription) 
        {
            NestedBindingDefinition = underlyingBindingDefinition;
            UseFormulaBindingDefinition = useFormulaBindingDefinition;
            CanNotify = underlyingBindingDefinition.CanNotify;

            DefinitionDescription = new BindingDefinitionDescription();
        }

        static public ExcelBindingDefinitionFormulaResult CreateInstance(ExcelTemplateDefinition templateDefinition, string definition)
        {
            try
            {
                definition = definition.Replace(FORMULA_RESULT_PREFIX, string.Empty);
                definition = definition.TrimEnd('}');

                //UseFormulaBindingDefinition
                string[] parts = definition.Split(';');
                if (parts.Count() > 2)
                    throw new ArgumentException(string.Format("dataAccessor '{0}' is invalid.", definition));

                string useFormulaDefinition = null;
                string underlyingDefinition = null;
                if (parts.Count() == 1)
                    underlyingDefinition = string.Format("{{{0}}}", parts[0].Trim());
                else
                {
                    useFormulaDefinition = string.Format("{{{0}}}", parts[0].Trim());
                    underlyingDefinition = string.Format("{{{0}}}", parts[1].Trim());
                }

                BindingDefinitionDescription bindingDefinitionDescription = BindingDefinitionDescription.CreateBindingDescription(underlyingDefinition, underlyingDefinition);
                IBindingDefinition underlyingBindingDefinition = BindingDefinitionFactory.CreateInstances(templateDefinition, bindingDefinitionDescription);

                IBindingDefinition useFormulaBindingDefinition = null;
                if (!string.IsNullOrEmpty(useFormulaDefinition))
                {
                    bindingDefinitionDescription = BindingDefinitionDescription.CreateBindingDescription(useFormulaDefinition, useFormulaDefinition);
                    useFormulaBindingDefinition = BindingDefinitionFactory.CreateInstances(templateDefinition, bindingDefinitionDescription);
                }
                ExcelBindingDefinitionFormulaResult ret = new ExcelBindingDefinitionFormulaResult(bindingDefinitionDescription, underlyingBindingDefinition, useFormulaBindingDefinition);
                return ret;
            }
            catch (Exception ex)
            {
                string message = string.Format("Cannot retrieve the formula result binding dataAccessor '{0}'. {1}", definition.EmptyIfNull(), ex.Message);
                throw new EtkException(message, ex);
            }
        }
        #endregion


        override public IBindingContextItem ContextItemFactory(IBindingContextElement parent)
        {
            IBindingContextItem nestedContextItem = NestedBindingDefinition.ContextItemFactory(parent);
            return new ExcelContextItemFormulaResult(parent, this);
        }

        //  cannot be reached... Managed in 'ExcelContextItemFormulaResult'
        override public object UpdateDataSource(object dataSource, object data)
        {
            return null;
        }

        //  cannot be reached... Managed in 'ExcelContextItemFormulaResult'
        override public object ResolveBinding(object dataSource)
        {
            return null;
        }

        override public bool MustNotify(object dataSource, object source, PropertyChangedEventArgs args)
        {
            return NestedBindingDefinition != null && NestedBindingDefinition.MustNotify(dataSource, source, args);
        }

        override public IEnumerable<INotifyPropertyChanged> GetObjectsToNotify(object dataSource)
        {
            return NestedBindingDefinition == null ? null : NestedBindingDefinition.GetObjectsToNotify(dataSource);
        }
    }
}
