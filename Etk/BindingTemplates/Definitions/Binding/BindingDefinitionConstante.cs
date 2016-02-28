namespace Etk.BindingTemplates.Definitions.Binding
{
    using Etk.BindingTemplates.Context;
    using Etk.BindingTemplates.Definitions.Binding;

    class BindingDefinitionConstante : BindingDefinition
    {
        public string Value
        { get; private set; }

        override public object UpdateDataSource(object dataSource, object data)
        {
            return Value;
        }

        override public object ResolveBinding(object dataSource)
        {
            return Value;
        }

        private BindingDefinitionConstante(BindingDefinitionDescription bindingDefinitionDescription) : base(bindingDefinitionDescription)
        { }

        static public BindingDefinitionConstante CreateInstance(BindingDefinitionDescription bindingDefinitionDescription)
        {
            bindingDefinitionDescription.IsReadOnly = true;
            return new BindingDefinitionConstante(bindingDefinitionDescription){Value = bindingDefinitionDescription.BindingExpression,
                                                                                IsBoundWithData = false,
                                                                                IsReadOnly = true};
        }
    }
}
