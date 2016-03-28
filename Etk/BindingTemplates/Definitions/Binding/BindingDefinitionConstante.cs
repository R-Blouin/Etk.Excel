namespace Etk.BindingTemplates.Definitions.Binding
{
    class BindingDefinitionConstante : BindingDefinition
    {
        public string Value
        { get; private set; }

        public override object UpdateDataSource(object dataSource, object data)
        {
            return Value;
        }

        public override object ResolveBinding(object dataSource)
        {
            return Value;
        }

        private BindingDefinitionConstante(BindingDefinitionDescription bindingDefinitionDescription) : base(bindingDefinitionDescription)
        { }

        public static BindingDefinitionConstante CreateInstance(BindingDefinitionDescription bindingDefinitionDescription)
        {
            bindingDefinitionDescription.IsReadOnly = true;
            return new BindingDefinitionConstante(bindingDefinitionDescription){Value = bindingDefinitionDescription.BindingExpression,
                                                                                IsBoundWithData = false,
                                                                                IsReadOnly = true};
        }
    }
}
