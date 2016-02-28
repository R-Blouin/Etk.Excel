namespace Etk.BindingTemplates.Definitions.Binding
{
    using System;
    using Etk.BindingTemplates.Context;

    class BindingDefinitionRoot : BindingDefinition
    {
        public override string Name
        { get { return null; } }

        override public object UpdateDataSource(object dataSource, object data)
        { 
            return null; 
        }

        override public object ResolveBinding(object dataSource)
        {
            return null;
        }

        private  BindingDefinitionRoot(BindingDefinitionDescription definitionDescription) : base(definitionDescription)
        {}

        #region static public methods
        static public BindingDefinitionRoot CreateInstance(Type sourceType)
        {
            BindingDefinitionDescription definitionDescription = new BindingDefinitionDescription() { IsReadOnly = true };

            BindingDefinitionRoot definition = new BindingDefinitionRoot(definitionDescription) { BindingType = sourceType };
            if (definition.BindingType != null)
            {
                definition.ManageCollectionStatus();
                definition.ManageEnumAndNullable();
            }
            return definition;
        }
        #endregion
    }
}
