using System;

namespace Etk.BindingTemplates.Definitions.Binding
{
    class BindingDefinitionRoot : BindingDefinition
    {
        public override string Name
        { get { return null; } }

        public override object UpdateDataSource(object dataSource, object data)
        { 
            return null; 
        }

        public override object ResolveBinding(object dataSource)
        {
            return null;
        }

        private  BindingDefinitionRoot(BindingDefinitionDescription definitionDescription) : base(definitionDescription)
        {}

        #region static public methods
        public static BindingDefinitionRoot CreateInstance(Type sourceType)
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
