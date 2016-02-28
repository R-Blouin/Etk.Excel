namespace Etk.ModelManagement
{
    using Etk.BindingTemplates.Definitions.Binding;
    using System;
    using System.Collections.Generic;

    public class ModelProperty : IModelProperty
    {
        #region properties and attributes
        public IBindingDefinition BindingDefinition
        { get; private set; }

        public IModelType Parent
        { get; private set; }

        public string Name
        { get; set; }

        public string Description
        { get; set; }

        public bool IsACollection
        { get { return BindingDefinition.IsACollection; } }

        public IModelType ModelType
        { get; private set; }

        public IEnumerable<IModelProperty>  Properties
        { get { return ModelType == null ? null : ModelType.GetProperties(); } }
        #endregion

        #region .ctors and factories
        public ModelProperty(IModelType parent, IBindingDefinition bindingDefinition)
        {
            Parent = parent;
            BindingDefinition = bindingDefinition;
            Name = BindingDefinition.Name;
            Description = BindingDefinition.Description;
        }
        #endregion

        #region public methods 
        public void ResolveDependencies(IModelDefinitionManager modelDefinition)
        {
            if (ModelType != null)
                return;

            ModelType = modelDefinition.GetModelType(Name);
            if (ModelType == null)
            {
                ModelType = Etk.ModelManagement.ModelType.CreateInstance(modelDefinition, BindingDefinition.BindingType);
                if (ModelType == null)
                    throw new Exception(string.Format("Cannot retrieve Model type for ModelProperty '{0}'", Name));
                modelDefinition.AddModelType(ModelType);
            }
        }
        #endregion
    }
}
