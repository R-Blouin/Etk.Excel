namespace Etk.ModelManagement
{
    using Etk.BindingTemplates.Definitions.Binding;

    /// <summary> For the primitive type</summary>
    public class ModelLeafProperty : IModelProperty
    {
        public IModelType Parent
        { get; private set; }

        public IBindingDefinition BindingDefinition
        { get; private set; }

        public string Name
        { get; set; }

        public string Description
        { get; set; }

        public bool IsACollection
        { get { return false; } }

        public ModelLeafProperty(IModelType parent, IBindingDefinition bindingDefinition)
        {
            Parent = parent;
            BindingDefinition = bindingDefinition;
            Name = BindingDefinition.Name;
            Description = BindingDefinition.Description;
        }
    }
}
