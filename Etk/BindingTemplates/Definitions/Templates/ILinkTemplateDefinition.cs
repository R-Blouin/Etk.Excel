namespace Etk.BindingTemplates.Definitions.Templates
{
    using Etk.BindingTemplates.Definitions.Binding;

    public interface ILinkedTemplateDefinition : IDefinitionPart
    {
        ITemplateDefinition TemplateDefinition { get; }
        IBindingDefinition BindingDefinition { get; }
        LinkedTemplatePositioning Positioning { get; }

        object ResolveBinding(object dataSource);
    }
}