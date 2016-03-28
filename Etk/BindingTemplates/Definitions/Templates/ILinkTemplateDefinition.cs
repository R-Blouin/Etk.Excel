using Etk.BindingTemplates.Definitions.Binding;

namespace Etk.BindingTemplates.Definitions.Templates
{
    public interface ILinkedTemplateDefinition : IDefinitionPart
    {
        ITemplateDefinition TemplateDefinition { get; }
        IBindingDefinition BindingDefinition { get; }
        LinkedTemplatePositioning Positioning { get; }

        object ResolveBinding(object dataSource);
    }
}