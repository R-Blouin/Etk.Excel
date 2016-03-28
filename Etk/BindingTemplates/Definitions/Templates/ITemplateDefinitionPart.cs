using System.Collections.Generic;
using Etk.BindingTemplates.Definitions.Binding;

namespace Etk.BindingTemplates.Definitions.Templates
{
    public interface ITemplateDefinitionPart
    {
        /// <summary> The template definition that owned that part.</summary>
        ITemplateDefinition Parent { get; }

        /// <summary> Get the binding parts of the template</summary>
        List<IDefinitionPart> BindingParts { get; }

        /// <summary> Contains the binding definitions owned by the current template</summary>
        List<IBindingDefinition> BindingDefinitions { get; }

        /// <summary> Contains the list of the templates use in the current one</summary>
        List<ILinkedTemplateDefinition> LinkedTemplates { get; }

        /// <summary> True if 'LinkedTemplates' has items.</summary>
        bool HasLinkedTemplates { get; }
    }
}