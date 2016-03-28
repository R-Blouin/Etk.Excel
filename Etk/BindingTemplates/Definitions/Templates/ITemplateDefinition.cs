using Etk.BindingTemplates.Definitions.Binding;
using Etk.ModelManagement.DataAccessors;

namespace Etk.BindingTemplates.Definitions.Templates
{
    public interface ITemplateDefinition
    {
        /// <summary> Template Name</summary>
        string Name { get; }

        /// <summary> Template Description</summary>
        string Description { get; }

        /// <summary> Get the main binding definition type.
        /// If supplied, then the 'BindingDefinitions' are resolved (if possible) at the template creation.</summary>
        /// Else the binding definition are resolved at runtime (late binding) 
        IBindingDefinition MainBindingDefinition { get; }

        /// <summary> Template Orientation (vertical or horizontal)</summary>
        Orientation Orientation { get; }

        /// <summary> Method info to invoke to set the datasource of the views based on tha template</summary>
        IDataAccessor DataAccessor{ get;}

        /// <summary> Contain the definition of the Header of the current template</summary>
        ITemplateDefinitionPart Header { get; }
        /// <summary> Contain the definition of the body of the current template</summary>
        ITemplateDefinitionPart Body { get; }
        /// <summary> Contains the definition of the Footer of the current template</summary>
        ITemplateDefinitionPart Footer { get; }

        /// <summary> Contains the type dynamically created at runtime from the template definition.
        /// Use by the ETK sort and filter process.</summary>
        BindingType BindingType { get; }
    }
}
