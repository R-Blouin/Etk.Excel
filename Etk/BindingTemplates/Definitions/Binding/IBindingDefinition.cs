namespace Etk.BindingTemplates.Definitions.Binding
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using Etk.BindingTemplates.Context;
    using Etk.BindingTemplates.Definitions;
    using Etk.BindingTemplates.Definitions.Decorators;
    using Etk.BindingTemplates.Definitions.EventCallBacks;

    public interface IBindingDefinition : IDefinitionPart
    {
        /// <summary> Contains the expression used to create the binding definition</summary>
        string BindingExpression { get; }
        /// <summary> Binding definition Name </summary>
        string Name { get; }
        /// <summary> Binding definition Description </summary>
        string Description { get; }

        /// <summary> Type of the object bound with the binding definition</summary>
        Type BindingType { get; }
        /// <summary> If truen then the bound property is a collection</summary>
        bool BindingTypeIsGeneric { get; }
        /// <summary> if 'BindingTypeIsGeneric' = true, contain the generic type (For IEnumerable<double>, then contains 'double'</summary>
        Type BindingGenericType { get; }
        /// <summary> if 'BindingTypeIsGeneric' = true, contain the generic type (For IEnumeravle<double>, then contains 'IEnumerable'</summary>
        Type BindingGenericTypeDefinition { get; }
        /// <summary> Binding Description </summary>
        bool IsACollection { get; }

        /// <summary> True if the BindingType is Nullable</summary>
        bool IsNullable { get; }
        /// <summary> True if the BindingType is a Enum</summary>
        bool IsEnum { get; }
        /// <summary> True if the contain of the cell has more than one line</summary>
        bool IsMultiLine { get; }
        /// <summary> Multiplicator to apply to the number of lines of a multi lines value</summary>
        double MultiLineFactor { get; }

        /// <summary> True if the binding definition is read only </summary>
        bool IsReadOnly { get; }
        /// <summary> True if the binding definition is optional </summary>
        bool IsOptional { get; }
        /// <summary> True if the binding definition is not a constante</summary>
        bool IsBoundWithData { get; }
        /// <summary>Update the data bound with the binding definition</summary>
        /// <param name="dataSource">The data source</param>
        /// <param name="data">The new value to inject to the object bound with the binding definition</param>
        /// <returns>The changed data</returns>
        object UpdateDataSource(object dataSource, object data);
        /// <summary>Invoke the binding definition: return the value of the binding definition</summary>
        /// <param name="dataSource">The data source</param>
        /// <returns></returns>
        object ResolveBinding(object dataSource);

        /// <summary> True if the binding definition can notify any modifications </summary>
        bool CanNotify { get; }
        bool MustNotify(object dataSource, object source, PropertyChangedEventArgs args);
        IEnumerable<INotifyPropertyChanged> GetObjectsToNotify(object dataSource);

        /// <summary>The factory used to create an context item (a contextual link between the data source, the template view, and the binding definition)</summary>
        /// <param name="parent">The parent context element</param>
        /// <returns>The new created context item</returns>
        IBindingContextItem ContextItemFactory(IBindingContextElement parent);

        /// <summary>If defined, the decorator used to modify the style of the rendered object used to present the data linked with the binding definition</summary>
        Decorator DecoratorDefinition { get; }

        /// <summary> Contains the callback to invoke when the bound object is selected</summary>
        EventCallback OnSelection {get ;}

        /// <summary> Contains the callback to invoke when the bound object is clicked (left double click in Excel)</summary>
        EventCallback OnClick { get; }
    }
}
