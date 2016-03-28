using Etk.ModelManagement.DataAccessors;

namespace Etk.ModelManagement
{
    public interface IModelAccessor
    {
        /// <summary> Group that owned this accessor.</summary>
        IModelAccessorGroup Parent { get; }
        /// <summary> Name of the Group that owned this accessor.</summary>
        string ParentName { get; }

        /// <summary> Accessor Ident.</summary>
        string Ident { get; }
        /// <summary> Accessor Name.</summary>
        string Name { get; }
        /// <summary> Accessor Description.</summary>
        string Description { get; }
        /// <summary> Data accessor (contains the information to retrieve the data)</summary>
        IDataAccessor DataAccessor { get; }
        /// <summary> Returned model type.</summary>
        IModelType ReturnModelType { get; }
        /// <summary>True is the return type is a collection of <see cref="IModelType"/> </summary>
        bool ReturnTypeIsACollection { get; }
    }
}
