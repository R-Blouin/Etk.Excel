using System.Collections.Generic;

namespace Etk.ModelManagement
{
    public interface IModelAccessorGroup
    {
        /// <summary>Model definition manager that owned the group.</summary>
        IModelDefinitionManager Parent{ get;}

        /// <summary> Accessor group Name.</summary>
        string Name { get; }
        /// <summary> Accessor group description.</summary>
        string Description { get; }

        /// <summary> Accessor group description.</summary>
        IEnumerable<IModelAccessor> Accessors { get; }
    }
}
