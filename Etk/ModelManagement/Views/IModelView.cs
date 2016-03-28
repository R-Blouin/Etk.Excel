using System.Collections.Generic;

namespace Etk.ModelManagement.Views
{
    /// <summary> Model view definition. Describe a view on a model type</summary>
    public interface IModelView : IModelViewPart 
    {
        /// <summary> Contain a reference to the view underlying model type.</summary>
        IModelType Parent { get; }

        /// <summary> Model View name.</summary>
        string Name { get; }
        /// <summary> Model View Description.</summary>
        string Description { get; }

        /// <summary> Parts of the view (can be ModelView or ModelProperty) </summary> 
        IEnumerable<IModelViewPart> Parts  { get; }
    }
}
