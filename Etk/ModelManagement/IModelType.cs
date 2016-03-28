using System;
using System.Collections.Generic;
using Etk.ModelManagement.Views;

namespace Etk.ModelManagement
{
    /// <summary> Model type definition </summary>
    public interface IModelType
    {
        IModelDefinitionManager Parent { get; }

        /// <summary>Model type name </summary>
        string Name { get; }
        /// <summary>Model type description</summary>
        string Description { get; }

        /// <summary>.Net Underlying type</summary>
        Type UnderlyingType { get; }

        /// <summary>Model type default view</summary>
        IEnumerable<IModelView> DefaultViews { get; }
        /// <summary>Get the model default views</summary>

        /// <returns>the model type properties</returns>
        IEnumerable<IModelProperty> GetProperties();

        ///// <summary>Get the model type property whose name is passed as parameter</summary>
        /// <param name="name">Name of the property to retrieve</param>
        ///// <returns>If exists, the model type property, if not null.</returns>
        IModelProperty GetProperty(string name);
    }
}
