namespace Etk.ModelManagement
{
    using System;
    using System.Collections.Generic;

    public interface IModelDefinitionManager
    {
        /// <summary> If true (if at least one model is loaded), the IModelDefinitionManager can be used.</summary>
        bool HasModels { get; }

        /// <summary> Add model configuration data from file</summary>
        /// <param name="configurationFilePath">configurationj file path</param>
        void RegisterModelFromFile(string configurationFilePath);

        /// <summary> Add model configuration data from xml string</summary>
        /// <param name="xmlString">string containing the xml data</param>
        void RegisterModelFromXml(string xmlString);

        /// <summary> Return all the accessort group defined in the model</summary>
        /// <returns>The model accessorsNames</returns>
        IEnumerable<IModelAccessorGroup> GetAccessorGroups();

        ///// <summary> Return all the accessorsNames defined in the model</summary>
        ///// <returns>The model accessorsNames</returns>
        //IEnumerable<IModelAccessor> GetAccessors();

        ///// <summary> Return the model accessor having 'name' as name</summary>
        /// <param name="name">name of the accessorsNames to return</param>
        /// <returns>The accessorsNames having 'name' as name.</returns>
        IModelAccessor GetAccessor(string name);

        /// <summary>Return all the model types</summary>
        /// <returns>All the types defined in the model.</returns>
        IEnumerable<IModelType> GetModelTypes();

        /// <summary> Return the model type having 'name' as name </summary>
        /// <param name="name">Name of the model type to return.</param>
        /// <returns>The model type that has 'name' as name.</returns>
        IModelType GetModelType(string name);

        /// <summary> Add a model type to the model</summary> 
        /// <param name="type">The model type to add</param>
        void AddModelType(IModelType type);

        ///// <summary> Invoke (check) the model
        ///// Normally, the model parts are resolved at their first use: if the model is not correctly built, an exception is thrown.
        ///// Using this method, you can force the model to resolve (check) all its part.
        ///// </summary>
        ///// <returns>Return an exception if errors found.</returns>
        //void ResolveModel();
    }
}
