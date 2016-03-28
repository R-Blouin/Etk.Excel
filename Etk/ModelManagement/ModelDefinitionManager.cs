using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using Etk.ModelManagement.Definitions.XmlDefinition;
using Etk.Tools.Extensions;
using Etk.Tools.Log;

namespace Etk.ModelManagement
{
    [Export]
    [PartCreationPolicy(CreationPolicy.Shared)]
    public class ModelDefinitionManager : IModelDefinitionManager
    {
        #region properties and attributes
        private readonly object syncObj = new object();

        internal Dictionary<string, IModelAccessorGroup> ModelAccessorGroupByName
        { get; set; }

        internal Dictionary<string, IModelAccessor> ModelAccessorByIdent
        { get; set; }

        internal Dictionary<string, IModelType> ModelTypeByName
        { get; set; }

        public bool  HasModels
        { get; protected set; }
        #endregion

        #region .ctors
        public ModelDefinitionManager()
        {
            ModelAccessorGroupByName = new Dictionary<string, IModelAccessorGroup>();
            ModelAccessorByIdent = new Dictionary<string, IModelAccessor>();
            ModelTypeByName = new Dictionary<string, IModelType>();
        }
        #endregion

        #region public methods
        /// <summary> Implements <see cref="IModelDefinitionManager.RegisterModelFromFile"/> </summary> 
        public void RegisterModelFromFile(string configurationFilePath)
        {
            try
            {
                Logger.Instance.LogFormat(LogType.Info, "Model configuration file '{0}' integration starting ...", configurationFilePath.EmptyIfNull());

                XmlModelConfiguration xmlModelConfiguration = XmlModelConfiguration.CreateInstanceFromFile(configurationFilePath);
                if (xmlModelConfiguration != null)
                    PopulateFromXmlModelConfiguration(xmlModelConfiguration);
            }
            catch (Exception ex)
            {
                throw new EtkException(string.Format("Retrieve model configuration from file '{0}' failed: {1}", configurationFilePath.EmptyIfNull(), ex.Message), ex);
            }
            finally
            {
                Logger.Instance.LogFormat(LogType.Info, "Model configuration file '{0}' integration finished", configurationFilePath.EmptyIfNull());
            }
        }

        /// <summary> Implements <see cref="IModelDefinitionManager.RegisterModelFromXml"/> </summary> 
        public void RegisterModelFromXml(string xmlString)
        {
            string configName = string.Empty;
            try
            {
                XmlModelConfiguration xmlModelConfiguration = XmlModelConfiguration.CreateInstanceFromXml(xmlString);
                if (xmlModelConfiguration != null)
                {
                    configName = xmlModelConfiguration.Name ?? string.Empty;
                    PopulateFromXmlModelConfiguration(xmlModelConfiguration);
                }
            }
            catch (Exception ex)
            {
                throw new EtkException(string.Format("Retrieve model configuration from xml '{0}' failed: {1}", configName, ex.Message), ex);
            }
            finally
            {
                Logger.Instance.LogFormat(LogType.Info, "Integration of model configuration xml '{0}' finished", configName);
            }
        }

        /// <summary> Implements <see cref="IModelDefinitionManager.GetAccessorGroups"/> </summary> 
        public IEnumerable<IModelAccessorGroup> GetAccessorGroups()
        {
            lock (syncObj)
            {
                List<IModelAccessorGroup> ret = null;
                if (ModelAccessorGroupByName != null)
                    ret = ModelAccessorGroupByName.Values.OrderBy(a => a.Name).ToList();
                return ret;
            }
        }

        /// <summary> Implements <see cref="IModelDefinitionManager.GetAccessors"/> </summary> 
        public IModelAccessor GetAccessor(string Ident)
        {
            lock (syncObj)
            {
                IModelAccessor ret = null;
                if (ModelAccessorByIdent != null)
                    ModelAccessorByIdent.TryGetValue(Ident, out ret);
                return ret;
            }
        }

        /// <summary> Implements <see cref="IModelDefinitionManager.GetModelTypes"/> </summary> 
        public IEnumerable<IModelType> GetModelTypes()
        {
            lock (syncObj)
            {
                return ModelTypeByName.Values.OrderBy(a => a.Name).ToList();
            }
        }

        /// <summary> Implements <see cref="IModelDefinitionManager.GetModelType"/> </summary> 
        public IModelType GetModelType(string name)
        {
            lock (syncObj)
            {
                if (string.IsNullOrEmpty(name))
                    return null;

                IModelType ret = null;
                ModelTypeByName.TryGetValue(name.ToUpper(), out ret);
                return ret;
            }
        }

        /// <summary> Add a model type from a .Net Type</summary> 
        public IModelType AddModelType(Type type)
        {
            if (type == null)
                return null;

            IModelType modelType = ModelType.CreateInstance(this, type);
            if (ModelTypeByName.ContainsKey(modelType.Name.ToUpper()))
                Logger.Instance.LogFormat(LogType.Warn, "The Model Type '{0}' is declared more than once.", modelType.Name);
            ModelTypeByName[modelType.Name.ToUpper()] = modelType;
            return modelType;
        }

        /// <summary> Implements <see cref="IModelDefinitionManager.AddModelType"/> </summary> 
        public void AddModelType(IModelType modelType)
        {
            if (modelType == null)
                return;

            if (ModelTypeByName.ContainsKey(modelType.Name.ToUpper()))
                Logger.Instance.LogFormat(LogType.Warn, "The Model Type '{0}' is register more than once.", modelType.Name);

            ModelTypeByName[modelType.Name.ToUpper()] = modelType;
        }

        ///// <summary> Implements <see cref="IModelDefinitionManager.GetModelType"/> </summary> 
        //public void ResolveModel()
        //{
        //    try
        //    {
        //        Logger.Instance.Log(LogType.Info, "Starting the model resolution...");

        //        bool onError = false;
        //        if (modelTypeByName != null)
        //        {
        //            foreach (IModelType type in modelTypeByName.Values)
        //            {
        //                try
        //                {
        //                    (type as ModelType).ResolveDependencies();
        //                }
        //                catch 
        //                { onError = true; }
        //            }
        //        }

        //        if (onError)
        //            throw new EtkException("Error(s) found");
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new EtkException(string.Format("Model resolution failed: {0}. Please check the log.", ex.Message));
        //    }
        //    finally
        //    {
        //        Logger.Instance.Log(LogType.Info, "Model resolution finished."); 
        //    }
        //}
        #endregion

        #region protected methods
        //protected abstract IModelAccessor CreateModelAccessors(XmlModelAccessor modelAccessorDefinition);

        protected void PopulateFromXmlModelConfiguration(XmlModelConfiguration xmlModelConfiguration)
        {
            lock (syncObj)
            {
                if (xmlModelConfiguration.ModelAccessorGroupDefinitions != null)
                {
                    HasModels = true;
                    foreach (XmlModelAccessorGroup xmlGroup in xmlModelConfiguration.ModelAccessorGroupDefinitions)
                    {
                        //IModelAccessor accessor = CreateModelAccessors(modelAccessorDefinition);
                        IModelAccessorGroup group = ModelAccessorGroup.CreateInstance(this, xmlGroup);
                        if (group != null)
                        {
                            if (ModelAccessorGroupByName.ContainsKey(group.Name))
                                Logger.Instance.LogFormat(LogType.Warn, "The model accessor group '{0}' is declared more than once.", group.Name);
                            ModelAccessorGroupByName[group.Name] = group;
                        }
                    }
                }

                if (xmlModelConfiguration.TypeDefinitions != null)
                {
                    foreach (XmlModelType typeDefinition in xmlModelConfiguration.TypeDefinitions)
                    {
                        ModelType modelType = ModelType.CreateInstance(this, typeDefinition);
                        if (modelType != null)
                        {
                            if (ModelTypeByName.ContainsKey(modelType.Name.ToUpper()))
                                Logger.Instance.LogFormat(LogType.Warn, "The model UnderlyingType '{0}' is declared more than once.", modelType.Name);
                            ModelTypeByName[modelType.Name.ToUpper()] = modelType;
                        }
                    }
                }

                // Resolve model accessorsNames dependencies
                foreach (ModelAccessor modelAccessor in ModelAccessorByIdent.Values)
                    modelAccessor.ResolveDependencies();

                // Resolve model types dependencies
                foreach (ModelType modelType in ModelTypeByName.Values.ToList())
                    modelType.ResolveDependencies();

                // Resolve views dependencies of the model types
                foreach (ModelType modelType in ModelTypeByName.Values.ToList())
                    modelType.ResolveViewsDependencies();
            }
        }
        #endregion
    }
}
