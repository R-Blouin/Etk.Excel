using System;
using System.Collections.Generic;
using System.Linq;
using Etk.ModelManagement.Definitions.XmlDefinition;
using Etk.ModelManagement.Views;
using Etk.Tools.Extensions;
using Etk.Tools.Log;
using Etk.Tools.Reflection;

namespace Etk.ModelManagement
{
    /// <summary> Model type definition </summary>
    public class ModelType : IModelType
    {
        #region properties and attributes
        private readonly ILogger log = Logger.Instance;

        private readonly object syncObj = new object();
        private bool dependenciesResolved;

        /// <summary> Implements <see cref="IModelType.Name"/> </summary> 
        public string Name
        { get; private set; }

        /// <summary> Implements <see cref="IModelType.Description"/> </summary> 
        public string Description
        { get; private set; }

        /// <summary> Implements <see cref="IModelType.UnderlyingType"/> </summary> 
        public Type UnderlyingType
        { get; private set; }

        /// <summary> Implements <see cref="IModelType.ParentElement"/> </summary> 
        public IModelDefinitionManager Parent
        { get; private set; }

        public Dictionary<string, IModelProperty> PropertiesByName
        { get; private set; }

        public Dictionary<string, ModelLinkProperty> LinkPropertiesByName
        { get; private set; }

        private List<ModelView> defaultViews;
        public IEnumerable<IModelView> DefaultViews => defaultViews;

        #endregion

        #region .ctors and factories
        private ModelType()
        {}

        /// <summary> Create instance from an xml definition</summary>
        /// <param name="modelDefinitionManager">The owner of the future ModelType</param>
        /// <param name="definition">The xml defintion</param>
        public static ModelType CreateInstance(IModelDefinitionManager modelDefinitionManager, XmlModelType definition)
        {
            ModelType ret = null;
            if (definition != null)
            {
                ret = new ModelType();
                ret.Parent = modelDefinitionManager;
                ret.PropertiesByName = new Dictionary<string, IModelProperty>();
                ret.LinkPropertiesByName = new Dictionary<string, ModelLinkProperty>();
                try
                {
                    if (string.IsNullOrEmpty(definition.Name))
                        throw new EtkException("'Name' is mandatory");

                    ret.Name = definition.Name.EmptyIfNull().Trim();
                    ret.Description = definition.Description;

                    string typeName = definition.Type.EmptyIfNull().Trim();
                    if (string.IsNullOrEmpty(typeName))
                        throw new Exception("'Type' cannot be null or emtpy");

                    // if (!string.IsNullOrEmpty(typeName))
                    {
                        ret.UnderlyingType = TypeHelpers.GetType(typeName);
                        if (ret.UnderlyingType == null)
                            throw new Exception($"Cannot retrieve Underlying type '{typeName}'");
                        ret.RetrieveProperties(ret.UnderlyingType, definition.PropertiesToIgnore);
                    }

                    ret.OverridePropertiesFromXml(definition.Properties);
                    ret.RetrieveLinkPropertiesFromXml(definition.LinkProperties);
                    ret.RetrieveDefaultViewsFromXml(definition.Views);
                }
                catch (Exception ex)
                {
                    throw new EtkException($"Cannot create 'ModelType' '{definition.Name.EmptyIfNull()}': {ex.Message}");
                }
            }
            return ret;
        }

        /// <summary>Create a model type from a given Type  </summary>
        /// <param name="modelDefinitionManager"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public static ModelType CreateInstance(IModelDefinitionManager modelDefinitionManager, Type type)
        {
            ModelType ret = null;
            if (type != null)
            {
                ret = new ModelType();
                ret.Parent = modelDefinitionManager;
                ret.PropertiesByName = new Dictionary<string, IModelProperty>();
                try
                {
                    ret.Name = type.Name;
                    ret.Description = type.Name;
                    ret.UnderlyingType = type;

                    ret.RetrieveProperties(ret.UnderlyingType, null);
                }
                catch (Exception ex)
                {
                    throw new EtkException($"Cannot create 'ModelType' '{type.Name}': {ex.Message}");
                }
            }
            return ret;
        }
        #endregion

        #region public methods 
        /// <summary> Implements <see cref="IModelType.GetProperties"/> </summary> 
        public IEnumerable<IModelProperty> GetProperties()
        {
            if (! dependenciesResolved)
                ResolveDependencies();

            List<IModelProperty> ret = new List<IModelProperty>();
            if(PropertiesByName != null)
                ret.AddRange(PropertiesByName.Values.OrderBy(p => p.Name));
            if (LinkPropertiesByName != null)
                ret.AddRange(LinkPropertiesByName.Values.OrderBy(p => p.Name));
            return ret;
        }

        /// <summary> Implements <see cref="IModelType.GetProperty"/> </summary> 
        public IModelProperty GetProperty(string name)
        {
            if (!dependenciesResolved)
                ResolveDependencies();

            if (string.IsNullOrEmpty(name))
                return null;

            IModelProperty modelProperty;
            if (PropertiesByName.TryGetValue(name.ToUpper(), out modelProperty))
                return modelProperty;

            ModelLinkProperty modelLinkProperty;
            if (LinkPropertiesByName != null && LinkPropertiesByName.TryGetValue(name.ToUpper(), out modelLinkProperty))
                return modelLinkProperty;

            return null;
        }

        /// <summary> Invoke the dependencies used in this Type</summary>
        public void ResolveDependencies()
        {
            lock (syncObj)
            {
                if (! dependenciesResolved)
                {
                    try
                    {
                        if (PropertiesByName != null)
                        {
                            IEnumerable<ModelProperty> modelProperties = PropertiesByName.Values.Where(p => p is ModelProperty).Cast<ModelProperty>();
                            foreach (ModelProperty property in modelProperties)
                                property.ResolveDependencies(Parent);
                        }

                        //if (!string.IsNullOrEmpty(reference))
                        //    RetrievePropertiesFromReference();
                        dependenciesResolved = true;
                        //Resolvelinkedproperties();

                        log.LogFormat(LogType.Info, "Model type '{0}' dependencies resolved", this.Name);
                    }
                    catch (Exception ex)
                    {
                        throw new EtkException(
                            $"Cannot resolve the dependencies of the ModelType '{Name.EmptyIfNull()}': {ex.Message}");
                    }
                }            
            }
        }

        /// <summary> Invoke the dependencies of the views described in this Type</summary>
        public void ResolveViewsDependencies()
        {
            lock (syncObj)
            {
                try
                {
                    if (defaultViews != null)
                    {
                        foreach (ModelView view in defaultViews)
                            view.ResolveDependencies();
                    }
                }
                catch (Exception ex)
                {
                    throw new EtkException(
                        $"Cannot resolve the views of the ModelType '{Name.EmptyIfNull()}': {ex.Message}");
                }
            }
        }
        #endregion

        #region private methods
        /// <summary> Retrieve the bound properties from the underlying type</summary>
        private void RetrieveProperties(Type type, List<string> propertiesToIgnore)
        {
            IEnumerable<IModelProperty> properties = ModelPropertyFactory.CreateInstances(this, type);
            if (properties != null)
            {
                foreach (IModelProperty p in properties)
                {
                    if (propertiesToIgnore == null || ! propertiesToIgnore.Contains(p.Name))
                        PropertiesByName[p.Name.ToUpper()] = p;
                }
            }
        }

        /// <summary> If Xml properties are defined, they will override the definition of the existing ones.</summary>
        private void OverridePropertiesFromXml(List<XmlModelProperty> properties)
        {
            if (properties == null)
                return;
            try
            {
                foreach (XmlModelProperty property in properties)
                {
                    if (string.IsNullOrEmpty(property.Name))
                        throw new EtkException("A property name cannot be null or empty");

                    IModelProperty existingProperty;
                    // Override an existing model property definition
                    if(PropertiesByName.TryGetValue(property.Name.ToUpper(), out existingProperty))
                    {
                        existingProperty.Description = property.Description;
                        if(! string.IsNullOrEmpty(property.NameToUse))
                        {
                            PropertiesByName.Remove(existingProperty.Name.ToUpper());
                            existingProperty.Name = property.NameToUse;
                            PropertiesByName[existingProperty.Name.ToUpper()] = existingProperty;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new EtkException(
                    $"Cannot retrieve properties for UnderlyingType '{Name.EmptyIfNull()}': {ex.Message}");
            }
        }

        /// <summary> Retrieve the link properties from the xml definitions.</summary>
        private void RetrieveLinkPropertiesFromXml(List<XmlModelLinkProperty> xmlLinkProperties)
        {
            if (xmlLinkProperties == null)
                return;

            foreach (XmlModelLinkProperty xmlLinkProperty in xmlLinkProperties)
            { 
                ModelLinkProperty linkProperty = new ModelLinkProperty(this, xmlLinkProperty);
                if (linkProperty != null)
                    LinkPropertiesByName[linkProperty.Name.ToUpper()] = linkProperty;
            }
        }

        /// <summary> Retrieve the default properties from the xml definitions.</summary>
        private void RetrieveDefaultViewsFromXml(List<XmlModelView> xmlViews)
        {
            if (xmlViews == null || xmlViews.Count() == 0)
                return;

            defaultViews = new List<ModelView>();
            foreach (XmlModelView xmlView in xmlViews)
            {
                ModelView view = new ModelView(this, xmlView);
                defaultViews.Add(view);
            }
        }

        //private void RetrievePropertiesFromReference()
        //{
        //    try
        //    {
        //        IModelType typeReference = ParentElement.GetModelType(reference);
        //        if (typeReference == null)
        //            throw new EtkException(string.Format("Cannot retrieve the referenced model type '{0}'", reference));

        //        (typeReference as ModelType).ResolveDependencies();

        //        foreach (IModelProperty property in typeReference.GetProperties())
        //            BoundPropertiesByName[property.Name] = property;
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new EtkException(string.Format("Model type '{0}': Cannot resolve the model type reference:{1}", Name.EmptyIfNull(), ex.Message));
        //    }
        //}

        //private void Resolvelinkedproperties()
        //{
        //    //foreach (ModelLinkProperty property in LinkPropertiesByName.Values)
        //    //{
        //    //}
        //}
        #endregion
    }
}
