namespace Etk.ModelManagement
{
    using DataAccessors;
    using Etk.Excel.UI.Extensions;
    using Etk.Excel.UI.Log;
    using Etk.ModelManagement.Definitions.XmlDefinition;
    using System;
    using System.Collections.Generic;
    using System.Linq;

    class ModelAccessor : IModelAccessor
    {
        private ILogger log = Logger.Instance;

        #region properties and attributes
        private string modelType;

        /// <summary>Model definition group that owned the accessor.</summary>
        public IModelAccessorGroup Parent
        { get; private set; }

        public string ParentName
        { get { return Parent.Name; } }

        /// <summary>Accessor Ident.</summary>
        public string Ident
        { get; private set; }

        /// <summary>Accessor Name.</summary>
        public string Name
        { get; private set; }

        /// <summary>Accessor Description</summary>
        public string Description
        { get; private set; }

        /// <summary> Name of the Model type returned by the 'DataAccessor' invocation.</summary>
        public IModelType ReturnModelType
        { get; private set; }

        /// <summary>Data Accessor (manage the data retrieving) </summary>
        public IDataAccessor DataAccessor
        { get; private set; }

        /// <summary>True is the return type is a collection of <see cref="IModelType"/> </summary>
        public bool ReturnTypeIsACollection
        { get; private set; }
        #endregion

        #region .ctors and factories
        private ModelAccessor()
        {}

        public static IModelAccessor CreateInstance(IModelAccessorGroup parent, XmlModelAccessor definition)
        {
            if (definition == null)
                return null;

            ModelAccessor accessor = new ModelAccessor();
            try
            {
                accessor.Parent = parent;
                accessor.Name = definition.Name.EmptyIfNull().Trim();
                accessor.Ident = definition.Ident.EmptyIfNull().Trim();
                if (string.IsNullOrEmpty(accessor.Ident))
                    accessor.Ident = accessor.Name;

                accessor.Description = definition.Description.EmptyIfNull().Trim();
                accessor.modelType = definition.ReturnModelType.EmptyIfNull().Trim();

                definition.InstanceName = definition.InstanceName.EmptyIfNull().Trim();
                definition.InstanceType = definition.InstanceType.EmptyIfNull().Trim();
                definition.Method = definition.Method.EmptyIfNull().Trim();

                definition.Method = definition.Method.EmptyIfNull().Trim();

                if (string.IsNullOrEmpty(accessor.Name))
                    throw new EtkException("'Name' is mandatory");
                if (string.IsNullOrEmpty(definition.Method))
                    throw new EtkException("'Method' is mandatory");

                DataAccessorInstanceType dataAccessorInstanceType = ModelManagement.DataAccessors.DataAccessor.AccessorInstanceTypeFrom(definition.InstanceType);
                accessor.DataAccessor = ModelManagement.DataAccessors.DataAccessor.CreateInstance(definition.Method, dataAccessorInstanceType, definition.InstanceName);
            }
            catch (Exception ex)
            {
                throw new EtkException(string.Format("Cannot create 'ModelAccessor' '{0} {1}': {2}", definition.Name.EmptyIfNull(), definition.Method.EmptyIfNull(), ex.Message));
            }
            return accessor;
        }
        #endregion

        #region public method
        public void ResolveDependencies()
        {
            try
            {
                Type accessorReturnType = DataAccessor.ReturnType.IsGenericType ? DataAccessor.ReturnType.GetGenericArguments()[0] : DataAccessor.ReturnType;
                ReturnTypeIsACollection = DataAccessor.ReturnType.GetInterfaces().Any(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(ICollection<>));

                string modelTypeToTest = string.IsNullOrEmpty(modelType) ? accessorReturnType.Name : modelType;

                ReturnModelType = this.Parent.Parent.GetModelType(modelTypeToTest);
                if (ReturnModelType != null)
                {
                    if (!ReturnModelType.UnderlyingType.Equals(accessorReturnType) || ! accessorReturnType.IsAssignableFrom(ReturnModelType.UnderlyingType))
                        throw new EtkException(string.Format("'Model type '{0}' underlying type '{1}' is not compatible with accessor return type '{2}'", modelTypeToTest, ReturnModelType.UnderlyingType.Name, accessorReturnType.Name));
                }
                else
                    ReturnModelType = ((ModelDefinitionManager)this.Parent.Parent).AddModelType(accessorReturnType);

                if(ReturnModelType == null)
                    throw new EtkException("Cannot retrieve the accessor return type");

                log.LogFormat(LogType.Info, "Model accessor '{0}' dependencies resolved", this.Ident);
            }
            catch(Exception ex)
            {
                throw new EtkException(string.Format("Cannot resolve dependencies of 'ModelAccessor' '{0}': {1}", Ident, ex.Message));
            }
        }
        #endregion
    }
}
