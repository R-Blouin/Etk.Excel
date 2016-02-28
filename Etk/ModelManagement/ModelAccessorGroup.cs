namespace Etk.ModelManagement
{
    using Etk.ModelManagement.Definitions.XmlDefinition;
    using Etk.Excel.UI.Extensions;
    using Etk.Excel.UI.Log;
    using System;
    using System.Collections.Generic;

    class ModelAccessorGroup : IModelAccessorGroup
    {
        #region properties and attributes
        private ILogger log = Logger.Instance;

        /// <summary>Model definition manager that owned the accessor.</summary>
        public IModelDefinitionManager Parent
        { get; private set; }

        List<IModelAccessor> accessors;
        /// <summary>Model definition manager that owned the accessor.</summary>
        public IEnumerable<IModelAccessor> Accessors
        { get { return accessors; } }

        /// <summary>Accessor Name.</summary>
        public string Name
        { get; private set; }

        /// <summary>Accessor Description</summary>
        public string Description
        { get; private set; }
        #endregion

        #region .ctors and factories
        private ModelAccessorGroup()
        {
            accessors = new List<IModelAccessor>();
        }

        public static IModelAccessorGroup CreateInstance(ModelDefinitionManager parent, XmlModelAccessorGroup definition)
        {
            if (definition == null)
                return null;

            ModelAccessorGroup group = new ModelAccessorGroup();
            try
            {
                if (string.IsNullOrEmpty(definition.Name))
                    throw new EtkException("'Name' is mandatory");

                group.Parent = parent;
                group.Name = definition.Name.EmptyIfNull().Trim();
                group.Description = definition.Description.EmptyIfNull().Trim();

                if (definition.Accessors != null)
                {
                    foreach (XmlModelAccessor xmlAccessor in definition.Accessors)
                    {
                        //IModelAccessor accessor = CreateModelAccessors(modelAccessorDefinition);
                        IModelAccessor accessor = ModelAccessor.CreateInstance(group, xmlAccessor);
                        if (accessor != null)
                        {
                            group.accessors.Add(accessor);

                            if (parent.ModelAccessorByIdent.ContainsKey(accessor.Ident))
                                Logger.Instance.LogFormat(LogType.Warn, "The model accessor '{0}' is declared more than once.", accessor.Ident);
                            parent.ModelAccessorByIdent[accessor.Ident] = accessor;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new EtkException(string.Format("Cannot create 'Model Accessor Group' '{0}': {2}", definition.Name.EmptyIfNull(), ex.Message));
            }
            return group;
        }
        #endregion
    }
}
