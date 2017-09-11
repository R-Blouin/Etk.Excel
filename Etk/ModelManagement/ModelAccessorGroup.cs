using System;
using System.Collections.Generic;
using Etk.ModelManagement.Definitions.XmlDefinition;
using Etk.Tools.Extensions;
using Etk.Tools.Log;

namespace Etk.ModelManagement
{
    class ModelAccessorGroup : IModelAccessorGroup
    {
        #region properties and attributes
        //private ILogger log = Logger.Instance;

        /// <summary>Model definition manager that owned the accessor.</summary>
        public IModelDefinitionManager Parent
        { get; private set; }

        readonly List<IModelAccessor> accessors;
        /// <summary>Model definition manager that owned the accessor.</summary>
        public IEnumerable<IModelAccessor> Accessors => accessors;

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
                throw new EtkException($"Cannot create 'Model Accessor Group' '{definition.Name.EmptyIfNull()}': {ex.Message}");
            }
            return group;
        }
        #endregion
    }
}
