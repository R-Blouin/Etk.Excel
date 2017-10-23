using System;
using System.Collections.Generic;
using System.Linq;
using Etk.ModelManagement.Definitions.XmlDefinition;
using Etk.Tools.Extensions;

namespace Etk.ModelManagement
{
    public class ModelLinkProperty : IModelProperty 
    {
        #region properties and attributes
        public IModelType Parent
        { get; private set; }

        public string Name
        { get; set; }

        public string Description
        { get; set; }

        public IModelAccessor LinkedModelAccessor
        { get; private set; }

        public IEnumerable<string> Keys
        { get; private set; }

        public bool IsACollection
        { get; private set; }
        #endregion

        #region .ctors and factoriez
        public ModelLinkProperty(ModelType parent, XmlModelLinkProperty xmlLinkProperty)
        {
            try
            {
                Parent = parent;
                Name = xmlLinkProperty.Name.EmptyIfNull().Trim();
                Description = xmlLinkProperty.Description.EmptyIfNull().Trim();
                string accessorName = xmlLinkProperty.Accessor.EmptyIfNull().Trim();

                if (parent == null)
                    throw new EtkException("'Parent Type' is not defined");
                if (string.IsNullOrEmpty(Name))
                    throw new EtkException("'Name' cannot be null or empty");
                if (string.IsNullOrEmpty(accessorName))
                    throw new EtkException("'Accessor' cannot be null or empty");
                IModelAccessor accessor = Parent.Parent.GetAccessor(accessorName);
                if (accessor == null)
                    throw new EtkException($"Cannot find the 'Accessor' '{accessorName}'");

                LinkedModelAccessor = accessor;
                IsACollection = LinkedModelAccessor.ReturnTypeIsACollection;

                IEnumerable<string> keys = null;
                if (! string.IsNullOrEmpty(xmlLinkProperty.Keys))
                    keys = xmlLinkProperty.Keys.Split(',').Select(k => k.Trim()).Where(k => string.IsNullOrEmpty(k));
            }
            catch (Exception ex)
            {
                throw new EtkException($"Linked Property '{Name.EmptyIfNull()}' creation failed: {ex.Message}");
            }
        }
        #endregion
    }
}
