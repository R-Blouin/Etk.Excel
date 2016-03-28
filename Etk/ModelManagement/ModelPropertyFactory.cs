using System;
using System.Collections.Generic;
using System.Reflection;
using Etk.BindingTemplates.Definitions.Binding;

namespace Etk.ModelManagement
{
    class ModelPropertyFactory
    {
        public static IEnumerable<IModelProperty> CreateInstances(IModelType parent, Type type)
        {
            List<IModelProperty> ret = new List<IModelProperty>();
            if (type != null)
            {
                foreach (PropertyInfo pi in type.GetProperties())
                {
                    IBindingDefinition bindingDefinition = BindingDefinitionFactory.CreateInstance(pi);
                    if (bindingDefinition != null)
                    {
                        if (pi.PropertyType.Namespace.Equals("System"))
                            ret.Add(new ModelLeafProperty(parent, bindingDefinition));
                        else
                            ret.Add(new ModelProperty(parent, bindingDefinition));
                    }
                }
            }
            return ret;
        } 
    }
}
