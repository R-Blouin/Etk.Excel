using System;
using System.Collections.Generic;
using System.Linq;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.Tools.Collections;
using Etk.Tools.Emit;
using System.Reflection;

namespace Etk.BindingTemplates.Definitions.Binding
{
    public class BindingType
    {
        #region properties and attributes
        //private static readonly object syncObj = new object();
        private static int classIdent;

        private static readonly TypeBuilderFactory typeBuilderFactory = new TypeBuilderFactory("BindedTypeAssembly");

        public Type BindType
        { get;  }

        public ReadOnlyDictionary<string, BindingTypeProperty> PropertyByName
        { get; }
        #endregion

        #region .ctors
        private BindingType(Type type,Dictionary<string, BindingTypeProperty> propertyByName)
        {
            BindType = type;
            PropertyByName = new ReadOnlyDictionary<string, BindingTypeProperty>(propertyByName);
        }
        #endregion

        #region factory
        public static BindingType CreateInstance(TemplateDefinition template)
        {
            BindingType bindingType = null;
            if (template != null)
            {
                List<IBindingDefinition> definitionsToUse = new List<IBindingDefinition>();
                if(template.Header != null)
                    definitionsToUse.AddRange(template.Header.BindingDefinitions.Where(d => ! (d is BindingDefinitionConstante)));
                if(template.Body != null)
                    definitionsToUse.AddRange(template.Body.BindingDefinitions.Where(d => ! (d is BindingDefinitionConstante)));
                if(template.Footer != null)
                    definitionsToUse.AddRange(template.Footer.BindingDefinitions.Where(d => ! (d is BindingDefinitionConstante)));

                if (definitionsToUse.Count > 0)
                {
                    List<EmitProperty> emitProperties = new List<EmitProperty>();
                    Dictionary<string, string> descriptionByName = new Dictionary<string, string>();
                    foreach (IBindingDefinition definition in definitionsToUse)
                    {
                        if (definition != null && ! string.IsNullOrEmpty(definition.Name))
                        {
                            emitProperties.Add(new EmitProperty(definition.Name, definition.BindingType ?? typeof(object)));
                            descriptionByName[definition.Name] = string.IsNullOrEmpty(definition.Description) ? definition.Name : definition.Description;
                        }
                    }

                    if (emitProperties.Count > 0)
                    {
                        Type type;
                        //lock (syncObj)
                        {
                            Dictionary<string, BindingTypeProperty> propertyByName = new Dictionary<string, BindingTypeProperty>();
                            type = typeBuilderFactory.CreateType($"BindType{classIdent++}", emitProperties);

                            foreach (PropertyInfo pi in type.GetProperties())
                            {
                                string name = pi.Name;
                                while (propertyByName.ContainsKey(name))
                                    name = name + "_";

                                propertyByName[name] = new BindingTypeProperty(name, descriptionByName[pi.Name], pi.GetGetMethod(), pi.GetSetMethod());
                            }
                            bindingType = new BindingType(type, propertyByName);
                        }
                    }
                }
            }
            return bindingType;
        }
        #endregion
    }
}
