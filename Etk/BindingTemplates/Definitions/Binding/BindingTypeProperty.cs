using System.Reflection;

namespace Etk.BindingTemplates.Definitions.Binding
{
    public class BindingTypeProperty
    {
        public string Name
        { get; private set; }

        public string Description
        { get; private set; }

        public MethodInfo GetMethod
        { get; private set; }

        public MethodInfo SetMethod
        { get; private set; }

        public BindingTypeProperty(string name, string description, MethodInfo getMethod, MethodInfo setMethod)
        {
            Name = name;
            Description = description;
            SetMethod = setMethod;
            GetMethod = getMethod;
        }
    }
}
