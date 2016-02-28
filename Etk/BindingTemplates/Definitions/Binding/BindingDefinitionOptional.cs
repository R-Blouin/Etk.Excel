namespace Etk.BindingTemplates.Definitions.Binding
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Text.RegularExpressions;
    using Etk.BindingTemplates.Context;

    class BindingDefinitionOptional : BindingDefinition
    {
        private Dictionary<Type, IBindingDefinition> bindingDefinitionByType = new Dictionary<Type, IBindingDefinition>();
        
        public override string Name
        { 
        //            //if (string.IsNullOrEmpty(definitionDescription.Name))
        //            //definitionDescription.Name = definitionDescription.BindingExpression.Replace('.', '_');
            get { return string.IsNullOrEmpty(DefinitionDescription.Name) ? DefinitionDescription.BindingExpression : DefinitionDescription.Name; }
        }

        #region .ctors and factories
        private BindingDefinitionOptional(BindingDefinitionDescription definitionDescription) : base(definitionDescription)
        { }
        
        static public BindingDefinitionOptional CreateInstance(BindingDefinitionDescription definitionDescription)
        {
            if (string.IsNullOrEmpty(definitionDescription.Name))
            {
                definitionDescription.Name = definitionDescription.BindingExpression.Replace('.', '_');
                MatchCollection ret = BindingDefinition.ValidCharExtract.Matches(definitionDescription.Name);
                StringBuilder sb = new StringBuilder();
                foreach (Match m in ret)
                    sb.Append(m.Value);
                definitionDescription.Name = sb.ToString();
            }
            BindingDefinitionOptional definition = new BindingDefinitionOptional(definitionDescription) { IsOptional = true };

            return definition;
        }
        #endregion

        #region public methods
        public override object UpdateDataSource(object dataSource, object data)
        {
            return null;
        }

        public override object ResolveBinding(object dataSource)
        {
            return null;
        }

        public IBindingDefinition CreateRealBindingDefinition(Type type)
        { 
            IBindingDefinition definition = null;
            if (!bindingDefinitionByType.TryGetValue(type, out definition))
            {
                definition = BindingDefinitionFactory.CreateInstance(type, DefinitionDescription) ?? this;
                bindingDefinitionByType[type] = definition;
            }
            return definition;
        }

        public override IBindingContextItem ContextItemFactory(IBindingContextElement parent)
        {
            BindingContextItem ret;
            if (parent.DataSource == null)
                ret = new BindingContextItem(parent, this);
            else
            {
                IBindingDefinition realBindingDefinition = CreateRealBindingDefinition(parent.DataSource.GetType());
                if (realBindingDefinition.CanNotify)
                    ret = new BindingContextItemCanNotify(parent, realBindingDefinition);
                else
                    ret = new BindingContextItem(parent, realBindingDefinition);
            }
            ret.Init();
            return ret;
        }
        #endregion
    }
}
