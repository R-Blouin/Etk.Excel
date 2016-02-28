namespace Etk.Excel.BindingTemplates.Controls.NamedRange
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Text;
    using Etk.BindingTemplates.Context;
    using Etk.BindingTemplates.Definitions.Binding;
    using Etk.BindingTemplates.Definitions.Decorators;
    using Etk.BindingTemplates.Definitions.EventCallBacks;
    using Etk.Excel.BindingTemplates.Definitions;
    using Etk.Excel.UI.Extensions;

    class ExcelBindingDefinitionNamedRange : IBindingDefinition
    {
        #region attributes and properties
        public const string NAMEDRANGE_TEMPLATE_PREFIX = "<NR";
        private const string POS_KEYWORD = "[POS]";
        private const string ALL_POS_KEYWORD = "[ALLPOS]";

        private bool usePos = false;
        private bool useAllPos = false;
        private string rootName;
        ExcelNamedRangeDefinition definition;
        private IBindingDefinition nameBindingDefinition;

        public IBindingDefinition NestedBindingDefinition
        { get; private set; }

        #region from IBindingDefinition
        public string Name
        { get { return NestedBindingDefinition != null ? NestedBindingDefinition.Name : string.Empty; } }

        public string Description
        { get { return NestedBindingDefinition != null ? NestedBindingDefinition.Description : string.Empty; } }

        public string BindingExpression
        {
            get { return NestedBindingDefinition == null ? null : NestedBindingDefinition.BindingExpression; }
        }

        public bool IsACollection
        {
            get { return NestedBindingDefinition == null ? false : NestedBindingDefinition.IsACollection; }
        }

        public bool IsReadOnly
        {
            get { return NestedBindingDefinition == null ? false : NestedBindingDefinition.IsReadOnly; }
        }

        public bool IsEnum
        {
            get { return NestedBindingDefinition == null ? false : NestedBindingDefinition.IsEnum; }
        }

        public bool IsNullable
        {
            get { return NestedBindingDefinition == null ? false : NestedBindingDefinition.IsNullable; }
        }

        public bool CanNotify
        {
            get { return NestedBindingDefinition == null ? false : NestedBindingDefinition.CanNotify; }
        }

        public bool IsOptional
        {
            get { return NestedBindingDefinition == null ? false : NestedBindingDefinition.IsOptional; }
        }

        public bool IsBoundWithData
        {
            get { return NestedBindingDefinition == null ? false : NestedBindingDefinition.IsBoundWithData; }
        }

        public Type BindingType
        {
            get { return NestedBindingDefinition == null ? null : NestedBindingDefinition.BindingType; }
        }

        public bool BindingTypeIsGeneric
        {
            get { return NestedBindingDefinition == null ? false : NestedBindingDefinition.BindingTypeIsGeneric; }
        }

        public Type BindingGenericType
        {
            get { return NestedBindingDefinition == null ? null : NestedBindingDefinition.BindingGenericType; }
        }

        public Type BindingGenericTypeDefinition
        {
            get { return NestedBindingDefinition == null ? null : NestedBindingDefinition.BindingGenericTypeDefinition; }
        }

        public BindingPartType PartType
        {
            get { return NestedBindingDefinition == null ? BindingPartType.BindingDefinition : NestedBindingDefinition.PartType; }
        }

        public Decorator DecoratorDefinition
        {
            get { return NestedBindingDefinition == null ? null : NestedBindingDefinition.DecoratorDefinition; }
        }

        public EventCallback OnSelection
        {
            get { return NestedBindingDefinition == null ? null : NestedBindingDefinition.OnSelection; }
        }

        public EventCallback OnClick
        {
            get { return NestedBindingDefinition == null ? null : NestedBindingDefinition.OnClick; }
        }

        public bool IsMultiLine
        {
            get { return NestedBindingDefinition == null ? false : NestedBindingDefinition.IsMultiLine; }
        }

        public double MultiLineFactor
        {
            get { return NestedBindingDefinition == null ? 0 : NestedBindingDefinition.MultiLineFactor; }
        }
        #endregion
        #endregion

        #region .ctors and factories}
        private ExcelBindingDefinitionNamedRange(ExcelNamedRangeDefinition definition, string rootName, IBindingDefinition nestedBindingDefinition, IBindingDefinition nameBindingDefinition)
        {
            NestedBindingDefinition = nestedBindingDefinition;
            this.definition = definition;
            this.rootName = rootName;
            this.nameBindingDefinition = nameBindingDefinition;
            if (nameBindingDefinition == null)
            {
                if (definition.Name.Contains(POS_KEYWORD))
                    usePos = true;
                if (definition.Name.Contains(ALL_POS_KEYWORD))
                    useAllPos = true;
            }
        }

        static public ExcelBindingDefinitionNamedRange CreateInstance(ExcelTemplateDefinitionPart templateDefinition, ExcelNamedRangeDefinition definition, IBindingDefinition nestedBindingDefinition)
        {
            try
            {
                string trimmedName = definition.Name.Trim();
                IBindingDefinition nameBindingDefinition = null;
                string rootName = null;
                int pos = trimmedName.IndexOf('{');
                if (pos != - 1)
                {
                    rootName = trimmedName.Remove(pos);
                    string expression = trimmedName.Substring(pos);
                    BindingDefinitionDescription bindingDefinitionDescription = BindingDefinitionDescription.CreateBindingDescription(expression, expression);
                    if(bindingDefinitionDescription != null && ! string.IsNullOrEmpty(bindingDefinitionDescription.BindingExpression))
                    {
                        if(bindingDefinitionDescription.BindingExpression.Contains(ALL_POS_KEYWORD) || bindingDefinitionDescription.BindingExpression.Contains(POS_KEYWORD))
                            throw new ArgumentException(string.Format("Cannot mixte the keywords '{0}' and '{1}' with binding dataAccessor", POS_KEYWORD, ALL_POS_KEYWORD));
                        nameBindingDefinition = BindingDefinitionFactory.CreateInstances(templateDefinition.Parent as ExcelTemplateDefinition, bindingDefinitionDescription);
                    }
                }
                return new ExcelBindingDefinitionNamedRange(definition, rootName, nestedBindingDefinition, nameBindingDefinition);
            }
            catch (Exception ex)
            {
                string message = string.Format("Cannot create create the named caller binding dataAccessor '{0}'. {1}", definition.Name, ex.Message);
                throw new EtkException(message, ex);
            }
        }

        static public ExcelNamedRangeDefinition RetrieveNamedRangeDefinition(string definition)
        {
            ExcelNamedRangeDefinition ret = null;
            try
            {
                ret = definition.Deserialize<ExcelNamedRangeDefinition>();
                if (string.IsNullOrWhiteSpace(ret.Name))
                    throw new ArgumentException("The 'Name' attribute is mandatory");
            }
            catch (Exception ex)
            {
                string message = string.Format("Cannot retrieve the named caller dataAccessor '{0}'. {1}", definition.EmptyIfNull(), ex.Message);
                throw new EtkException(message, ex);
            }
            return ret;
        }
        #endregion

        public IBindingContextItem ContextItemFactory(IBindingContextElement owner)
        {
            string name = null;
            IBindingContextItem nestedContextItem = NestedBindingDefinition == null ? null : NestedBindingDefinition.ContextItemFactory(owner);
            if (nameBindingDefinition != null)
            {
                object obj = nameBindingDefinition.ResolveBinding(owner.DataSource);
                name = rootName.EmptyIfNull() + (obj == null ? string.Empty : obj.ToString());
            }
            else
            {
                if (usePos)
                    name = definition.Name.Replace(POS_KEYWORD, "_" + owner == null ? string.Empty : owner.Index.ToString());
                else if(useAllPos)
                {
                    StringBuilder nameBuilder = new StringBuilder();
                    IBindingContextElement currentOwner = owner;
                    while(currentOwner != null)
                    {
                        nameBuilder.Insert(0, "_" + currentOwner.Index);
                        currentOwner =  currentOwner.ParentPart.ParentContext.Parent;
                    }
                    nameBuilder.Insert(0, definition.Name.Replace(ALL_POS_KEYWORD, string.Empty));
                    name = nameBuilder.ToString();
                }
                else 
                    name = definition.Name;
            }
            return new ExcelContextItemNamedRange(owner, name, this, nestedContextItem);
        }

        public object UpdateDataSource(object dataSource, object data)
        {
            return NestedBindingDefinition == null ? null : NestedBindingDefinition.UpdateDataSource(dataSource, data);
        }

        public object ResolveBinding(object dataSource)
        {
            return NestedBindingDefinition == null ? null : NestedBindingDefinition.ResolveBinding(dataSource);
        }

        public bool MustNotify(object dataSource, object source, PropertyChangedEventArgs args)
        {
            return NestedBindingDefinition != null && NestedBindingDefinition. MustNotify(dataSource, source, args);
        }

        public IEnumerable<INotifyPropertyChanged> GetObjectsToNotify(object dataSource)
        {
            return NestedBindingDefinition == null ? null : NestedBindingDefinition.GetObjectsToNotify(dataSource);
        }

        public bool IsSelected()
        {
            return false;
        }

        public bool IsDoubleLeftClicked()
        {
            return false;
        }
    }
}
