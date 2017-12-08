using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.BindingTemplates.Definitions.Decorators;
using Etk.Excel.BindingTemplates.Definitions;
using Etk.Tools.Extensions;
using Etk.BindingTemplates.Definitions.EventCallBacks;

namespace Etk.Excel.BindingTemplates.Controls.NamedRange
{
    class ExcelBindingDefinitionNamedRange : IBindingDefinition
    {
        #region attributes and properties
        public const string NAMEDRANGE_TEMPLATE_PREFIX = "<NR";
        private const string POS_KEYWORD = "[POS]";
        private const string ALL_POS_KEYWORD = "[ALLPOS]";

        private readonly bool usePos;
        private readonly bool useAllPos;
        private readonly string rootName;
        private readonly ExcelNamedRangeDefinition definition;
        private readonly IBindingDefinition nameBindingDefinition;

        public IBindingDefinition NestedBindingDefinition
        { get; private set; }

        #region from IBindingDefinition
        public string Name => NestedBindingDefinition != null ? NestedBindingDefinition.Name : string.Empty;

        public string Description => NestedBindingDefinition != null ? NestedBindingDefinition.Description : string.Empty;

        public string BindingExpression => NestedBindingDefinition?.BindingExpression;

        public bool IsACollection => NestedBindingDefinition?.IsACollection ?? false;

        public bool IsReadOnly => NestedBindingDefinition?.IsReadOnly ?? false;

        public bool IsEnum => NestedBindingDefinition?.IsEnum ?? false;

        public bool IsNullable => NestedBindingDefinition?.IsNullable ?? false;

        public bool CanNotify => NestedBindingDefinition?.CanNotify ?? false;

        public bool IsOptional => NestedBindingDefinition?.IsOptional ?? false;

        public bool IsBoundWithData => NestedBindingDefinition?.IsBoundWithData ?? false;

        public Type BindingType => NestedBindingDefinition?.BindingType;

        public bool BindingTypeIsGeneric => NestedBindingDefinition?.BindingTypeIsGeneric ?? false;

        public Type BindingGenericType => NestedBindingDefinition?.BindingGenericType;

        public Type BindingGenericTypeDefinition => NestedBindingDefinition?.BindingGenericTypeDefinition;

        public BindingPartType PartType => NestedBindingDefinition?.PartType ?? BindingPartType.BindingDefinition;

        public Decorator DecoratorDefinition => NestedBindingDefinition?.DecoratorDefinition;

        public EventCallback OnSelection => NestedBindingDefinition?.OnSelection;

        public EventCallback OnClick => NestedBindingDefinition?.OnClick;

        public bool IsMultiLine => NestedBindingDefinition?.IsMultiLine ?? false;

        public double MultiLineFactor => NestedBindingDefinition?.MultiLineFactor ?? 0;

        public EventCallback MultiLineFactorResolver => NestedBindingDefinition?.MultiLineFactorResolver;
        public SpecificEventCallback OnAfterRendering => NestedBindingDefinition?.OnAfterRendering;

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

        public static ExcelBindingDefinitionNamedRange CreateInstance(ExcelTemplateDefinitionPart templateDefinition, ExcelNamedRangeDefinition definition, IBindingDefinition nestedBindingDefinition)
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
                    BindingDefinitionDescription bindingDefinitionDescription = BindingDefinitionDescription.CreateBindingDescription(templateDefinition.Parent, expression, expression);
                    if(bindingDefinitionDescription != null && ! string.IsNullOrEmpty(bindingDefinitionDescription.BindingExpression))
                    {
                        if(bindingDefinitionDescription.BindingExpression.Contains(ALL_POS_KEYWORD) || bindingDefinitionDescription.BindingExpression.Contains(POS_KEYWORD))
                            throw new ArgumentException($"Cannot mixte the keywords '{POS_KEYWORD}' and '{ALL_POS_KEYWORD}' with binding dataAccessor");
                        nameBindingDefinition = BindingDefinitionFactory.CreateInstances(templateDefinition.Parent as ExcelTemplateDefinition, bindingDefinitionDescription);
                    }
                }
                return new ExcelBindingDefinitionNamedRange(definition, rootName, nestedBindingDefinition, nameBindingDefinition);
            }
            catch (Exception ex)
            {
                string message = $"Cannot create create the named caller binding dataAccessor '{definition.Name}'. {ex.Message}";
                throw new EtkException(message);
            }
        }

        public static ExcelNamedRangeDefinition RetrieveNamedRangeDefinition(string definition)
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
                string message = $"Cannot retrieve the named caller dataAccessor '{definition.EmptyIfNull()}'. {ex.Message}";
                throw new EtkException(message);
            }
            return ret;
        }
        #endregion

        public IBindingContextItem ContextItemFactory(IBindingContextElement owner)
        {
            string name;
            IBindingContextItem nestedContextItem = NestedBindingDefinition?.ContextItemFactory(owner);
            if (nameBindingDefinition != null)
            {
                object obj = nameBindingDefinition.ResolveBinding(owner.DataSource);
                name = rootName.EmptyIfNull() + (obj?.ToString() ?? string.Empty);
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
            return NestedBindingDefinition?.UpdateDataSource(dataSource, data);
        }

        public object ResolveBinding(object dataSource)
        {
            return NestedBindingDefinition?.ResolveBinding(dataSource);
        }

        public bool MustNotify(object dataSource, object source, PropertyChangedEventArgs args)
        {
            return NestedBindingDefinition != null && NestedBindingDefinition. MustNotify(dataSource, source, args);
        }

        public IEnumerable<INotifyPropertyChanged> GetObjectsToNotify(object dataSource)
        {
            return NestedBindingDefinition?.GetObjectsToNotify(dataSource);
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
