using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Decorators;
using Etk.BindingTemplates.Definitions.EventCallBacks;

namespace Etk.BindingTemplates.Definitions.Binding
{
    public abstract class BindingDefinition : IBindingDefinition
    {
        #region attributes and properties
        protected static Regex ValidCharExtract = new Regex("[a-zA-Z_]");

        protected BindingDefinitionDescription DefinitionDescription;

        public BindingPartType PartType => BindingPartType.BindingDefinition;

        /// <summary> Implements <see cref="IBindingDefinition.BindingExpression"/> </summary>
        public string BindingExpression => DefinitionDescription.BindingExpression;

        /// <summary> Implements <see cref="IBindingDefinition.IsACollection"/> </summary>
        public bool IsACollection
        { get; protected set; }

        //public bool IsALinqCollection
        //{ get; protected set; }

        /// <summary> Implements <see cref="IBindingDefinition.BindingType"/> </summary>
        public Type BindingType
        { get; protected set; }

        /// <summary> Implements <see cref="IBindingDefinition.BindingTypeIsGeneric"/> </summary>
        public bool BindingTypeIsGeneric
        { get; protected set; }

        /// <summary> Implements <see cref="IBindingDefinition.BindingGenericType"/> </summary>
        public Type BindingGenericType
        { get; protected set; }

        /// <summary> Implements <see cref="IBindingDefinition.BindingGenericTypeDefinition"/> </summary>
        public Type BindingGenericTypeDefinition
        { get; protected set; }

        /// <summary> Implements <see cref="IBindingDefinition.IsOptional"/> </summary>
        public Type CollectionType
        { get; protected set; }

        /// <summary> Implements <see cref="IBindingDefinition.IsReadOnly"/> </summary>
        public bool IsReadOnly
        { get; protected set; }

        /// <summary> Implements <see cref="IBindingDefinition.CanNotify"/> </summary>
        public bool CanNotify
        { get; protected set; }

        /// <summary> Implements <see cref="IBindingDefinition.IsOptional"/> </summary>
        public bool IsOptional
        { get; protected set; }

        /// <summary> Implements <see cref="IBindingDefinition.IsBoundWithData"/> </summary>
        public bool IsBoundWithData
        { get; protected set; }

        /// <summary> Implements <see cref="IBindingDefinition.Name"/> </summary>
        public virtual string Name
        { 
            get 
            {
                if (DefinitionDescription == null)
                    return string.Empty; 
                return string.IsNullOrEmpty(DefinitionDescription.Name) ? DefinitionDescription.BindingExpression : DefinitionDescription.Name; 
            } 
        }

        /// <summary> Implements <see cref="IBindingDefinition.Description"/> </summary>
        public virtual string Description
        { 
            get 
            {
                if (DefinitionDescription == null)
                    return string.Empty;
                return string.IsNullOrEmpty(DefinitionDescription.Description) ? Name : DefinitionDescription.Description; 
            } 
        }

        /// <summary> Implements <see cref="IBindingDefinition.DecoratorDefinition"/> </summary>
        public Decorator DecoratorDefinition => DefinitionDescription.Decorator;

        /// <summary> Implements <see cref="IBindingDefinition.IsNullable"/> </summary>
        public bool IsNullable
        { get; protected set; }

        /// <summary> Implements <see cref="IBindingDefinition.IsEnum"/> </summary>
        public bool IsEnum 
        { get; protected set;}

        /// <summary> Implements <see cref="IBindingDefinition.OnSelection"/> </summary>
        public EventCallback OnSelection => DefinitionDescription.OnSelection;

        /// <summary> Implements <see cref="IBindingDefinition.OnClick"/> </summary>
        public EventCallback OnClick => DefinitionDescription.OnLeftDoubleClick;

        /// <summary> Implements <see cref="IBindingDefinition.IsMultiLine"/> </summary>
        public bool IsMultiLine => DefinitionDescription.IsMultiLine;

        public double MultiLineFactor => DefinitionDescription.MultiLineFactor;

        public EventCallback MultiLineFactorResolver => DefinitionDescription.MultiLineFactorResolver;

        public SpecificEventCallback OnAfterRendering => DefinitionDescription.OnAfterRendering;

        public string Formula => DefinitionDescription.Formula;
        #endregion

        #region .ctors
        protected BindingDefinition(BindingDefinitionDescription bindingDefinitionDescription)
        {
            DefinitionDescription = bindingDefinitionDescription;
            CanNotify = false;
            IsBoundWithData = true;
            IsReadOnly = DefinitionDescription?.IsReadOnly ?? true;
        }
        #endregion

        protected void ManageCollectionStatus()
        {
            IsACollection = BindingType.GetInterfaces().Any(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(ICollection<>));
            if (IsACollection)
            {
                if (BindingType.IsArray)
                {
                    CollectionType = BindingType;
                    BindingType = BindingType.GetElementType();
                }
                else
                {
                    Type[] types = BindingType.GetGenericArguments();
                    if (types == null || types.Count() != 1)
                        throw new BindingTemplateException($"'{BindingType.FullName}': Only collection with one generic argument are taken into account.");
                    CollectionType = BindingType;
                    BindingType = types[0];
                }
            }
        }

        protected void ManageEnumAndNullable()
        {
            IsNullable = BindingType.IsGenericType && BindingType.GetGenericTypeDefinition() == typeof(Nullable<>);
            IsEnum = BindingType.IsGenericType ? BindingType.GetGenericArguments()[0].IsEnum : BindingType.IsEnum;
        }

        /// <summary> Implements <see cref="IBindingDefinition.DecoratorDefinition"/> </summary>
        public abstract object UpdateDataSource(object dataSource, object data);

        /// <summary> Implements <see cref="IBindingDefinition.ResolveBinding"/> </summary>
        public abstract object ResolveBinding(object dataSource);

        public virtual bool MustNotify(object dataSource, object source, PropertyChangedEventArgs args)
        {
            return false;
        }

        public virtual IEnumerable<INotifyPropertyChanged> GetObjectsToNotify(object dataSource)
        {
            return null;
        }

        public virtual IBindingContextItem ContextItemFactory(IBindingContextElement parent)
        {
            BindingContextItem ret;
            if (parent.DataSource == null)
                ret = new BindingContextItem(parent, this);
            else
                ret = CanNotify ? new BindingContextItemCanNotify(parent, this) : new BindingContextItem(parent, this);
            ret.Init();
            return ret;
        }
    }
}