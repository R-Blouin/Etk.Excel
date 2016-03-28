using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Etk.BindingTemplates.Definitions.Binding
{
    class BindingDefinitionHierarchical : BindingDefinition
    {
        private IBindingDefinition realBindingDefinition;
        private IBindingDefinition childBindingDefinition;

        #region override 'BindingDefinition' methods
        public override object UpdateDataSource(object dataSource, object data)
        {
            dataSource = realBindingDefinition.ResolveBinding( dataSource);
            return childBindingDefinition.UpdateDataSource(dataSource, data);
        }

        public override object ResolveBinding(object dataSource)
        {
            dataSource = realBindingDefinition.ResolveBinding(dataSource);
            return childBindingDefinition.ResolveBinding(dataSource);
        }

        public override IEnumerable<INotifyPropertyChanged> GetObjectsToNotify(object dataSource)
        {
            List<INotifyPropertyChanged> notifyPropertyChangedList = new List<INotifyPropertyChanged>();
            dataSource = realBindingDefinition.ResolveBinding(dataSource);
            return childBindingDefinition.GetObjectsToNotify(dataSource);
        }

        public override bool MustNotify(object dataSource, object source, PropertyChangedEventArgs args)
        {
            dataSource = realBindingDefinition.ResolveBinding(dataSource);
            return childBindingDefinition.MustNotify(dataSource, source, args);
        }
        #endregion

        #region .ctors and factories
        private BindingDefinitionHierarchical(BindingDefinitionDescription definitionDescription) : base(definitionDescription)
        { }

        public static BindingDefinitionHierarchical CreateInstance(Type type, BindingDefinitionDescription definitionDescription)
        {
            try
            {
                if (string.IsNullOrEmpty(definitionDescription.Name))
                    definitionDescription.Name = definitionDescription.BindingExpression.Replace('.', '_');

                BindingDefinitionHierarchical definition = new BindingDefinitionHierarchical(definitionDescription);

                string realBindingDefinitionExpression = definitionDescription.BindingExpression.Split('.')[0];

                BindingDefinitionDescription realBindingDefinitionDescription = new BindingDefinitionDescription() { BindingExpression = realBindingDefinitionExpression, IsReadOnly = definitionDescription.IsReadOnly };
                definition.realBindingDefinition = BindingDefinitionFactory.CreateInstance(type, realBindingDefinitionDescription);

                string childBindingExpression = definitionDescription.BindingExpression.Substring(realBindingDefinitionExpression.Length + 1);
                BindingDefinitionDescription childBindingDescription = new BindingDefinitionDescription() { BindingExpression = childBindingExpression, IsReadOnly = definitionDescription.IsReadOnly };
                definition.childBindingDefinition = BindingDefinitionFactory.CreateInstance(definition.realBindingDefinition.BindingType, childBindingDescription);

                definition.CanNotify = definition.childBindingDefinition.CanNotify;

                definition.BindingType = definition.childBindingDefinition.BindingType; // type;
                return definition;
            }
            catch (Exception ex)
            {
                throw new BindingTemplateException(string.Format("Cannot create the 'Hierarchical BindingDefinition' '{0}'. {1}", definitionDescription.BindingExpression, ex.Message));
            }
        }
        #endregion
    }
}
