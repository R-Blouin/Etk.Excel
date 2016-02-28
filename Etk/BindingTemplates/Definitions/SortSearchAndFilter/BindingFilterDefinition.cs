namespace Etk.BindingTemplates.Definitions.SortSearchAndFilter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Etk.BindingTemplates.Context;
    using Etk.BindingTemplates.Context.SortSearchAndFilter;
    using Etk.BindingTemplates.Definitions.Binding;
    using Etk.BindingTemplates.Definitions.Templates;
    using Etk.BindingTemplates.Views;

    public abstract class BindingFilterDefinition : IDefinitionPart
    {
        #region attribuets and properties
        private string definition;
        private IEnumerable<string> propertyToFilterPath;

        public BindingPartType PartType
        {
            get { return BindingPartType.FilterDefinition; }
        }

        /// <summary>Template definition part that owns the filter</summary>
        public ITemplateDefinitionPart FilterOwner
        { get; protected set; }

        /// <summary>Name of the template property to filter</summary>
        public string DefinitionToFilterName
        { get; protected set; }

        /// <summary>Binding definition to filter</summary>
        public IBindingDefinition DefinitionToFilter
        { get; protected set; }

        /// <summary>Watermark</summary>
        public string Watermark
        { get; protected set; }
        #endregion

        #region .ctors
        protected BindingFilterDefinition(ITemplateDefinitionPart templateDefinitionPart, string definition, string watermark, IEnumerable<string> path)
        {
            FilterOwner = templateDefinitionPart;
            Watermark = watermark;
            this.definition = definition;
            DefinitionToFilterName = path.Last();
            DefinitionToFilterName = DefinitionToFilterName.Replace('.', '_');
            if (path.Count() > 1)
                propertyToFilterPath = path.Take(path.Count() - 1);
        }
        #endregion

        #region public methods
        public void Init()
        {
            if (FilterOwner.Parent.Body != null && FilterOwner.Parent.Body.BindingDefinitions != null)
            {
                // If the current template is the target, then filter its body elements
                if (propertyToFilterPath == null || propertyToFilterPath.Count() == 0)
                    DefinitionToFilter = FilterOwner.Parent.Body.BindingDefinitions.FirstOrDefault(def => def.IsBoundWithData && def.Name.Equals(DefinitionToFilterName));
                else
                {
                    if (FilterOwner.Parent.Body.LinkedTemplates != null)
                    {
                        ILinkedTemplateDefinition workingLinkedTemplateDefinition = FilterOwner.Parent.Body.LinkedTemplates.FirstOrDefault(lt => lt.TemplateDefinition.Name.Equals(propertyToFilterPath.First()));
                        if (propertyToFilterPath.Count() > 1)
                        {
                            string[] pathParts = propertyToFilterPath.Skip(1).ToArray();
                            for (int i = 0; i < pathParts.Length; i++)
                            {
                                if (workingLinkedTemplateDefinition == null)
                                    break;

                                List<ILinkedTemplateDefinition> links = new List<ILinkedTemplateDefinition>();
                                //if(workingLinkedTemplateDefinition.DefinitionToFilterOwner.Header != null && workingLinkedTemplateDefinition.DefinitionToFilterOwner.Header.LinkedTemplates != null)
                                //    links.AddRange(workingLinkedTemplateDefinition.DefinitionToFilterOwner.Header.LinkedTemplates);
                                if (workingLinkedTemplateDefinition.TemplateDefinition.Body != null && workingLinkedTemplateDefinition.TemplateDefinition.Body.LinkedTemplates != null)
                                    links.AddRange(workingLinkedTemplateDefinition.TemplateDefinition.Body.LinkedTemplates);
                                //if (workingLinkedTemplateDefinition.DefinitionToFilterOwner.Footer != null && workingLinkedTemplateDefinition.DefinitionToFilterOwner.Footer.LinkedTemplates != null)
                                //    links.AddRange(workingLinkedTemplateDefinition.DefinitionToFilterOwner.Footer.LinkedTemplates);

                                workingLinkedTemplateDefinition = links.FirstOrDefault(lt => lt.TemplateDefinition.Name.Equals(pathParts[i]));
                            }
                        }
                        if (workingLinkedTemplateDefinition != null)
                        {
                            List<IBindingDefinition> bindingDefinitions = new List<IBindingDefinition>();
                            //if(workingLinkedTemplateDefinition.DefinitionToFilterOwner.Header != null && workingLinkedTemplateDefinition.DefinitionToFilterOwner.Header.BindingDefinitions != null)
                            //    bindingDefinitions.AddRange(workingLinkedTemplateDefinition.DefinitionToFilterOwner.Header.BindingDefinitions);
                            if (workingLinkedTemplateDefinition.TemplateDefinition.Body != null && workingLinkedTemplateDefinition.TemplateDefinition.Body.BindingDefinitions != null)
                                bindingDefinitions.AddRange(workingLinkedTemplateDefinition.TemplateDefinition.Body.BindingDefinitions);
                            //if (workingLinkedTemplateDefinition.DefinitionToFilterOwner.Footer != null && workingLinkedTemplateDefinition.DefinitionToFilterOwner.Footer.BindingDefinitions != null)
                            //    bindingDefinitions.AddRange(workingLinkedTemplateDefinition.DefinitionToFilterOwner.Footer.BindingDefinitions);

                            DefinitionToFilter = bindingDefinitions.FirstOrDefault(def => def.Name.Equals(DefinitionToFilterName)); //&& def.IsBoundWithData
                            if (DefinitionToFilter != null)
                                FilterOwner = workingLinkedTemplateDefinition.TemplateDefinition.Body;
                        }
                    }
                }
            }
            if (DefinitionToFilter == null)
                throw new Exception(string.Format("Cannot resolve the path to the filter '{0}'.", definition));
        }

        public string GetFilterExpression(string filterValue)
        {
            if (string.IsNullOrEmpty(filterValue))
                return null;
            if (DefinitionToFilter.BindingType.IsValueType)
                return string.Format("{0}.ToString().ToUpper().Contains(\"{1}\")", DefinitionToFilter.Name, filterValue.ToUpper());
            else
                return string.Format("{0} != null && {0}.ToString().ToUpper().Contains(\"{1}\")", DefinitionToFilter.Name, filterValue.ToUpper());
        }

        abstract public BindingFilterContextItem CreateContextItem(ITemplateView view, IBindingContextElement parent);
        #endregion
    }
}
