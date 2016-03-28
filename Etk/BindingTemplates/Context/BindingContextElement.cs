using System;
using System.Collections.Generic;
using Etk.BindingTemplates.Context.SortSearchAndFilter;
using Etk.BindingTemplates.Definitions;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.BindingTemplates.Definitions.SortSearchAndFilter;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.BindingTemplates.Views;
using Etk.SortAndFilter;

namespace Etk.BindingTemplates.Context
{
    class BindingContextElement : IBindingContextElement
    {
        private bool disposed;

        public IBindingContextPart ParentPart
        { get; private set; }

        public object DataSource
        { get; private set; }

        public int Index
        { get; private set; }

        public List<IBindingContextItem> BindingContextItems
        { get; private set; }

        public List<IBindingContext> LinkedBindingContexts
        { get; private set; }

        public object Element
        { get; private set; }

        #region .ctors
        public BindingContextElement(IBindingContextPart parent, object dataSource, int index)
        {
            BindingContextItems = new List<IBindingContextItem>();
            LinkedBindingContexts = new List<IBindingContext>();

            ParentPart = parent;
            DataSource = dataSource;
            Index = index;

            Init();
        }

        ~BindingContextElement()
        {
            // Don't dispose from here: When call from here, Excel is exiting
            //Dispose();
        }
        #endregion

        #region private method
        private void Init()
        {
            try
            {
                bool createElement = ParentPart.TemplateDefinitionPart.Parent.BindingType != null && (ParentPart.ExternalSorterAndFilter != null || ParentPart.ParentContext.TemplatedFilters != null || ((TemplateView) ParentPart.ParentContext.Owner).SorterDefinition != null);
                if (createElement)
                    Element = Activator.CreateInstance(ParentPart.TemplateDefinitionPart.Parent.BindingType.BindType);

                List<BindingFilterContextItem> newTemplatesFilters = null; 
                foreach (IDefinitionPart definitionPart in ParentPart.TemplateDefinitionPart.BindingParts)
                {
                    if (definitionPart != null)
                    {
                        switch (definitionPart.PartType)
                        {
                            case BindingPartType.BindingDefinition:
                                IBindingContextItem item = ((IBindingDefinition) definitionPart).ContextItemFactory(this);
                                BindingContextItems.Add(item);
                            break;
                            case BindingPartType.FilterDefinition:
                                BindingFilterContextItem filter = ((BindingFilterDefinition) definitionPart).CreateContextItem(ParentPart.ParentContext.Owner, this);
                                if (newTemplatesFilters == null)
                                    newTemplatesFilters = new List<BindingFilterContextItem>();
                                BindingContextItems.Add(filter);
                                if (! string.IsNullOrEmpty(filter.FilterExpression))
                                    newTemplatesFilters.Add(filter);
                            break;
                            case BindingPartType.SearchDefinition:
                                BindingSearchContextItem search = ((BindingSearchDefinition) definitionPart).CreateContextItem(ParentPart.ParentContext.Owner);
                                BindingContextItems.Add(search);
                            break;
                        }
                    }
                }

//#if DEBUG
                LinkedBindingContexts = new List<IBindingContext>();
                foreach (ILinkedTemplateDefinition lt in ParentPart.TemplateDefinitionPart.LinkedTemplates)
                {
                    object resolvedBinding = lt.ResolveBinding(DataSource);
                    List<IFilterDefinition> templatedFilters = null;
                    if (newTemplatesFilters == null)
                        templatedFilters = ParentPart.ParentContext.TemplatedFilters;
                    else
                    {
                        templatedFilters = new List<IFilterDefinition>();
                        if (ParentPart.ParentContext.TemplatedFilters != null)
                            templatedFilters.AddRange(ParentPart.ParentContext.TemplatedFilters);
                        templatedFilters.AddRange(newTemplatesFilters);
                    }
                    BindingContext linkedContext = new BindingContext(this, ParentPart.ParentContext.Owner, lt.TemplateDefinition, resolvedBinding, templatedFilters);
                    LinkedBindingContexts.Add(linkedContext);
                }
//#else
//                IBindingContext[] contexts = new IBindingContext[ParentElement.FilterOwner.LinkedTemplates.Count];
//                if (contexts.Any())
//                {
//                    Parallel.For(0, ParentElement.DefinitionToFilterOwner.Body.LinkedTemplates.Count, i =>
//                    {
//                        ILinkedTemplateDefinition lt = ParentElement.DefinitionToFilterOwner.Body.LinkedTemplates[i];
//                        object resolvedBinding = lt.ResolveBinding(DataSource);
//                        contexts[i] = new BindingContext(this, lt.DefinitionToFilterOwner, resolvedBinding, externalSortsAndFilters);
//                    });
//                }
//                LinkedBindingContexts = contexts.ToList();
//#endif
            }
            catch (Exception ex)//@@
            {
                throw ex;
            }
        }
        #endregion

        #region public method
        public void Dispose()
        {
            if (!disposed)
            {
                disposed = true;
                foreach (IBindingContextItem item in BindingContextItems)
                {
                    if (item != null)
                        item.Dispose();
                }
                BindingContextItems.Clear();

                foreach (IBindingContext part in LinkedBindingContexts)
                    part.Dispose();
                LinkedBindingContexts.Clear();
            }
        }
        #endregion
    }
}
