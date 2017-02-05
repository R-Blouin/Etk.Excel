using System.Collections.Generic;
using System.Linq;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.SortAndFilter;

namespace Etk.BindingTemplates.Context
{
    class BindingContextPart : IBindingContextPart
    {
        public IBindingContext ParentContext
        { get; private set; }
        
        public ISortersAndFilters ExternalSorterAndFilter
        { get; private set; }

        public ITemplateDefinitionPart TemplateDefinitionPart
        { get; private set; }

        public IEnumerable<IBindingContextElement> Elements
        { get; private set; }

        public IEnumerable<IBindingContextElement> ElementsToRender
        { get; private set; }

        public BindingContextPartType PartType 
        { get; private set; }

        #region .ctors and factories
        private BindingContextPart(IBindingContext parent, ITemplateDefinitionPart templateDefinitionPart, BindingContextPartType partType)
        {
            ParentContext = parent;
            TemplateDefinitionPart = templateDefinitionPart;
            PartType = partType; 
        }

        public static BindingContextPart CreateBodyBindingContextPart(BindingContext parent, ITemplateDefinitionPart templateDefinitionPart, List<object> dataSourceAsList, ISortersAndFilters externalSorterAndFilter, ISortersAndFilters templatedSorterAndFilter)
        {
            BindingContextPart ret = new BindingContextPart(parent, templateDefinitionPart, BindingContextPartType.Body);
            ret.ExternalSorterAndFilter = parent.ExternalSortsAndFilters;

            int elementIndex = 0;
            // Bug => Closure !!!!
            // ret.Elements = dataSourceAsList.Select(ds => new BindingContextElement(ret, ds, elementIndex++));
            List<IBindingContextElement> tmpElements = new List<IBindingContextElement>();
            foreach (object obj in dataSourceAsList)
                tmpElements.Add(new BindingContextElement(ret, obj, elementIndex++));
            ret.Elements = tmpElements;

            // The sorterers and filters defined outside templates have the priority 
            if (parent.ExternalSortsAndFilters != null)
            {
                Dictionary<object, IBindingContextElement> contextItemByElement = ret.Elements.ToDictionary(e => e.Element, e => e);
                IEnumerable<object> elements = parent.ExternalSortsAndFilters.Execute(contextItemByElement.Keys) as IEnumerable<object>;
                ret.ElementsToRender = elements.Select(e => { IBindingContextElement el = null;
                                                              contextItemByElement.TryGetValue(e, out el);
                                                              return el;});
            }
            else
                ret.ElementsToRender = ret.Elements;

            // Manage the filters defined in the templates
            if (templatedSorterAndFilter != null)
            {
                Dictionary<object, IBindingContextElement> contextItemByElement = ret.ElementsToRender.ToDictionary(e => e.Element, e => e);
                IEnumerable<object> elements = templatedSorterAndFilter.Execute(contextItemByElement.Keys) as IEnumerable<object>;
                ret.ElementsToRender = elements.Select(e => { IBindingContextElement el = null;
                                                              contextItemByElement.TryGetValue(e, out el);
                                                              return el; });
            }
            return ret;
        }
        #endregion

        /// <summary>For the header/footer</summary>
        public static BindingContextPart CreateHeaderOrFooterBindingContextPart(IBindingContext parent, ITemplateDefinitionPart templateDefinitionPart, BindingContextPartType partType, object dataSource)
        {
            BindingContextPart ret = new BindingContextPart(parent, templateDefinitionPart, partType);

            ret.Elements = new BindingContextElement[] { new BindingContextElement(ret, dataSource, 0) };
            ret.ElementsToRender = ret.Elements.ToList();
            return ret;
        }

        public void Dispose()
        {
            if (Elements != null)
            {
                foreach (IBindingContextElement element in Elements)
                    element.Dispose();
                Elements = null;
                ElementsToRender = null;
            }
        }
    }
}
