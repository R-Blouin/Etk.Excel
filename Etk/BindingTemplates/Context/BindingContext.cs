using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.BindingTemplates.Views;
using Etk.SortAndFilter;

namespace Etk.BindingTemplates.Context
{
    public class BindingContext : IBindingContext
    {
        public ITemplateView Owner
        { get; protected set; }

        public ITemplateDefinition TemplateDefinition
        { get; protected set; }

        public IBindingContextElement Parent
        { get; private set; }

        //public int Occurrences
        //{ get; private set; }

        public object DataSource
        { get; private set; }

        public ISortersAndFilters ExternalSortsAndFilters
        { get; protected set; }

        public IBindingContextPart Header
        { get; private set; }

        public IBindingContextPart Body
        { get; private set; }

        public IBindingContextPart Footer
        { get; private set; }

        public List<IFilterDefinition> TemplatedFilters
        { get; private set; }

        #region .ctors
        public BindingContext(IBindingContextElement parent, ITemplateView owner, ITemplateDefinition templateDefinition, object dataSource, List<IFilterDefinition> templatedFilters)
        {
            try
            {
                if (owner == null)
                    throw new ArgumentNullException("The parameter 'owner' cannot be null");
                if (templateDefinition == null)
                    throw new ArgumentNullException("The parameter 'templateDefinition' cannot be null");

                Owner = owner;
                TemplateDefinition = templateDefinition;
                TemplatedFilters = templatedFilters;

                //TemplatedSortsAndFilters = templatedSortsAndFilters;
                Parent = parent;
                DataSource = dataSource;
                
                if (DataSource != null)
                {
                    List<object> dataSourceAsList;
                    IBindingDefinition dataSourceType;
                    if (DataSource is IEnumerable)
                    {
                        dataSourceAsList = (DataSource as IEnumerable).Cast<object>().ToList();
                        dataSourceType = BindingDefinitionRoot.CreateInstance(dataSourceAsList.GetType());
                    }
                    else
                    {
                        dataSourceAsList = new List<object>();
                        dataSourceAsList.Add(DataSource); //new object[] { DataSource };
                        dataSourceType = BindingDefinitionRoot.CreateInstance(DataSource.GetType());
                    }

                    if (TemplateDefinition.MainBindingDefinition != null)
                        CheckType(TemplateDefinition.MainBindingDefinition, dataSourceType);

                    ISortersAndFilters externalSortersAndFilters = null;
                    if (owner.ExternalSortersAndFilters != null)
                        owner.ExternalSortersAndFilters.TryGetValue(TemplateDefinition, out externalSortersAndFilters);

                    //Occurrences = dataSourceAsList.Count;
                    if (TemplateDefinition.Body != null)
                    {
                        IEnumerable<IFilterDefinition> templatedFiltersToTakeIntoAccount = null;
                        if (templatedFilters != null)
                        {
                            IEnumerable<IFilterDefinition>  templatedFiltersToTakeIntoAccountFound = templatedFilters.Where(tf => tf.TemplateDefinition == templateDefinition);
                            if (templatedFiltersToTakeIntoAccountFound.Any())
                                templatedFiltersToTakeIntoAccount = templatedFiltersToTakeIntoAccountFound;
                        }

                        ISorterDefinition[] sortersDefinition = null;
                        if (((TemplateView)owner).SorterDefinition != null && ((TemplateView)owner).SorterDefinition.TemplateDefinition == templateDefinition)
                            sortersDefinition =  new ISorterDefinition[] {((TemplateView)owner).SorterDefinition};

                        ISortersAndFilters sortersAndFilters = null;
                        if (templatedFilters != null || sortersDefinition != null)
                            sortersAndFilters = SortersAndFilterersFactory.CreateInstance(templateDefinition, templatedFiltersToTakeIntoAccount, sortersDefinition);
                        Body = BindingContextPart.CreateBodyBindingContextPart(this, TemplateDefinition.Body, dataSourceAsList, externalSortersAndFilters, sortersAndFilters);
                    }

                    if (TemplateDefinition.Header != null)
                        Header = BindingContextPart.CreateHeaderOrFooterBindingContextPart(this, TemplateDefinition.Header, DataSource);
                    if (TemplateDefinition.Footer != null)
                        Footer = BindingContextPart.CreateHeaderOrFooterBindingContextPart(this, TemplateDefinition.Footer, DataSource);
                }
            }
            catch (Exception ex)
            {
                string message = string.Format("Create the 'BindingContext' for template '{0}' failed . {1}", templateDefinition == null ? string.Empty : templateDefinition.Name, ex.Message);
                throw new EtkException(message);
            }
        }
        #endregion

        #region private methods
        private void CheckType(IBindingDefinition mainBindingDef, IBindingDefinition dataSourceType)
        {
            bool checkTypesOk = false;
            if (dataSourceType.IsACollection)
                checkTypesOk = true;
                //checkTypesOk = mainBindingDef.BindingType == dataSourceType.BindingType || mainBindingDef.BindingType.IsAssignableFrom(dataSourceType.BindingType);
            else
                checkTypesOk = mainBindingDef.BindingType == dataSourceType.BindingType || mainBindingDef.BindingType.IsAssignableFrom(dataSourceType.BindingType);

            if (! checkTypesOk)
            {
                Type expected = mainBindingDef.BindingType;
                Type got = dataSourceType.BindingType;
                throw new BindingTemplateException(string.Format("DataSource has not got the right UnderlyingType. '{0}' (or a UnderlyingType derivated from) was expected, got '{1}'", expected.Name, got.Name));
            }
        }
        #endregion

        #region #region public methods
        public void Dispose()
        {           
            if(Header != null)
            {
                Header.Dispose();
                Header = null;
            }
            if (Body != null)
            {
                Body.Dispose();
                Body = null;
            }
            if (Footer != null)
            {
                Footer.Dispose();
                Footer = null;
            }
        }
        #endregion
    }
}
