using System;
using System.Collections.Generic;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Context.SortSearchAndFilter;
using Etk.BindingTemplates.Definitions.SortSearchAndFilter;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.SortAndFilter;

namespace Etk.BindingTemplates.Views
{
    public abstract class TemplateView : ITemplateView
    {
        #region attributes and properties
        protected object syncRoot = new object();

        private Dictionary<ITemplateDefinition, ISortersAndFilters> externalSortersAndFilters;
        public Dictionary<ITemplateDefinition, ISortersAndFilters> ExternalSortersAndFilters
        {
            get { return externalSortersAndFilters; }
            set{ externalSortersAndFilters = value;}
        }

        public Guid Ident
        { get; protected set; }

        public ITemplateView Parent
        { get; protected set; }

        public IBindingContext BindingContext
        { get; internal set; }

        public bool IsDisposed
        { get; protected set; }

        public ITemplateDefinition TemplateDefinition
        { get; protected set; }

        /// <summary>To keep a trace of the filters defined in templates</summary>
        public Dictionary<object, Dictionary<BindingFilterDefinition, string>> FilterValueByFilterDefinitionByElement
        { get; protected set; }

        /// <summary>Contains the template search value </summary>
        public abstract string SearchValue
        { get; set; }

        /// <summary>Contains the 'ISorterDefinition' attached to the view</summary>
        public ISorterDefinition SorterDefinition
        { get; set; }
        #endregion
        
        #region .ctors
        public TemplateView(ITemplateDefinition templateDefinition)
        {
            Ident = Guid.NewGuid();
            if (templateDefinition == null)
            {
                string message = "Cannot create a 'view' without a 'template dataAccessor'.";
                throw new EtkException(message);
            }
            TemplateDefinition = templateDefinition;
            FilterValueByFilterDefinitionByElement = new Dictionary<object, Dictionary<BindingFilterDefinition, string>>();
        }
        #endregion

        #region public methods
        public virtual void Clear()
        {
            SetDataSource(null);
        }

        public object GetDataSource()
        {
            return BindingContext == null ? null : BindingContext.DataSource;
        }

        public virtual void SetDataSource(object dataSource)
        {
            lock (syncRoot)
            {
                FilterValueByFilterDefinitionByElement.Clear();
                CreateBindingContext(dataSource);
            }
        }

        public virtual void CreateBindingContext(object dataSource)
        {
            try
            {
                if (BindingContext != null)
                {
                    BindingContext.Dispose();
                    BindingContext = null;
                }
                if (dataSource != null)
                {
                    List<IFilterDefinition> templatedFilters = RootBindingFilter.CreateInstances(this, dataSource);
                    BindingContext = new Context.BindingContext(null, this, this.TemplateDefinition, dataSource, templatedFilters);
                }
            }
            catch (Exception ex)
            {
                string message = string.Format("Binding template '{0}', 'SetDataSource' failed. {1}", TemplateDefinition.Name, ex.Message);
                throw new EtkException(message);
            }
        }

        //public void ApplyFilter()
        //{ }

        public virtual void Dispose()
        {
            lock (syncRoot)
            {
                if (!IsDisposed)
                {
                    if (BindingContext != null)
                        BindingContext.Dispose();
                    BindingContext = null;
                    IsDisposed = true;
                }
            }
        }

        public abstract void ExecuteSearch();
        #endregion
    }
}
