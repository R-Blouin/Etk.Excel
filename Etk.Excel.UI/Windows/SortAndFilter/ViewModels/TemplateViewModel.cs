using System;
using System.Collections.Generic;
using System.Linq;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.SortAndFilter;
using Etk.Excel.MvvmBase;

namespace Etk.Excel.UI.Windows.BindingTemplate.SortAndFilter.ViewModels
{
    class TemplateViewModel : ViewModelBase, IDisposable
    {
        #region attributes and properties
        public static event Action<TemplateViewModel, BindingDefinitionViewModel> BindingDefinitionSelected;

        private SortAndFilterViewModel parent;

        public ITemplateDefinition TemplateDefinition
        { get; private set; }

        private List<BindingDefinitionViewModel> bindingDefinitions = new List<BindingDefinitionViewModel>();
        public List<BindingDefinitionViewModel> BindingDefinitions
        { get { return bindingDefinitions; } }

        public string Name
        { get { return TemplateDefinition.Name; } }

        public string Description
        { get { return TemplateDefinition.Description; } }

        private BindingDefinitionViewModel selectedBindingDefinition;
        public BindingDefinitionViewModel SelectedBindingDefinition
        {
            get { return selectedBindingDefinition; } 
            set 
            {
                selectedBindingDefinition = value;
                if (selectedBindingDefinition != null)
                {
                    value.IsSortOrFilterSelected = true;
                    parent.SelectedTemplate = this;
                }
                OnPropertyChanged("SelectedBindingDefinition");
            } 
        }
        #endregion

        #region .ctors
        public TemplateViewModel(SortAndFilterViewModel parent, ITemplateDefinition templateDefinition, IEnumerable<IBindingContextItem> items)
        {
            this.parent = parent;
            //&&this.owner = owner;
            //DefinitionToFilterOwner = templateDefinition;
            //foreach (IBindingDefinition bindingDefinition in templateDefinition.BindingDefinitions.Where(b => b.IsBoundWithData))
            //{
            //    IEnumerable<IBindingContextItem> childItems = items.Where(e => e.BindingDefinition == bindingDefinition);
            //    if(childItems.Any())
            //        bindingDefinitions.Add(new BindingDefinitionViewModel(this, bindingDefinition, childItems));
            //}

            //BindingDefinitionSelected += OnBindingDefinitionSelected;
        }
        #endregion

        #region methods
        private void OnBindingDefinitionSelected(TemplateViewModel template, BindingDefinitionViewModel bindingDefinition)
        {
            if (this != template)
                this.SelectedBindingDefinition = null;
            else
                this.SelectedBindingDefinition = bindingDefinition;
        }

        public void BindingDefinitionSelectedRequest(TemplateViewModel template, BindingDefinitionViewModel bindingDefinition)
        {
            if (BindingDefinitionSelected != null)
                BindingDefinitionSelected(template, bindingDefinition);

            parent.SelectedDefinition = bindingDefinition;
        }

        public void FilterChanged()
        {
            parent.FilterChanged();
        }

        public void Dispose()
        {
            BindingDefinitionSelected -= OnBindingDefinitionSelected;
            bindingDefinitions.ForEach(b => b.Dispose());
        }

        public ISortersAndFilters GetSorterAndFilterer()
        {
            ISortersAndFilters sortAndFilter = null;

            IEnumerable<IFilterDefinition> filterElements = GetFilters();
            IEnumerable<ISorterDefinition> sorterElements = GetSorters();

            if (filterElements.Any() || sorterElements.Any())
                sortAndFilter = SortersAndFilterersFactory.CreateInstance(TemplateDefinition, filterElements, sorterElements);
            return sortAndFilter;
        }

        private IEnumerable<IFilterDefinition> GetFilters()
        {
            IEnumerable<BindingDefinitionViewModel> filters = bindingDefinitions.Where(b => b.IsFilterOnValues || b.IsFilterWithConditions);
            List<IFilterDefinition> elements = new List<IFilterDefinition>();
            if (filters.Count() > 0)
            {
                foreach (BindingDefinitionViewModel bindingDefinition in filters)
                {
                    if (bindingDefinition.IsFilterOnValues)
                    {
                        if (bindingDefinition.ValueSelectionList != null && bindingDefinition.ValueSelectionList.Count > 0)
                        {
                            IEnumerable<ValueSelection> selectedValues = bindingDefinition.ValueSelectionList.Where(v => v.IsSelected);
                            if (selectedValues.Count() == 0)
                                throw new EtkException(string.Format("{0} none values selected.", bindingDefinition.Description));

                            if (selectedValues.Count() != bindingDefinition.ValueSelectionList.Count)
                            {
                                bool OrEquals = selectedValues.Count() <= bindingDefinition.ValueSelectionList.Count / 2;
                                IFilterDefinition element = new FilterOnValues(null, bindingDefinition.BindingDefinition, selectedValues.Select(s => s.Value), OrEquals);
                                elements.Add(element);
                            }
                        }
                    }
                    else
                    { 
                        
                    }
                }
            }
            return elements;
        }

        private IEnumerable<ISorterDefinition> GetSorters()
        {
            List<ISorterDefinition> sortElements = new List<ISorterDefinition>();
            IEnumerable<BindingDefinitionViewModel> sorters = bindingDefinitions.Where(b => b.IsSortAscending || b.IsSortDescending);
            if (sorters.Any())
            {
                foreach (BindingDefinitionViewModel sorter in sorters)
                {
                    //ISorterDefinition sortElement = SortDefinitionFactory.CreateInstance(TemplateDefinition, sorter.BindingDefinition, sorter.IsSortDescending, sorter.IsNoCaseSensitive);
                    //sortElements.Add(sortElement);
                }
            }
            return sortElements;
        }
        #endregion
    }
}
