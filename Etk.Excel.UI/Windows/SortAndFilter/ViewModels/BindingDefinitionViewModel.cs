using System;
using System.Collections.Generic;
using System.Linq;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.Excel.MvvmBase;
using Etk.Excel.UI.Windows.BindingTemplate.SortAndFilter.ViewModels;

namespace Etk.Excel.UI.Windows.SortAndFilter.ViewModels
{
    class BindingDefinitionViewModel : ViewModelBase, IDisposable
    {
        #region attributes and properties
        private static event Action<TemplateViewModel, BindingDefinitionViewModel> SortOrFilterSelected;

        private readonly TemplateViewModel parent;
        private readonly List<IBindingContextItem> items;

        public IBindingDefinition BindingDefinition
        { get; private set; }

        public string Description
        { 
            get 
            {
                if (string.IsNullOrEmpty(BindingDefinition.Description))
                    return BindingDefinition.Name;
                return $"{BindingDefinition.Name} ({BindingDefinition.Description})";
            } 
        }

        private bool isSortAscending;
        public bool IsSortAscending
        {
            get { return isSortAscending; }
            set 
            {
                isSortAscending = value;
                OnPropertyChanged("IsSortAscending");
                if (isSortAscending)
                {
                    isSortDescending = false;
                    OnPropertyChanged("IsSortDescending");
                }
                IsSortOrFilterSelected = true;
            }
        }

        private bool isSortDescending;
        public bool IsSortDescending
        {
            get { return isSortDescending; }
            set
            {
                isSortDescending = value;
                OnPropertyChanged("IsSortDescending");
                if (isSortAscending)
                {
                    isSortAscending = false;
                    OnPropertyChanged("IsSortAscending");
                }
                IsSortOrFilterSelected = true;
            }
        }

        private bool isFilterWithConditions;
        public bool IsFilterWithConditions
        {
            get { return isFilterWithConditions; }
            set 
            {
                isFilterWithConditions = value;
                OnPropertyChanged("IsFilterWithConditions");
                if (isFilterWithConditions)
                {
                    isFilterOnValues = false;
                    OnPropertyChanged("IsFilterOnValues");
                }
                parent.FilterChanged();
                IsSortOrFilterSelected = true;
            }
        }

        private bool isFilterOnValues;
        public bool IsFilterOnValues
        {
            get { return isFilterOnValues; }
            set
            {
                isFilterOnValues = value;
                OnPropertyChanged("IsFilterOnValues");
                if (isFilterOnValues)
                {
                    isFilterWithConditions = false;
                    OnPropertyChanged("IsFilterWithConditions");
                }
                parent.FilterChanged();
                IsSortOrFilterSelected = true;
            }
        }

        private bool isSortOrFilterSelected;
        public bool IsSortOrFilterSelected
        {
            get { return isSortOrFilterSelected; }
            set
            {
                if (isSortOrFilterSelected != value)
                {
                    isSortOrFilterSelected = value;
                    if (isSortOrFilterSelected)
                    {
                        SortOrFilterSelected(parent, this);
                        parent.BindingDefinitionSelectedRequest(parent, this);
                    }
                    OnPropertyChanged("IsSortOrFilterSelected");
                }
            }
        }

        private bool isNoCaseSensitive;
        public bool IsNoCaseSensitive
        {
            get { return isNoCaseSensitive; }
            set
            {
                isNoCaseSensitive = value;
                OnPropertyChanged("IsNoCaseSensitive");
            }
        }        

        public List<ValueSelection> ValueSelectionList
        { get; set;}
        #endregion

        #region .ctors
        public BindingDefinitionViewModel(TemplateViewModel parent, IBindingDefinition bindingDefinition, IEnumerable<IBindingContextItem> items)
        {
            this.parent = parent;
            BindingDefinition = bindingDefinition;
            this.items = items.ToList();

            if (items.Any())
            {
                ValueSelectionList = this.items.Select(i => i.ResolveBinding())
                                               .Distinct()
                                               .Select(v => new ValueSelection(v))
                                               .OrderBy(s => s.Value)
                                               .ToList();
            }

            SortOrFilterSelected += OnSortOrFilterSelected;
        }
        #endregion

        #region methods
        private void OnSortOrFilterSelected(TemplateViewModel parentTemplate,BindingDefinitionViewModel from)
        {
            if (from != this)
                IsSortOrFilterSelected = false;
        }

        public void Dispose()
        {
            SortOrFilterSelected -= OnSortOrFilterSelected;
        }
        #endregion
    }
}
