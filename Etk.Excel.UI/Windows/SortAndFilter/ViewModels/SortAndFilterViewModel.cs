using System;
using System.Collections.Generic;
using System.Windows.Input;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.BindingTemplates.Views;
using Etk.Excel.BindingTemplates.Views;
using Etk.Excel.MvvmBase;
using Etk.Excel.UI.MvvmBase;
using Etk.Excel.UI.Windows.BindingTemplate.SortAndFilter.ViewModels;
using Etk.SortAndFilter;

namespace Etk.Excel.UI.Windows.SortAndFilter.ViewModels
{
    class SortAndFilterViewModel : ViewModelBase, IDisposable
    {
        #region command
        RelayCommand filterSelectAllCommand;
        public ICommand FilterSelectAllCommand => filterSelectAllCommand ?? (filterSelectAllCommand = new RelayCommand(param => selectedDefinition?.ValueSelectionList?.ForEach(v => v.IsSelected = true))); 

        RelayCommand filterUnSelectAllCommand;
        public ICommand FilterUnSelectAllCommand => filterUnSelectAllCommand ?? (filterUnSelectAllCommand = new RelayCommand(param =>selectedDefinition?.ValueSelectionList?.ForEach(v => v.IsSelected = false)));

        RelayCommand applySortAndFilterCommand;

        public ICommand ApplySortAndFilterCommand => applySortAndFilterCommand ?? (applySortAndFilterCommand = new RelayCommand(param => ApplyExternalSortAndFilter()));

        RelayCommand resetAllCommand;
        public ICommand ResetAllCommand => resetAllCommand ?? (resetAllCommand = new RelayCommand(param => { ResetAllFiltersCommand.Execute(null);
                                                                                                             ResetAllSortersCommand.Execute(null);}));        

        RelayCommand resetAllFiltersCommand;
        public ICommand ResetAllFiltersCommand => resetAllFiltersCommand ?? (resetAllFiltersCommand = new RelayCommand(param => { TemplateViewModels.ForEach(t => t.BindingDefinitions.ForEach(b => { b.IsFilterWithConditions = false;
                                                                                                                                                                  b.IsFilterOnValues = false;}));}));

        RelayCommand resetAllSortersCommand;
        public ICommand ResetAllSortersCommand => resetAllSortersCommand ?? (resetAllSortersCommand = new RelayCommand(param => { TemplateViewModels.ForEach(t => t.BindingDefinitions.ForEach(b => { b.IsSortAscending = false;
                                                                                                                                                                  b.IsSortDescending = false;}));}));
        #endregion

        #region properties and attributes
        private readonly ITemplateView rootTemplateView;
        //private ISorterAndFilter sorterAndFilterer;

        private TemplateViewModel selectedTemplate;
        public TemplateViewModel SelectedTemplate
        {
            get { return selectedTemplate; }
            set
            {
                selectedTemplate = value;
                OnPropertyChanged("SelectedTemplate");
            }
        }

        public List<TemplateViewModel> TemplateViewModels
        { get; set; }

        private BindingDefinitionViewModel selectedDefinition;
        public BindingDefinitionViewModel SelectedDefinition
        {
            get { return selectedDefinition; }
            set
            {
                selectedDefinition = value;
                OnPropertyChanged("SelectedDefinition");
                FilterChanged();
            }
        }

        public bool SelectedDefinitionFilterOnValueEnabled => selectedDefinition != null && selectedDefinition.IsFilterOnValues;

        public bool SelectedDefinitionHasFilterOnValue => ! SelectedDefinitionHasFilterOnCondition;

        public bool SelectedDefinitionHasFilterOnCondition => selectedDefinition != null && selectedDefinition.IsFilterWithConditions;

        #endregion

        #region .ctors 
        public SortAndFilterViewModel(ITemplateView rootTemplateView)
        {
            this.rootTemplateView = rootTemplateView;

            TemplateViewModels = new List<TemplateViewModel>();
            PopulateTemplates(this.rootTemplateView.TemplateDefinition, new IBindingContext[] {this.rootTemplateView.BindingContext});
        }
        #endregion

        #region methods
        private void PopulateTemplates(ITemplateDefinition templateDefinition, IEnumerable<IBindingContext> bindingContexts)
        {
            //IEnumerable<IBindingContextElement> elements = bindingContexts.SelectMany(e => e.Elements)
            //                                                              .Where(e => e != null);
            //if(elements.Any())
            //{
            //    IEnumerable<IBindingContextItem> items = elements.SelectMany(e => e.BindingContextItems);
            //    if (templateDefinition.IsSortableAndFilterable)
            //    {
            //        TemplateViewModel templateViewModel = new TemplateViewModel(this, templateDefinition, items);
            //        TemplateViewModels.Add(templateViewModel);
            //    }

            //    //&&foreach (ILinkedTemplateDefinition childTemplate in templateDefinition.LinkedTemplates)
            //    //{
            //    //    IEnumerable<IBindingContext> childContexts = elements.SelectMany(e => e.LinkedBindingContexts)
            //    //                                                         .Where(e => e.DefinitionToFilterOwner.Name.Equals(childTemplate.DefinitionToFilterOwner.Name));

            //    //    if (bindingContexts.Any())
            //    //        PopulateTemplates(childTemplate.DefinitionToFilterOwner, childContexts);
            //    //}
            //}
        }

        private void ApplyExternalSortAndFilter()
        {
            Dictionary<ITemplateDefinition, ISortersAndFilters> sortersAndFilterers = new Dictionary<ITemplateDefinition, ISortersAndFilters>();
            foreach(TemplateViewModel template in TemplateViewModels)
            {
                ISortersAndFilters sorterAndFilterer = template.GetSorterAndFilterer();
                if (sorterAndFilterer != null)
                    sortersAndFilterers[template.TemplateDefinition] = sorterAndFilterer; 
            }
            if (sortersAndFilterers.Count > 0)
                rootTemplateView.ExternalSortersAndFilters = sortersAndFilterers;
            else
                rootTemplateView.ExternalSortersAndFilters = null;

            rootTemplateView.SetDataSource(rootTemplateView.GetDataSource());
            ETKExcel.TemplateManager.Render(rootTemplateView as IExcelTemplateView);
            //@à((RootTemplateView) rootTemplateView).RenderView();
            //MetroFormWpfContainer.CloseWindowCommandRouted.Execute(null, null);
        }

        public void FilterChanged()
        {
            OnPropertyChanged("SelectedDefinitionHasFilterOnValue");
            OnPropertyChanged("SelectedDefinitionFilterOnValueEnabled");
            OnPropertyChanged("SelectedDefinitionHasFilterOnCondition");
        }

        public void Dispose()
        {
            TemplateViewModels.ForEach(t => t.Dispose());
        }
        #endregion
    }
}
