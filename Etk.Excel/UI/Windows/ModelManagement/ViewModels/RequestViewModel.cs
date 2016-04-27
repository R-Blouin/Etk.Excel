using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;
using Etk.Excel.UI.MvvmBase;
using Etk.Excel.UI.Windows.Wizard;
using Etk.ModelManagement;
using Etk.ModelManagement.Views;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.UI.Windows.ModelManagement.ViewModels
{
    public class RequestViewModel : ViewModelBase, IWizardStep
    {
        #region command
        private RelayCommand selectFirstOutputRange;
        /// <summary> Select the View first ouptup concernedRange</summary>
        public ICommand SelectFirstOutputRange
        {
            get
            {
                return selectFirstOutputRange ?? (selectFirstOutputRange = new RelayCommand(param =>
                        {
                            try
                            {
                                SelectingRange = true;

                                if (FirstOutputRange != null)
                                    FirstOutputRange.Select();
                                ExcelInterop.Range range = ETKExcel.ExcelApplication.RangeSelectionDialog("Select First output range");
                                if (range != null)
                                    FirstOutputRange = range;

                                caller.Select();
                            }
                            finally
                            {
                                SelectingRange = false;
                            }
                        }));
                    }
        }
        #endregion

        #region attributes and properties
        private WizardViewModel parent;

        private string name;
        /// <summary> View Name</summary>
        public string Name
        {
            get { return name; }
            set
            {
                name = value;
                OnPropertyChanged("Name");
                if (canNext != null)
                    canNext();
            }
        }

        private string description;
        /// <summary> View Description</summary>
        public string Description
        {
            get { return description; }
            set
            {
                description = value;
                OnPropertyChanged("Description");
            }
        }

        bool selectingRange;
        public bool SelectingRange
        {
            get { return selectingRange; }
            set
            {
                selectingRange = value;
                OnPropertyChanged("SelectingRange");
            }
        }

        private ExcelInterop.Range caller;
        private ExcelInterop.Range firstOutputRange;
        public ExcelInterop.Range FirstOutputRange
        {
            get { return firstOutputRange; }
            set
            {
                firstOutputRange = value;
                OnPropertyChanged("FirstOutputRangeAddress");
                if (canNext != null)
                    canNext();
            }
        }

        /// <summary> View first output concernedRange</summary>
        public string FirstOutputRangeAddress
        {
            get { return firstOutputRange == null ? string.Empty : string.Format("'{0}'!{1}", FirstOutputRange.Parent.Name, FirstOutputRange.AddressLocal); }
        }

        public IEnumerable<IModelAccessor> accessors;
        public IEnumerable<IModelAccessor> Accessors
        {
            get { return accessors; }
            set
            {
                accessors = value;
                OnPropertyChanged("Accessors");
            }
        }

        private IModelAccessor selectedAccessor;
        public IModelAccessor SelectedAccessor
        {
            get { return selectedAccessor; }
            set
            {
                selectedAccessor = value;
                OnPropertyChanged("SelectedAccessor");

                parent.ViewProperties = null;
                if (canNext != null)
                    canNext();
            }
        }

        private object accessorsSelectedItem;
        public object AccessorsSelectedItem
        {
            get { return accessorsSelectedItem; }
            set
            {
                if (value != null && value is IModelAccessor)
                {
                    accessorsSelectedItem = value;
                    SelectedAccessor = accessorsSelectedItem as IModelAccessor;
                }
                else
                {
                    accessorsSelectedItem = SelectedAccessor;
                    OnPropertyChanged("AccessorsSelectedItem");
                }
            }
        }

        public bool HasFilter
        { get { return !HasNoFilter; } }

        public bool HasNoFilter
        { get { return string.IsNullOrEmpty(FilterOnAccessors); } }

        private string filterOnAccessors;
        public string FilterOnAccessors
        {
            get { return filterOnAccessors; }
            set
            {
                filterOnAccessors = value;
                OnPropertyChanged("FilterOnAccessors");
                OnPropertyChanged("HasNoFilter");
                OnPropertyChanged("HasFilter");
                OnPropertyChanged("FilteredAccessors");
            }
        }

        /// <summary> Filtered accessors</summary>
        public IEnumerable<IModelAccessor> FilteredAccessors
        {
            get
            {
                if (Accessors == null || FilterOnAccessors == null)
                    return null;

                string filterOnAccessorsUpper = FilterOnAccessors.ToUpper();
                return Accessors.Where(a => a.Name != null && a.Name.ToUpper().Contains(filterOnAccessorsUpper))
                      .Union(Accessors.Where(a => a.Description != null && a.Description.ToUpper().Contains(filterOnAccessorsUpper)))
                      .Distinct()
                      .OrderBy(a => a.Name);
            }
        }
        #endregion

        #region .ctors
        public RequestViewModel(WizardViewModel parent, ExcelInterop.Range caller, ExcelInterop.Range firstOutputRange)
        {
            this.parent = parent;
            this.caller = caller;
            FirstOutputRange = firstOutputRange;
            IEnumerable<IModelAccessorGroup> accessorGroups = ETKExcel.ModelDefinitionManager.GetAccessorGroups();

            if (accessorGroups != null)
                Accessors = accessorGroups.SelectMany(g => g.Accessors).OrderBy(a => a.Name).ToList();
        }
        #endregion

        #region IWizardStep interface implementation
        public object GetNextStepData()
        {
            if (parent.ViewProperties == null)
            {
                IModelView modelView = null;
                if (selectedAccessor.ReturnModelType.DefaultViews != null && selectedAccessor.ReturnModelType.DefaultViews.Any())
                    modelView = selectedAccessor.ReturnModelType.DefaultViews.ElementAt(0);
                else
                    modelView = new ModelView(selectedAccessor.ReturnModelType, Name, Description);
                parent.ViewProperties = new ViewPropertiesViewModel(parent, modelView);
            }
            return parent.ViewProperties;
        }

        public bool OnNext(object parameters)
        {
            return true;
        }

        public bool OnCancel()
        {
            return true;
        }

        public bool CheckCanNext()
        {
            if (!string.IsNullOrEmpty(name) && selectedAccessor != null && firstOutputRange != null)
                return true;
            return false;
        }

        event Action canNext;
        event Action IWizardStep.CanNext
        {
            add{ canNext += value; }
            remove { canNext -= value; }
        }
        #endregion
    }
}
