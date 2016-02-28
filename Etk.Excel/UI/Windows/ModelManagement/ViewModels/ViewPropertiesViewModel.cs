namespace Etk.Excel.UI.Windows.ModelManagement.ViewModels
{
    using Etk.Excel.UI.MvvmBase;
    using Etk.ModelManagement;
    using Etk.ModelManagement.Views;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Wizard;

    public class ViewPropertiesViewModel : ViewModelBase, IWizardStep
    {
        private WizardViewModel parent;
        private IModelType returnModelType;

        public ModelView ModelView
        { get; private set; }

        private string name;
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

        private List<IModelProperty> sourceProperties;
        public List<IModelProperty> SourceProperties
        {
            get { return sourceProperties; }
            set
            {
                sourceProperties = value;
                OnPropertyChanged("SourceProperties");
                if (canNext != null)
                    canNext();
            }
        }

        public List<IModelProperty> selectedViewProperties;
        public List<IModelProperty> SelectedViewProperties
        {
            get { return selectedViewProperties; }
            set { selectedViewProperties = value; }
        }

        public ViewPropertiesViewModel(WizardViewModel parent, IModelView modelView)
        {
            this.parent = parent;
            SourceProperties = new List<IModelProperty>();
            returnModelType = parent.Request.SelectedAccessor.ReturnModelType;
            SourceProperties.AddRange(returnModelType.GetProperties());

            Name = returnModelType.Name;
            //rootModelType = modelType;
            //this.selectedViewProperties = new List<IModelProperty>();
            //if (selectedViewProperties == null && modelType.DefaultViews != null && modelType.DefaultViews.Any())
            //{
            //    IModelView defaultView = modelType.DefaultViews.FirstOrDefault(v => v.IsDefault);
            //    if(defaultView != null && defaultView)
            //    this.selectedViewProperties.AddRange(modelType.DefaultViews);
            //}
        }

        #region IWizardStep interface implementation
        public object GetNextStepData()
        {
            return false;
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
            return selectedViewProperties != null && selectedViewProperties.Any();
        }

        event Action canNext;
        event Action IWizardStep.CanNext
        {
            add { canNext += value; }
            remove { canNext -= value; }
        }
        #endregion
    }
}
