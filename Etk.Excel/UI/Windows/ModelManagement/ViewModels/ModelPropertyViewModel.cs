using Etk.Excel.UI.MvvmBase;
using Etk.ModelManagement;

namespace Etk.Excel.UI.Windows.ModelManagement.ViewModels
{
    class ModelPropertyViewModel : ViewModelBase
    {
        #region attributes and proeprties
        public bool IsSelected
        { get; set; }

        private string header;
        public string Header
        {
            get { return header; }
            set
            {
                header = value;
                OnPropertyChanged("Header");
            }
        }

        public ModelProperty ModelProperty
        { get; set; }

        public string Name
        { get { return ModelProperty.Name; } }

        public string Description
        { get { return ModelProperty.Description; } }
        #endregion

        #region .ctors
        public ModelPropertyViewModel(ModelProperty modelProperty, string header)
        {
            ModelProperty = modelProperty;
            Header = header;
        }
        #endregion

        #region IDropTarget implementation
        //public void DragOver(IDropInfo dropInfo)
        //{
        //}

        //public void Drop(IDropInfo dropInfo)
        //{
        //}
        #endregion
    }
}
