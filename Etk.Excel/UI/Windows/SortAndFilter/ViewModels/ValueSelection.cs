using Etk.Excel.UI.MvvmBase;

namespace Etk.Excel.UI.Windows.BindingTemplate.SortAndFilter.ViewModels
{
    class ValueSelection : ViewModelBase
    {
        public object Value
        { get; private set; }

        public string ValueString
        { get; private set; }

        private bool isSelected;
        public bool IsSelected
        { 
            get { return isSelected;}
            set
            {
                isSelected = value;
                OnPropertyChanged("IsSelected");
            }
        }

        public ValueSelection(object value)
        {
            Value = value;
            ValueString = value == null ? "<null>" : value.ToString();
        }
    }
}
