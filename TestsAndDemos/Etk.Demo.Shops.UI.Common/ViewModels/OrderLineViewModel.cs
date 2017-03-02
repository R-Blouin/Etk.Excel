using System.ComponentModel;
using Etk.Tests.Data.Shops.DataType;

namespace Etk.Demo.Shops.UI.Common.ViewModels
{
    public class OrderLineViewModel : INotifyPropertyChanged
    {
        public OrderLine Line { get; private set; }

        public OrderLineViewModel(OrderLine line)
        {
            Line = line;
        }

        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged(string propertyName)
        {
            if (this.PropertyChanged != null)
                this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion
    }
}
