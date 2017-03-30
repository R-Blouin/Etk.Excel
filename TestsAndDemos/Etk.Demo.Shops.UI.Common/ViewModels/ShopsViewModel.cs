using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Etk.Tests.Data.Shops;

namespace Etk.Demo.Shops.UI.Common.ViewModels
{
    public class ShopsViewModel : INotifyPropertyChanged
    {
        #region proeprties
        public IEnumerable<ShopViewModel> Shops
        { get; private set; }

        private ShopViewModel selectedShop;
        public ShopViewModel SelectedShop
        {
            get { return selectedShop; }
            set
            {
                selectedShop = value;
                OnPropertyChanged("SelectedShop");
                OnPropertyChanged("ShopsToDisplay");
            }
        }

        public IEnumerable<ShopViewModel> ShopsToDisplay
        {
            get { return selectedShop == Shops.First() ? Shops.Skip(1).ToArray() : new[] { selectedShop}; }
        }
        #endregion

        #region .ctors
        public ShopsViewModel()
        {
            List<ShopViewModel> shops = new List<ShopViewModel>();
            shops.Add(new ShopViewModel(null));
            shops.AddRange(ShopManager.Shops.Select(s => new ShopViewModel(s)).OrderBy(s => s.Name));

            Shops = shops;
            SelectedShop = Shops.First();
        }
        #endregion

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
