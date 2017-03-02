using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Etk.Tests.Data.Shops;
using Etk.Tests.Data.Shops.DataType;

namespace Etk.Demo.Shops.UI.Common.ViewModels
{
    public class ShopViewModel : INotifyPropertyChanged
    {
        private Shop shop;

        public int Ident
        { 
            get { return shop.Id; }
        }

        public string Name
        {
            get { return shop == null ? null : shop.Name; }
            set 
            {
                if(shop != null)
                    shop.Name = value;
                OnPropertyChanged("Name");
            }
        }

        public string Address
        {
            get { return shop.Address.Street; }
            set
            {
                shop.Address.Street = value;
                OnPropertyChanged("Address");
            }
        }

        public string City
        {
            get { return shop.Address.City; }
            set
            {
                shop.Address.City = value;
                OnPropertyChanged("City");
            }
        }

        public string Phone
        {
            get { return shop.ReceptionPhone; }
            set
            {
                shop.ReceptionPhone = value;
                OnPropertyChanged("Phone");
            }
        }

        public IEnumerable<CustomerViewModel> Customers
        {
            get;
            private set;
        }

        private bool? detailsVisibility;
        public bool? DetailsVisibility
        {
            get { return detailsVisibility;} 
            set { detailsVisibility = value;}
        }

        public ShopViewModel(Shop shop)
        {
            this.shop = shop;
            if(shop != null)
            {
                IEnumerable<Customer> customers = CustomerManager.GetCustomers(shop.CustomerIds);
                if (customers != null)
                    Customers = customers.Select(c => new CustomerViewModel(c)).ToArray();
            }
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
