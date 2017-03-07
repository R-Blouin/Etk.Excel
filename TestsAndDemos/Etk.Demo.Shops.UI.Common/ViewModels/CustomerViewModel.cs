using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Etk.Tests.Data.Shops.DataType;

namespace Etk.Demo.Shops.UI.Common.ViewModels
{
    public class CustomerViewModel
    {
        public Customer Customer { get; set; }

        public IEnumerable<OrderViewModel> Orders
        { get; private set; }

        public CustomerViewModel(Customer customer)
        {
            Customer = customer;
            Orders = Customer.Orders.Select(o => new OrderViewModel(o)).ToArray();
        }

        public void DisplayName()
        {
            MessageBox.Show($"Yo !!! {Customer.Forename} {Customer.Surname}");
        }
    }
}