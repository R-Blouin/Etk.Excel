using Etk.Tests.Data.Shops.DataType;
using System.Collections.Generic;
using System.Linq;

namespace Etk.Demo.Shops.UI.Common.Controls.ViewModels
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
    }
}