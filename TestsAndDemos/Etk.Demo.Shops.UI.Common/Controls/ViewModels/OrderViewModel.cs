using Etk.Tests.Data.Shops.DataType;
using System.Collections.Generic;
using System.Linq;

namespace Etk.Demo.Shops.UI.Common.Controls.ViewModels
{
    public class OrderViewModel
    {
        public Order Order { get; private set; }
        
        public IEnumerable<OrderLineViewModel> Lines
        { get; private set; }

        public OrderViewModel(Order order)
        {
            Order = order;
            Lines = order.Lines.Select(l => new OrderLineViewModel(l)).ToArray();
        }
    }
}
