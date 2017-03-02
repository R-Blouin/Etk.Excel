using System.Collections.Generic;
using System.Linq;
using Etk.Tests.Data.Shops.DataType;

namespace Etk.Demo.Shops.UI.Common.ViewModels
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
