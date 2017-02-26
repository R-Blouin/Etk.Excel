namespace Etk.Tests.Data.Shops
{
    using Etk.Tests.Data.Shops.DataType;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Xml.Serialization;
    
    public static class OrdersManager
    {       
        #region attributes and properties
        static private OrderList orderList;

        public static IEnumerable<Order> Orders
        {
            get
            {
                if (orderList == null && orderList.Orders != null)
                    return null;
                return orderList.Orders;
            }
        }
        #endregion

        #region .ctors
        /// <summary> It's ugly but it's JUST to have some test data.</summary>
        static OrdersManager()
        {
            CreateDefaultData();
        }
        #endregion

        #region public methods
        /// <summary>Return an order given its id</summary>
        /// <param name="id">Order id to retrieve</param>
        public static Order GetOrder(int id)
        {
            if (orderList == null && orderList.Orders != null)
                return null;
            return orderList.Orders.FirstOrDefault(o => o.Id == id);
        }

        /// <summary>Return a list of specific orders</summary>
        /// <param name="ids">the order ids to retrieve</param>
        public static IEnumerable<Order> GetOrders(IEnumerable<int> ids)
        {
            if (orderList == null && orderList.Orders != null)
                return null;
            if (ids == null || ! ids.Any())
                return null;
            return orderList.Orders.Where(o => ids.Contains(o.Id));
        }
        #endregion

        #region private methods
        static private void CreateDefaultData()
        {
            XmlSerializer xs = new XmlSerializer(typeof(OrderList));
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Etk.Tests.Data.Shops.Data.Orders.xml"))
            {
                orderList = xs.Deserialize(stream) as OrderList;
            }
        }
        #endregion
    }
}
