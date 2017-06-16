namespace Etk.Tests.Data.Shops
{
    using Etk.Tests.Data.Shops.DataType;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Xml.Serialization;
    
    public static class CustomerManager
    {
        #region attributes and properties
        static private CustomerList customerList;

        public static IEnumerable<Customer> Customers
        {
            get
            {
                if (customerList == null)
                    return null;
                return customerList.Customers;
            }
        }
        #endregion

        #region .ctors
        static CustomerManager()
        {
            CreateDefaultData();
        }
        #endregion

        #region private methods
        /// <summary> It's ugly but it's JUST to have some test data.</summary>
        private static void CreateDefaultData()
        {
            XmlSerializer xs = new XmlSerializer(typeof(CustomerList));
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Etk.Tests.Data.Shops.Data.Customers.xml"))
            {
                customerList = xs.Deserialize(stream) as CustomerList;
            }
        }
        #endregion

        #region public methods
        /// <summary>Return a list of specific customers</summary>
        /// <param name="ids">the customer ids to retrieve</param>
        public static IEnumerable<Customer> GetCustomers(IEnumerable<int> ids)
        {
            if (customerList == null)
                return null;
            if (ids == null || !ids.Any())
                return null;
            return customerList.Customers.Where(o => ids.Contains(o.Id));
        }

        public static int GetMaxOrdersLines()
        {
            return customerList.Customers.SelectMany(c => c.Orders).Max(o => o.Lines.Count);
        }

        /// <summary> Retrieve a customer by its Id</summary>
        /// <param name="customerId">Customer Id to retrieve</param>
        public static Customer GetCustomer(int customerId)
        {
            if (customerList == null)
                return null;
            return customerList.Customers.FirstOrDefault(c => c.Id == customerId);
        }

        /// <summary> Retrieve a specific order of a specific customer</summary>
        /// <param name="customerId">Customer Id to retrieve</param>
        /// <param name="orderId">Order Id to retrieve</param>
        public static Order GetCustomerOrder(int  customerId, int orderId)
        {
            if (customerList == null)
                return null;

            Order order = null;
            Customer customer = GetCustomer(customerId);
            if (customer != null)
            {
                IEnumerable<Order> orders = customer.Orders;
                if (orders != null)
                    order = orders.FirstOrDefault(o => o.Id == orderId);
            }
            return order;
        }

        /// <summary> Retrieve the orders of a specific customer. Return all orders if date is not set, if not, return all the orders for that date</summary>
        /// <param name="customerId">Customer Id to retrieve</param>
        /// <param name="date">date of the orders to retriev. Can be null</param>
        public static IEnumerable<Order> GetCustomerOrders(int customerId, DateTime? date)
        {
            if (customerList == null)
                return null;

            IEnumerable<Order> orders = null;
            Customer customer = GetCustomer(customerId);
            if (customer != null)
                orders = customer.Orders;
 
            if (orders != null && date.HasValue)
                orders = orders.Where(o => o.Date == date.Value);
            return orders;
        }
        #endregion
    }
}
