using Etk.Demos.Data.Shops.DataType;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Serialization;

namespace Etk.Demos.Data.Shops
{
    public static class CustomersManager
    {
        #region attributes and properties
        private static CustomerList customerList;

        public static IEnumerable<Customer> Customers
        { get { return customerList?.Customers;}}
        #endregion

        #region .ctors
        static CustomersManager()
        {
            XmlSerializer xs = new XmlSerializer(typeof(CustomerList));
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Etk.Demos.Data.Shops.Data.Customers.xml"))
            {
                customerList = xs.Deserialize(stream) as CustomerList;

                customerList.Customers.AddRange(customerList.Customers);
                customerList.Customers.AddRange(customerList.Customers);
                customerList.Customers.AddRange(customerList.Customers);
                customerList.Customers.AddRange(customerList.Customers);
            }
        }
        #endregion
    }
}
