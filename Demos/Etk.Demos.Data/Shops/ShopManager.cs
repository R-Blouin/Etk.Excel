using Etk.Demos.Data.Shops.DataType;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Serialization;

namespace Etk.Demos.Data.Shops
{
    public static class ShopManager
    {       
        #region attributes and properties
        static private ShopList shopList;

        public static IEnumerable<Shop> Shops
        {
            get
            {
                if (shopList == null)
                    return null;
                return shopList.Shops;
            }
        }
        #endregion

        #region .ctors
        static ShopManager()
        {
            CreateDefaultData();
        }
        #endregion

        #region private methods
        /// <summary>
        /// It's ugly but it's JUST to have some test data.
        /// </summary>
        static private void CreateDefaultData()
        {
            XmlSerializer xs = new XmlSerializer(typeof(ShopList));
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Etk.Demos.Data.Shops.Data.Shops.xml"))
            {
                shopList = xs.Deserialize(stream) as ShopList;
            }
        }
        #endregion

        #region public methods
        ///// <summary> Retrieve a shop by its Id</summary>
        ///// <param name="shopId">shop Id to retrieve</param>
        ///// <returns>A Customer having 'shopIdent' for id or null </returns>
        //static public Shop GetShop(int shopId)
        //{
        //    if (shopList == null && shopList.Shops == null)
        //        return null;
        //    return shopList.Shops.FirstOrDefault(c => c.Id == shopId);
        //}

        ///// <summary> Retrieve the customer of a specific shop.</summary>
        ///// <param name="shopId">shop Id to retrieve</param>
        //static public IEnumerable<Customer> GetShopCustomers(int shopId)
        //{
        //    if (shopList == null && shopList.Shops == null)
        //        return null;

        //    Shop shop = GetShop(shopId);
        //    if (shop == null)
        //        return null;

        //    return shop.Customers;
        //}
        #endregion
    }
}
