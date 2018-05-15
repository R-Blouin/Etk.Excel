namespace Etk.Tests.Data.Shops
{
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Xml.Serialization;
    using Etk.Tests.Data.Shops.DataType;
    
    public class ProductsManager
    {
        #region attributes and properties
        private static readonly object syncObj = new object();
        private ProductList productList;

        private static ProductsManager instance;
        public static ProductsManager Instance
        {
            get
            {
                lock (syncObj)
                {
                    if (instance == null)
                        instance = new ProductsManager();
                    return instance;
                }
            }
        }

        public IEnumerable<Product> Products
        {
            get
            {
                if (productList == null || productList.Products == null)
                    return null;
                return productList.Products;
            }
        }
        #endregion

        #region .ctors
        public ProductsManager()
        {
            CreateDefaultData();
        }
        #endregion

        #region public methods
        /// <summary>Return an product given its id</summary>
        /// <param name="id">Product id to retrieve</param>
        public Product GetProduct(int id)
        {
            if (productList == null && productList.Products != null)
                return null;
            return productList.Products.FirstOrDefault(o => o.Id == id);
        }

        ///// <summary>Return a list of specific products</summary>
        ///// <param name="ids">the product ids to retrieve</param>
        //public IEnumerable<Product> GetProducts(IEnumerable<int> ids)
        //{
        //    if (productList == null && productList.Products != null)
        //        return null;
        //    if (ids == null || !ids.Any())
        //        return null;
        //    return productList.Products.Where(o => ids.Contains(o.Id));
        //}
        #endregion

        #region private methods
        private void CreateDefaultData()
        {
            XmlSerializer xs = new XmlSerializer(typeof(ProductList));
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Etk.Tests.Data.Shops.Data.Products.xml"))
            {
                productList = xs.Deserialize(stream) as ProductList;
            }
        }
        #endregion
    }
}
