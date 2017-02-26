namespace Etk.Tests.Data.Shops.DataType
{
    using System;
    using System.ComponentModel;
    using System.Xml.Serialization;

    [Serializable]
    public class OrderLine : INotifyPropertyChanged
    {
        [XmlIgnore]
        public int OrderId
        { get; set; }

        [XmlAttribute]
        public int ProductId
        { get; set; }

        private int quantity;
        [XmlAttribute]
        public int Quantity
        {
            get { return quantity; }
            set
            {
                quantity = value;
                OnPropertyChanged("Quantity");
                OnPropertyChanged("Price");
                OnPropertyChanged("PriceFromFormula");
            }
        }

        [XmlIgnore]
        virtual public double Price
        {
            get { return Product == null ? 0 : Product.UnitPrice * quantity; }
        }

        [XmlIgnore]
        public bool UseFormulaForPriceFromFormula
        { get; set; }

        private double priceFromFormula;
        [XmlIgnore]
        public double PriceFromFormula
        {
            get 
            {
                if (!UseFormulaForPriceFromFormula)
                    return Product == null ? 0 : Product.UnitPrice * quantity;
                return priceFromFormula;
            }
            set
            {
                    priceFromFormula = value;
                    OnPropertyChanged("PriceFromFormula");
            }
        }

        private Product product;
        [XmlIgnore]
        public Product Product
        {
            get 
            { 
                if(product == null)
                    product = ProductsManager.Instance.GetProduct(ProductId);
                return product;
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
