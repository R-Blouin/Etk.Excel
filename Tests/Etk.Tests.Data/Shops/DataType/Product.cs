namespace Etk.Tests.Data.Shops.DataType
{
    using System;
    using System.ComponentModel;
    using System.Xml.Serialization;

    [Serializable]
    public class Product : INotifyPropertyChanged
    {
        [XmlAttribute]
        public int Id
        { get; set;}

        [XmlAttribute]
        public string Name
        { get; set;}

        [XmlAttribute]
        public string Description
        { get; set; }

        private double unitPrice;
        [XmlAttribute]
        public double UnitPrice
        { 
            get {return unitPrice;}
            set
            {
                unitPrice = value;
                OnPropertyChanged("UnitPrice");
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
