namespace Etk.Tests.Data.Shops.DataType
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Xml.Serialization;

    [Serializable]
    public class Shop : INotifyPropertyChanged
    {
        [XmlAttribute]
        public int Id
        { get; set; }

        private string name;
        [XmlAttribute]
        public string Name
        {
            get { return name; }
            set
            {
                name = value;
                OnPropertyChanged("Name");
            }
        }

        private string receptionPhone;
        [XmlAttribute]
        public string ReceptionPhone
        {
            get { return receptionPhone; }
            set
            {
                receptionPhone = value;
                OnPropertyChanged("ReceptionPhone");
            }
        }

        [XmlElement]
        public Address Address
        { get; set; }

        [XmlElement("CustomerId")]
        public List<int> CustomerIds
        { get; set; }

        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged(string propertyName)
        {
            if (this.PropertyChanged != null)
                this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion

        #region public methods
        public IEnumerable<Customer> GetCustomers()
        {
            return CustomerManager.GetCustomers(CustomerIds);
        }
        #endregion
    }
}
