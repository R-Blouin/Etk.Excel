namespace Etk.Demos.Data.Shops.DataType
{
    using System.ComponentModel;
    using System.Xml.Serialization;

    public class Address : INotifyPropertyChanged
    {
        private string street;
        [XmlAttribute]
        public string Street
        {
            get { return street; }
            set
            {
                street = value;
                OnPropertyChanged("Street");
            }
        }

        private string city;
        [XmlAttribute]
        public string City
        {
            get { return city; }
            set
            {
                city = value;
                OnPropertyChanged("City");
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
