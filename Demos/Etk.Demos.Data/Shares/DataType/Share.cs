using System.ComponentModel;
using System.Xml.Serialization;

namespace Etk.Demos.Data.Shares.DataType
{
    public class Share : INotifyPropertyChanged
    {
        [XmlAttribute]
        public string Code
        { get; set; }

        [XmlAttribute]
        public string Name
        { get; set; }

        [XmlAttribute]
        public string Currency
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
                OnPropertyChanged("Amount");
            }
        }

        private double last;
        [XmlAttribute]
        public double Last
        {
            get { return last; }
            set
            {
                last = value;
                OnPropertyChanged("Last");
                OnPropertyChanged("Amount");
            }
        }


        [XmlIgnore]
        public double Amount => last * quantity;

        private double amountRoundedToHundred;
        [XmlIgnore]
        public double AmountRoundedToHundred
        {
            get { return amountRoundedToHundred; }
            set
            {
                amountRoundedToHundred = value;
                OnPropertyChanged("AmountRoundedToHundred");
            }
        }


        #region .ctors
        public Share() {}
        #endregion

        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion
    }
}
