namespace Etk.Demos.Data.Shares.DataType
{
    using System.ComponentModel;
    using System.Xml.Serialization;

    public class Share : INotifyPropertyChanged
    {
        private bool canChange;
        [XmlIgnore]
        public bool CanChange
        {
            get { return canChange; }
            set
            {
                canChange = value;
                OnPropertyChanged("CanChange");
            }
        }

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
                if (CanChange)
                {
                    last = value;
                    OnPropertyChanged("Last");
                    OnPropertyChanged("Amount");
                }
            }
        }


        [XmlIgnore]
        public double Amount
        {
            get { return last * quantity; }
        }


        #region .ctors
        public Share()
        {
            CanChange = true;
        }
        #endregion

        #region public medthod
        public void ToogleCanChange()
        {
            CanChange = ! CanChange;    
        }
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
