namespace Etk.Tests.Data.Shops.DataType
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Xml.Serialization;

    public enum TestEnum
    { 
        EnumVal1,
        EnumVal2,
        EnumVal3,
        EnumVal4,
        EnumVal5,
    }

    [Serializable]
    public class Customer : INotifyPropertyChanged
    {
        private TestEnum enum1;
        [XmlIgnore]
        public TestEnum Enum1
        {
            get { return enum1; }
            set
            {
                enum1 = value;
                OnPropertyChanged("Enum1");
            }
        }

        private TestEnum? enumNullable1;
        [XmlIgnore]
        public TestEnum? EnumNullable1
        {
            get { return enumNullable1; }
            set
            {
                enumNullable1 = value;
                OnPropertyChanged("EnumNullable1");
            }
        }

        [XmlAttribute]
        public int Id
        { get; set; }

        private string forename;
        [XmlAttribute]
        public string Forename
        {
            get { return forename; }
            set
            {
                forename = value;
                OnPropertyChanged("Forename");
            }
        }

        private string surname;
        [XmlAttribute]
        public string Surname
        { 
            get {return surname;}
            set 
            {
                surname = value;
                OnPropertyChanged("Surname");
            } 
        }

        private string phoneNumber;
        [XmlAttribute]
        public string PhoneNumber
        {
            get { return phoneNumber; }
            set
            {
                phoneNumber = value;
                OnPropertyChanged("PhoneNumber");
            }
        }

        private string mobileNumber;
        [XmlAttribute]
        public string MobileNumber
        {
            get { return mobileNumber; }
            set
            {
                mobileNumber = value;
                OnPropertyChanged("MobileNumber");
            }
        }

        [XmlElement]
        public Address Address
        { get; set; }

        [XmlElement("OrderId")]
        public List<int> OrderIds
        { get; set; }


        //[XmlIgnore]
        //public List<Order> Orders
        //{ get; private set; }

        [XmlIgnore]
        public string Name
        { get { return string.Format("{0} {1}", Forename, Surname); } }

        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged(string propertyName)
        {
            if (this.PropertyChanged != null)
                this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion

        #region public methods
        public IEnumerable<Order> GetOrders()
        {
            return OrdersManager.GetOrders(OrderIds);
        }

        public IEnumerable<OrderLine> GetAllOrdersLines()
        {
            List<OrderLine> ret = new List<OrderLine>();

            IEnumerable<OrderLine> lines = OrdersManager.GetOrders(OrderIds).SelectMany(o => o.Lines);
            ret.AddRange(lines);
            ret.AddRange(lines);
            return ret;
        }
        #endregion
    }
}
