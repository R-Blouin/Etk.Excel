
using Etk.Excel.UI.MvvmBase;

namespace Etk.Tests.Templates.ExcelDna1.Tests.BasicValueTypes.DataTypes
{
    public enum TestEnum
    {
        enum1,
        enum2,
        enum3
    }

    public class BasicValueTypeContainer : ViewModelBase
    {
        private bool testBool;
        public bool TestBool
        {
            get { return testBool; }
            set
            {
                testBool = value;
                OnPropertyChanged("TestBool");
            }
        }

        private sbyte testByte;
        public sbyte TestByte
        {
            get { return testByte; }
            set
            {
                testByte = value;
                OnPropertyChanged("TestByte");
            }
        }

        private sbyte testSByte;
        public sbyte TestSByte
        {
            get { return testSByte; }
            set
            {
                testSByte = value;
                OnPropertyChanged("TestSByte");
            }
        }

        private char testChar;
        public char TestChar
        {
            get { return testChar; }
            set
            {
                testChar = value;
                OnPropertyChanged("TestChar");
            }
        }

        private short testShort;
        public short TestShort
        {
            get { return testShort; }
            set
            {
                testShort = value;
                OnPropertyChanged("TestShort");
            }
        }

        private ushort testUShort;
        public ushort TestUShort
        {
            get { return testUShort; }
            set
            {
                testUShort = value;
                OnPropertyChanged("TestUShort");
            }
        }

        private int testInt;
        public int TestInt
        {
            get { return testInt; }
            set
            {
                TestInt = value;
                OnPropertyChanged("TestInt");
            }
        }

        public uint testUInt;
        public uint TestUInt
        {
            get { return testUInt; }
            set
            {
                testUInt = value;
                OnPropertyChanged("TestUInt");
            }
        }

        private long testLong;
        public long TestLong
        {
            get { return testLong; }
            set
            {
                testLong = value;
                OnPropertyChanged("TestLong");
            }
        }

        private ulong testULong;
        public ulong TestULong
        {
            get { return testULong; }
            set
            {
                testULong = value;
                OnPropertyChanged("TestULong");
            }
        }

        private double testDouble;
        public double TestDouble
        {
            get { return testDouble; }
            set
            {
                testDouble = value;
                OnPropertyChanged("TestDouble");
            }
        }

        private float testFloat;
        public float TestFloat
        {
            get { return testFloat; }
            set
            {
                testFloat = value;
                OnPropertyChanged("TestFloat");
            }
        }

        private decimal testDecimal;
        public decimal TestDecimal
        {
            get { return testDecimal; }
            set
            {
                testDecimal = value;
                OnPropertyChanged("TestDecimal");
            }
        }

        private TestEnum testEnum;
        public TestEnum TestEnum
        {
            get { return testEnum; }
            set
            {
                testEnum = value;
                OnPropertyChanged("TestEnum");
            }
        }

        private string testString;
        public string TestString
        {
            get { return testString; }
            set
            {
                testString = value;
                OnPropertyChanged("TestString");
            }
        }
    }
}
