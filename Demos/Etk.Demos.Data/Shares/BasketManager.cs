using Etk.Demos.Data.Shares.DataType;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Etk.Demos.Data.Shares
{
    public class BasketManager : INotifyPropertyChanged
    {
        #region attributes and properties
        private static readonly Lazy<BasketManager> instance = new Lazy<BasketManager>(() => Activator.CreateInstance(typeof(BasketManager), true) as BasketManager);
        public static BasketManager Instance => instance.Value;

        private readonly Random random = new Random();
        public Basket Basket
        { get; set; }

        private volatile bool isRunning;
        public bool IsRunning
        {
            get { return isRunning; }
            private set
            {
                isRunning = value;
                OnPropertyChanged("CommandRunningLabel");
                OnPropertyChanged("IsRunning");
                OnPropertyChanged("IsNotRunning");
            }
        }

        public bool IsNotRunning => !isRunning;

        public string CommandRunningLabel => isRunning ? "Stop" : "Start";
        #endregion

        #region .ctors
        public BasketManager()
        {
            // Create Data
            XmlSerializer xs = new XmlSerializer(typeof(Basket));
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Etk.Demos.Data.Shares.Basket.xml"))
            {
                Basket = xs.Deserialize(stream) as Basket;
            }

        }
        #endregion

        #region public methods
        public void StartStopChanging()
        {
            if (Basket.Shares != null)
            {
                if (IsRunning)
                    IsRunning = false;
                else
                {
                    Task task = new Task(() => ChangeShares());
                    task.Start();
                }
            }
        }

        public void StartChanging()
        {
            if (Basket.Shares != null)
            {
                if (!isRunning)
                {
                    Task task = new Task(ChangeShares);
                    task.Start();
                }
            }
        }

        public void StopChanging()
        {
            if (Basket.Shares != null)
            {
                if (IsRunning)
                    IsRunning = false;
            }
        }
        #endregion

        #region private methods
        private void ChangeShares()
        {
            IsRunning = true;
            while(isRunning)
            {
                int index = random.Next(Basket.Shares.Count());
                Share share = Basket.Shares[index];
                if (share.CanChange)
                {
                    int percent = random.Next(11) - 5;
                    double value = share.Last + share.Last * percent / 100;
                    share.Last = value <= 0 ? 0.1 : value;

                    if (Basket.WaitingTime != -1)
                        Thread.Sleep(Basket.WaitingTime);
                }
            }
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
