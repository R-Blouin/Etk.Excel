// http://stackoverflow.com/questions/1427471/observablecollection-not-noticing-when-item-in-it-changes-even-with-inotifyprop

using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;

namespace Etk.Tools.Collections
{
    public class TrulyObservableCollection<T> : ObservableCollection<T> where T : INotifyPropertyChanged
    {
        public TrulyObservableCollection() : base()
        {
            HookupCollectionChangedEvent();
        }

        public TrulyObservableCollection(IEnumerable<T> collection) : base(collection)
        {
            foreach (T item in collection)
                item.PropertyChanged += ItemPropertyChanged;

            HookupCollectionChangedEvent();
        }

        public TrulyObservableCollection(List<T> list) : base(list)
        {
            list.ForEach(item => item.PropertyChanged += ItemPropertyChanged);

            HookupCollectionChangedEvent();
        }

        private void ItemPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            NotifyCollectionChangedEventArgs a = new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset);
            OnCollectionChanged(a);
        }

        private void TrulyObservableCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.NewItems != null)
            {
                foreach (object item in e.NewItems)
                    (item as INotifyPropertyChanged).PropertyChanged += ItemPropertyChanged;
            }
            if (e.OldItems != null)
            {
                foreach (object item in e.OldItems)
                    (item as INotifyPropertyChanged).PropertyChanged -= ItemPropertyChanged;
            }
        }

        private void HookupCollectionChangedEvent()
        {
            CollectionChanged += TrulyObservableCollectionChanged;
        }
    }
}
