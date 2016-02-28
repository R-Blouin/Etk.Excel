namespace Etk.Excel.UI.Windows.ModelManagement.Controls
{
    using Etk.Excel.UI.Windows.ModelManagement.ViewModels;
    using System.Windows.Controls;
    using System.Windows.Data;

    /// <summary>
    /// Logique d'interaction pour FilterOnValue.xaml
    /// </summary>
    public partial class RequestProperties : UserControl
    {
        public RequestProperties(RequestViewModel viewModel)
        {
            InitializeComponent();
            this.DataContext = viewModel;
            lstAccessors.ItemsSource = viewModel.Accessors;

            CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(lstAccessors.ItemsSource);
            if (view != null)
            {
                PropertyGroupDescription groupDescription = new PropertyGroupDescription("Parent");
                view.GroupDescriptions.Add(groupDescription);
            }
        }
    }
}
