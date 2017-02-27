using Etk.Demo.Shops.UI.Common.Controls.ViewModels;
using MahApps.Metro.Controls;

namespace Etk.Demo.Shops.UI.Wpf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new ShopsViewModel();
        }
    }
}
