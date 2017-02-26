using Etk.Demo.Shops.UI.Common.Controls.ViewModels;
using System.Windows.Controls;

namespace Etk.Demo.Shops.UI.Common.Controls
{
    /// <summary>
    /// Logique d'interaction pour FilterOnValue.xaml
    /// </summary>
    public partial class ShopsControl : UserControl
    {
        public ShopsControl()
        {
            InitializeComponent();
            DataContext = new ShopsViewModel();
        }
    }
}
