using System.Windows.Controls;
using Etk.Excel.UI.Windows.ModelManagement.ViewModels;

namespace Etk.Excel.UI.Windows.ModelManagement.Controls
{
    /// <summary>
    /// Logique d'interaction pour FilterOnValue.xaml
    /// </summary>
    public partial class ViewProperties : UserControl
    {
        public ViewProperties(ViewPropertiesViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }
    }
}
