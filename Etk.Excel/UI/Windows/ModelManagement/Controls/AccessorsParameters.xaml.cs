namespace Etk.Excel.UI.Windows.ModelManagement.Controls
{
    using ViewModels;
    using System.Windows.Controls;

    /// <summary>
    /// Logique d'interaction pour FilterOnValue.xaml
    /// </summary>
    public partial class Accessorsparameters : UserControl
    {
        public Accessorsparameters(AccessorsParametersViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }
    }
}
