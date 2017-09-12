using System.Windows;
using Etk.BindingTemplates.Views;
using Etk.Excel.UI.Windows.SortAndFilter.ViewModels;

namespace Etk.Excel.UI.Windows.SortAndFilter
{
    /// <summary>
    /// Logique d'interaction pour SortAndFilerWindow.xaml
    /// </summary>
    public partial class SortAndFilerWindow //: MetroWindow
    {
        public SortAndFilerWindow(ITemplateView rootTemplateView)
        {
            if (System.Windows.Application.Current == null)
            {
                new System.Windows.Application();
                System.Windows.Application.Current.ShutdownMode = ShutdownMode.OnExplicitShutdown;
            }


            InitializeComponent();

            SortAndFilterViewModel viewModel = new SortAndFilterViewModel(rootTemplateView);
            this.DataContext = viewModel;
        }
    }
}
