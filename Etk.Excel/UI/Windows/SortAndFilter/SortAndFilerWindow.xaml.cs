namespace Etk.Excel.UI.Windows.SortAndFilter
{
    using BindingTemplate.SortAndFilter.ViewModels;
    using Etk.BindingTemplates.Views;
    using System.Windows;

    /// <summary>
    /// Logique d'interaction pour SortAndFilerWindow.xaml
    /// </summary>
    public partial class SortAndFilerWindow //: MetroWindow
    {
        public SortAndFilerWindow(ITemplateView rootTemplateView)
        {
            if (Application.Current == null)
            {
                new Application();
                Application.Current.ShutdownMode = ShutdownMode.OnExplicitShutdown;
            }


            InitializeComponent();

            SortAndFilterViewModel viewModel = new SortAndFilterViewModel(rootTemplateView);
            this.DataContext = viewModel;
        }
    }
}
