using System.Windows.Interop;
using Etk.BindingTemplates.Views;

namespace Etk.Excel.UI.Windows.SortAndFilter
{
    static class SortAndFilterManagement
    {
        public static void DisplaySortAndFilterWindow(ITemplateView templateView)
        {
            DisplaySortAndFilterWindow(null, templateView);
        }

        public static void DisplaySortAndFilterWindow(System.Windows.Forms.IWin32Window owner, ITemplateView templateView)
        {
            //SortAndFilterViewModel viewModel = new SortAndFilterViewModel(templateView);
            SortAndFilerWindow window = new SortAndFilerWindow(templateView);
            if (owner != null)
            {
                WindowInteropHelper windowInteropHelper = new WindowInteropHelper(window);
                windowInteropHelper.Owner = owner.Handle;
                window.ShowDialog();
            }
            else 
                window.ShowDialog();
        }
    }
}
