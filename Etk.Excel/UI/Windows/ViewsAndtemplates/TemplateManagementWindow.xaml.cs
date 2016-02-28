namespace Etk.Excel.UI.Windows.ViewsAndtemplates
{
    using ViewModels;

    /// <summary>
    /// Logique d'interaction pour TemplateManagementWindow.xaml
    /// </summary>
    public partial class TemplateManagementWindow
    {
        private TemplateManagementViewModel viewModel;

        public TemplateManagementWindow(TemplateManagementViewModel viewModel)
        {
            this.viewModel = viewModel;
            DataContext = this.viewModel;

            InitializeComponent();
            //this.Resources.MergedDictionaries.Add(EtkWpfApplication.EtkWpfMainResources);
            //viewModel.PropertyChanged += OnPropertyChanged;
        }

        private void OnBindingDefinitionSelection(object sender, System.Windows.RoutedEventArgs e)
        {
            //this.MetroDialogOptions.ColorScheme = MetroDialogColorScheme.Accented;
            //CustomDialog dialog = this.Resources["BindingDefinitionsSelectionDlg"] as CustomDialog;

            //BindingDefinitionSelectionViewModel dlgViewmodel = new BindingDefinitionSelectionViewModel();
            //dialog.DataContext = dlgViewmodel;

            //dlgViewmodel.OnApply += () => { Task hide = this.HideMetroDialogAsync(dialog);
            //                                hide.Start();
            //                                hide.Wait();};
            //dlgViewmodel.OnCancel += () => { Task hide = this.HideMetroDialogAsync(dialog);
            //                                 hide.Start();
            //                                 hide.Wait();};
            //this.ShowMetroDialogAsync(dialog);
        }
    }
}
