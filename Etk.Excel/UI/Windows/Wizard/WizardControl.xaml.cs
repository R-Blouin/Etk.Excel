namespace Etk.Excel.UI.Windows.Wizard
{
    using System.Collections.Generic;
    using System.Windows.Controls;

    /// <summary>
    /// Logique d'interaction pour ExcelDynamicViewManagementWindow.xaml
    /// </summary>
    public partial class WizardControl : UserControl
    {
        List<WizardStepContainer> steps = new List<WizardStepContainer>();

        public WizardControlViewModel ViewModel
        { get; protected set; }

        public WizardControl()
        {
            InitializeComponent();

            ViewModel = new WizardControlViewModel();
            ViewModel.ChangeStepViewModel = (id, vm) => steps[id].ChangeStepViewModel(vm);

            DataContext = ViewModel;
        }

        public void AddStep(int stepId, UserControl control)
        {
            WizardStepContainer stepContainer = new WizardStepContainer(ViewModel, stepId, control);
            ViewModel.AddStep(stepContainer.ViewModel.StepViewModel);

            StepsContainer.Children.Add(stepContainer);
            steps.Add(stepContainer);
        }
    }
}