using System.Windows.Controls;

namespace Etk.Excel.UI.Windows.Wizard
{
    public partial class WizardStepContainer : UserControl
    {
        WizardControlViewModel parent;

        public WizardStepContainerViewModel ViewModel
        { get; private set; }

        public WizardStepContainer(WizardControlViewModel parent, int step, UserControl nestedControl)
        {
            this.parent = parent;

            InitializeComponent();

            ViewModel = new WizardStepContainerViewModel(parent, step, nestedControl.DataContext as IWizardStep);
            StepContainer.Children.Add(nestedControl);
            DataContext = ViewModel;
        }

        public void ChangeStepViewModel(IWizardStep stepViewModel)
        {
            ViewModel.StepViewModel = stepViewModel;
            (StepContainer.Children[0] as UserControl).DataContext = ViewModel.StepViewModel;
        }
    }
}
