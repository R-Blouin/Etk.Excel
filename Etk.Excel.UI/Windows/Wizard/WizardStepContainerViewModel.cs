using Etk.Excel.MvvmBase;

namespace Etk.Excel.UI.Windows.Wizard
{
    public class WizardStepContainerViewModel : ViewModelBase
    {
        public int Step
        { get; private set; }

        public WizardControlViewModel ParentWizard
        { get; private set; }

        public bool StepVisible => ParentWizard.CurrentStep == Step;

        IWizardStep stepViewModel;
        public IWizardStep StepViewModel
        {
            get { return stepViewModel; }
            set
            {
                stepViewModel = value;
                OnPropertyChanged("StepViewModel");
            }
        }

        public WizardStepContainerViewModel(WizardControlViewModel parent, int step, IWizardStep stepViewModel)
        {
            ParentWizard = parent;
            Step = step;
            this.stepViewModel = stepViewModel;

            parent.PropertyChanged += (o, e) => { if (e.PropertyName.Equals("CurrentStep"))
                                                    OnPropertyChanged("StepVisible"); };
        }
    }
}
