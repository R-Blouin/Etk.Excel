using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;
using Etk.Excel.UI.MvvmBase;
using Etk.Excel.MvvmBase;

namespace Etk.Excel.UI.Windows.Wizard
{
    public class WizardControlViewModel : ViewModelBase
    {
        #region command
        private RelayCommand previousCommand;
        /// <summary> PreviousCommand command</summary>
        public ICommand PreviousCommand
        {
            get { return previousCommand ?? (previousCommand = new RelayCommand(param => CurrentStep -= 1)); }
        }

        private RelayCommand nextCommand;
        /// <summary> Next command</summary>
        public ICommand NextCommand
        {
            get
            {
                return nextCommand ?? (nextCommand = new RelayCommand(param =>
                                                                     {
                                                                         if (currentStep != stepMax && steps[currentStep] != null)
                                                                         {
                                                                             object parameters = steps[currentStep].GetNextStepData();
                                                                             int nextStep = currentStep + 1;
                                                                             if (parameters is IWizardStep)
                                                                             {
                                                                                 IWizardStep stepViewModel = (IWizardStep)  parameters;
                                                                                 if (steps[nextStep] != null)
                                                                                     steps[nextStep].CanNext -= CanNext;

                                                                                 steps[nextStep] = stepViewModel;
                                                                                 steps[nextStep].OnNext(parameters);
                                                                                 steps[nextStep].CanNext += CanNext;

                                                                                 if(ChangeStepViewModel != null)
                                                                                     ChangeStepViewModel(nextStep, stepViewModel);
                                                                             }
                                                                             if (steps[nextStep] != null)
                                                                             {
                                                                                 CanNext();
                                                                                 CurrentStep = nextStep;
                                                                             }
                                                                         }
                                                                     }));
            }
        }

        private RelayCommand finishCommand;
        /// <summary> Finish command</summary>
        public ICommand FinishCommand
        {
            get { return finishCommand ?? (finishCommand = new RelayCommand(param => CurrentStep = currentStep)); }
        }
        #endregion

        #region attributes and properties
        private int stepMax;
        public int StepMax
        {
            get { return stepMax; }
            set
            {
                stepMax = value;
                OnPropertyChanged("NextEnabled");
                OnPropertyChanged("PreviousEnabled");
                OnPropertyChanged("FinishEnabled");
            }
        }

        public Action<int, IWizardStep> ChangeStepViewModel;

        private int currentStep;
        public int CurrentStep
        {
            get { return currentStep; }
            set
            {
                currentStep = value;
                OnPropertyChanged("CurrentStep");
                OnPropertyChanged("NextEnabled");
                OnPropertyChanged("PreviousEnabled");
                OnPropertyChanged("FinishEnabled");
            }
        }

        /// <summary>Next is enabled</summary>
        public bool NextEnabled => currentStep >= 0 && currentStep < stepMax && steps[currentStep] != null && steps[currentStep].CheckCanNext();

        /// <summary>Previous is enabled</summary>
        public bool PreviousEnabled => currentStep > 0;

        /// <summary> Finish is enabled</summary>
        public bool FinishEnabled
        { get { return steps.FirstOrDefault(s => s != null && ! s.CheckCanNext()) == null; }}

        private readonly List<IWizardStep> steps;
        public IEnumerable<IWizardStep> Steps => steps;

        #endregion

        #region .ctors and factories
        public WizardControlViewModel()
        {
            currentStep = 0;
            steps = new List<IWizardStep>();
        }
        #endregion

        #region public methods
        public void AddStep(IWizardStep step)
        {
            if (step != null)
                step.CanNext += CanNext;
            steps.Add(step);
            StepMax = steps.Count;
        }

        public void CanNext()
        {
            if (currentStep != stepMax)
            {
                OnPropertyChanged("NextEnabled");
                OnPropertyChanged("FinishEnabled");
            }
        }
        #endregion
    }
}
