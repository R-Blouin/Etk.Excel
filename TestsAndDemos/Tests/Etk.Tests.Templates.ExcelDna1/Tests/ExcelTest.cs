namespace Etk.Tests.Templates.ExcelDna1.Tests
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Etk.Excel.BindingTemplates.Views;
    using Etk.Excel.UI.MvvmBase;
    using Etk.Tests.Templates.ExcelDna1.Extensions;

    abstract class ExcelTest : ViewModelBase, IExcelTest
    {
        #region propertis and attributes
        public string Description
        { get; private set; }

        private bool done;
        public bool Done
        {
            get { return done; }
            private set
            {
                done = value;
                OnPropertyChanged("Done");
            }
        }

        private bool success;
        public bool Success
        {
            get { return success; }
            protected set
            {
                success = value;
                OnPropertyChanged("Success");
            }
        }

        private string errors;
        public string Errors
        {
            get { return errors; }
            protected set
            {
                errors = value;
                OnPropertyChanged("Errors");
            }
        }

        protected List<string> StepsErrorMessages
        { get; private set; }
        #endregion

        #region .ctors
        protected ExcelTest(string description)
        {
            Description = description;
            StepsErrorMessages = new List<string>();
        }
        #endregion

        #region public methods
        public void InitTestStatus()
        {
            Success = Done = false;
            Errors = null;
        }

        public void Execute(IExcelTemplateView view)
        {
            try
            {
                RealExecute(view);
            }
            catch (Exception ex)
            {
                StepsErrorMessages.Add(ex.ToString(null));
            }
            finally
            {
                Done = true;
                if (! StepsErrorMessages.Any())
                {
                    Success = true;
                    Errors = null;
                }
                else
                {
                    Success = false;
                    Errors = string.Join("\r\n", StepsErrorMessages.ToArray());
                }
            }
        }
        #endregion

        #region protected and private methods
        abstract protected void RealExecute(IExcelTemplateView view);
        #endregion
    }
}
