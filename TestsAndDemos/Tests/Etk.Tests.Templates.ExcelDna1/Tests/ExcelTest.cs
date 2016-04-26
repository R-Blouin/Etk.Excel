namespace Etk.Tests.Templates.ExcelDna1.Tests
{
    using System;
    using Etk.Excel.BindingTemplates.Views;
    using Etk.Excel.UI.MvvmBase;
    using Etk.Tests.Templates.ExcelDna1.Extensions;

    abstract class ExcelTest : ViewModelBase, IExcelTest
    {
        #region p^ropertis and attributes
        protected IExcelTemplateView View
        { get; private set; }

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

        private string exception;
        public string Exception
        {
            get { return exception; }
            protected set
            {
                exception = value;
                OnPropertyChanged("Exception");
            }
        }
        #endregion

        #region .ctors
        protected ExcelTest(IExcelTemplateView view, string description)
        {
            View = view;
            Description = description;

            Init();
        }
        #endregion

        #region public methods
        public void Execute()
        {
            try
            {
                Init();
                RealExecute();
            }
            catch (Exception ex)
            {
                Success = false;
                Exception = ex.ToString(null);
            }
            finally
            {
                Done = true;
            }
        }
        #endregion

        #region protected and private methods
        abstract protected void RealExecute();

        private void Init()
        {
            Success = Done = false;
            Exception = null;
        }
        #endregion
    }
}
