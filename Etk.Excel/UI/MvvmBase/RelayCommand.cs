/// From: Josh Smith
/// http://msdn.microsoft.com/en-us/magazine/dd419663.aspx

namespace Etk.Excel.UI.MvvmBase
{
    using System;
    using System.Windows.Input;
    
    public class RelayCommand : ICommand
    {
        #region Fields
        private readonly Action<object> _execute;
        private readonly Predicate<object> _canExecute;
        #endregion

        #region Constructors
        public RelayCommand(Action<object> execute)
            : this(execute, null)
        {
        }

        public RelayCommand(Action<object> execute, Predicate<object> canExecute)
        {
            if (execute == null)
            {
                throw new ArgumentNullException("execute");
            }

            this._execute = execute;
            this._canExecute = canExecute;
        }
        #endregion

        #region ICommand Members
        public bool CanExecute(object parameter)
        {
            return this._canExecute == null ? true : this._canExecute(parameter);
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public void Execute(object parameter)
        {
            this._execute(parameter);
        }
        #endregion
    }

}
