using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Threading;
using Etk.Tools.Log;
using System.Collections.Generic;

namespace Etk.Excel.Application
{
    class ExcelPostListAsynchronousManager
    {
        private volatile bool waitExcelBusy = false;
        private volatile bool tryAgain = true;

        private readonly Dispatcher dispatcher;
        private readonly IEnumerable<Action> actions;
        private Action postExecutionAction;

        #region .ctors
        public ExcelPostListAsynchronousManager(Dispatcher dispatcher, IEnumerable<Action> actions, Action postExecutionAction = null)
        {
            this.actions = actions;
            this.postExecutionAction = postExecutionAction;
            this.dispatcher = dispatcher;
        }
        #endregion

        #region public methods
        public void Execute()
        {
            Thread thread = new Thread(ExecuteActions);
            thread.Name = "PostListAsynchronousActions";
            thread.IsBackground = true;
            thread.Start();
        }
        #endregion

        #region private methods
        private void ExecuteActions()
        {
            try
            {
                foreach (Action action in actions)
                {
                    if (action != null)
                    {
                        tryAgain = true;
                        while (tryAgain)
                        {
                            DispatcherOperation operation = dispatcher.BeginInvoke(new Action(() =>
                            {
                                try
                                {
                                    action();
                                }
                                catch (COMException)
                                {
                                    waitExcelBusy = true;
                                }
                                catch (Exception ex)
                                {
                                    string message = $"'ExcelPostAsynchronousManager.ExecuteAction' failed.{ex.Message}";
                                    Logger.Instance.LogException(LogType.Error, ex, message);
                                }
                            }));
                            operation.Wait();
                            if (waitExcelBusy)
                            {
                                Thread.Sleep(50);
                                waitExcelBusy = false;
                                tryAgain = true;
                            }
                            else
                                tryAgain = false;
                        }
                    }
                }

                ExecutePostExecutionAction();
            }
            finally
            {
            }
        }

        private void ExecutePostExecutionAction()
        {
            //int cpt = 0;
            while (postExecutionAction != null)
            {
                DispatcherOperation operation = dispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        postExecutionAction();
                        postExecutionAction = null;
                    }
                    catch (COMException)
                    {
                        waitExcelBusy = true;
                    }
                    catch (Exception ex)
                    {
                        postExecutionAction = null;
                        string message = $"'ExcelPostAsynchronousManager.ExecutePostExecutionAction' failed.{ex.Message}";
                        Logger.Instance.LogException(LogType.Error, ex, message);
                    }
                }));
                operation.Wait();
                if (waitExcelBusy)
                {
                    //Interlocked.Increment(ref cpt);
                    //if (cpt >= 10)
                    //{
                    //    postExecutionAction = null;
                    //    Logger.Instance.Log(LogType.Error, "'ExcelPostAsynchronousManager.ExecutePostExecutionAction' failed 10 times. Execution stopped !");
                    //}
                    //else
                    {
                        Thread.Sleep(50);
                        waitExcelBusy = false;
                    }
                }
            }
        }
        #endregion
    }
}
