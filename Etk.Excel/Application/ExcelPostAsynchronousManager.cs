using System;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Threading;
using Etk.Tools.Log;
using System.Collections.Generic;

namespace Etk.Excel.Application
{
    class ExcelPostAsynchronousManager : IDisposable
    {
        private volatile bool waitExcelBusy = false;
        private volatile bool isDisposed;
        private readonly object syncObj = new object();
        private readonly BlockingCollection<Action> actions;
        private readonly Dispatcher dispatcher;
        private readonly CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();

        private readonly Thread thread;

        #region .ctors
        public ExcelPostAsynchronousManager(Dispatcher dispatcher)
        {
            actions = new BlockingCollection<Action>(new ConcurrentStack<Action>());
            this.dispatcher = dispatcher;
            thread = new Thread(Execute);
            thread.Name = "PostAsynchronousActions";
            thread.IsBackground = true;
            thread.Start();
        }
        #endregion
        
        #region public methods
        public void PostAction(Action action)
        {
            if (isDisposed || dispatcher == null)
                return;

            if(action != null)
                actions.Add(action);
        }

        public void PostActions(IEnumerable<Action> listOfActions)
        {
            if (isDisposed || dispatcher == null || listOfActions == null)
                return;

            foreach(Action action in listOfActions)
                actions.Add(action);
        }

        public void Dispose()
        {
            try
            {
                lock (syncObj)
                {
                    if (!isDisposed)
                    {
                        isDisposed = true;
                        cancellationTokenSource.Cancel();
                    }
                }
            }
            catch
            {}
        }
        #endregion

        #region private methods
        private void Execute()
        {
            try
            {
                while (!isDisposed)
                {
                    Action action = actions.Take(cancellationTokenSource.Token);
                    if (action != null)
                    {
                        DispatcherOperation operation = dispatcher.BeginInvoke(new Action(() =>
                                                        {
                                                            try
                                                            {
                                                                if (!isDisposed)
                                                                    action();
                                                            }
                                                            catch (COMException)
                                                            {
                                                                actions.Add(action);
                                                                waitExcelBusy = true;
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                string message = string.Format("'ExcelPostAsynchronousManager.ExecuteAction' failed.{0}", ex.Message);
                                                                Logger.Instance.LogException(LogType.Error, ex, message);
                                                            }
                                                        }));
                        operation.Wait();
                        if (waitExcelBusy)
                        {
                            Thread.Sleep(50);
                            waitExcelBusy = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (ex is OperationCanceledException)
                    Logger.Instance.Log(LogType.Info, "PostAsynchronousManager properly ended");
                else
                    Logger.Instance.LogException(LogType.Error, ex, string.Format("PostAsynchronousManager not properly ended", ex.Message));
            }
            finally
            {
                actions.Dispose();
            }
        }
        #endregion
    }
}
