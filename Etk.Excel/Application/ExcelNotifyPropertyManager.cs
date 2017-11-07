using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Etk.Tools.Log;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Etk.Excel.BindingTemplates.Controls.WithFormula;

namespace Etk.Excel.Application
{
    class ExcelNotifyPropertyManager : IDisposable
    {
        private int sleepTime = 0;
        private volatile bool waitExcelBusy;
        private bool isDisposed;
        private readonly object syncObj = new object();
        private readonly BlockingCollection<ExcelNotityPropertyContext> contextItems;
        private readonly CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();

        private readonly ExcelApplication ExcelApplication;
        
        private readonly Thread thread;

        #region .ctors
        public ExcelNotifyPropertyManager(ExcelApplication excelApplication)
        {
            contextItems = new BlockingCollection<ExcelNotityPropertyContext>();
            ExcelApplication = excelApplication;

            thread = new Thread(Execute);
            thread.Name = "NotifyPropertiesChanged";
            thread.IsBackground = true;
            thread.Start();
        }
        #endregion
        
        #region public methods
        public void NotifyPropertyChanged(ExcelNotityPropertyContext context)
        {
            if (isDisposed)
                return;

            if (contextItems.FirstOrDefault(i => i.ContextItem == context.ContextItem && ! i.ChangeColor) != null)
                return;
            else
                contextItems.Add(context);
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
                    if (waitExcelBusy)
                    {
                        Thread.Sleep(sleepTime);
                        waitExcelBusy = false;
                    }

                    ExcelNotityPropertyContext context = contextItems.Take(cancellationTokenSource.Token);
                    if (context != null)
                        (ETKExcel.ExcelApplication as ExcelApplication).ExcelDispatcher.BeginInvoke(new Action(() => ExecuteNotity(context)));
                }
            }
            catch (Exception ex)
            {
                if (ex is OperationCanceledException)
                    Logger.Instance.Log(LogType.Info, "ExcelNotifyPropertyManager properly ended");
                else
                    Logger.Instance.LogException(LogType.Error, ex,
                        $"ExcelNotifyPropertyManager not properly ended:{ex.Message}");
            }
            finally
            {
                contextItems.Dispose();
            }
        }

        private void ExecuteNotity(ExcelNotityPropertyContext context)
        {
            if (isDisposed || context.ContextItem.IsDisposed || !context.View.IsRendered)
                return;

            ExcelInterop.Worksheet worksheet = null;
            ExcelInterop.Range range = null;
            bool enableEvent = ExcelApplication.Application.EnableEvents;
            try
            {
                worksheet = context.View.FirstOutputCell.Worksheet;
                KeyValuePair<int, int> kvp = context.Param;
                range = worksheet.Cells[context.View.FirstOutputCell.Row + kvp.Key, context.View.FirstOutputCell.Column + kvp.Value];
                if (range != null)
                {
                    object value = context.ContextItem.ResolveBinding();
                    if (value != null && value is Enum)
                        value = ((Enum)value).ToString();

                    if (! object.Equals(range.Value2, value))
                    {
                        if (enableEvent)
                            ExcelApplication.Application.EnableEvents = false;
                        range.Value2 = value;
                        if (context.ContextItem is ExcelContextItemWithFormula)
                        {
                            range.Calculate();
                            ((ExcelContextItemWithFormula)context.ContextItem).UpdateTarget(range.Value2);
                        }
                        context.View.CurrentSelectedCell?.Select();
                    }
                    context.ContextItem.BindingDefinition.DecoratorDefinition?.Resolve(range, context.ContextItem);
                }
                sleepTime = 0;
            }
            catch (COMException comEx)
            {
                waitExcelBusy = true;
                NotifyPropertyChanged(context);
                if (sleepTime < 1000)
                    sleepTime += 10;
            }
            catch (Exception ex)
            {
                string message = $"'ExecuteNotity' failed.{ex.Message}";
                Logger.Instance.LogException(LogType.Error, ex, message);
            }
            finally
            {
                if(worksheet != null)
                {
                    ExcelApplication.ReleaseComObject(worksheet);
                    worksheet = null;
                }
                try
                {

                    if (ExcelApplication.Application.EnableEvents != enableEvent)
                        ExcelApplication.Application.EnableEvents = enableEvent;
                }
                catch (COMException comEx)
                { }
            }
            range = null;
        }
        #endregion
    }
}
