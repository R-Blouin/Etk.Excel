using System;
using System.Runtime.InteropServices;
using System.Threading;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Decorators;
using Etk.BindingTemplates.Definitions.EventCallBacks;
using Etk.Tools.Log;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.BindingTemplates.Decorators
{
    public class ExcelRangeSimpleDecorator : Decorator
    {
        private readonly ILogger log = Logger.Instance;

        private static EventCallbacksManager eventCallbacksManager;
        private static EventCallbacksManager EventCallbacksManager => eventCallbacksManager ??
                                                                      (eventCallbacksManager = CompositionManager.Instance.GetExportedValue<EventCallbacksManager>());

        #region .ctors and factories
        /// <summary> Constructor</summary>
        public ExcelRangeSimpleDecorator(string ident, EventCallback callback) : base(ident, null, callback)    
        {
            try
            {
                CheckParameters(callback);
            }
            catch(Exception ex)
            {
                log.LogFormat(LogType.Warn, $"'ExcelRangeDecorator2' constructor failed:{ex.Message}");
            }
        }
        #endregion

        #region public methods
        public bool Resolve(object sender, IBindingContextElement element)
        {
            ExcelInterop.Range concernedRange = sender as ExcelInterop.Range;
            if (concernedRange == null)
                return false;

            try
            {
                ExcelInterop.Range concernedRangeFirstCell = concernedRange.Cells[1, 1];

                // We delete the previous concernedRange comment 
                ExcelInterop.Comment comment = concernedRangeFirstCell.Comment;
                comment?.Delete();

                // Invoke decorator resolver
                object result = EventCallbacksManager.DecoratorInvoke(Callback, concernedRange, element.DataSource, null);
                if (result != null)
                {
                    string commentStr = result as string;
                    if (!string.IsNullOrEmpty(commentStr))
                    {
                        concernedRange.AddComment(commentStr);
                        ExcelInterop.Comment addedComment = concernedRange.Comment;
                        ExcelInterop.Shape shape = addedComment.Shape;
                        ExcelInterop.TextFrame textFrame = shape.TextFrame;
                        textFrame.AutoSize = true;
                    }
                    return commentStr != null;
                }
                return false;
            }
            catch (COMException comEx)
            {
                if (comEx.ErrorCode == ETKExcel.EXCEL_BUSY)
                {
                    Thread.Sleep(ETKExcel.WAITINGTIME_EXCEL_BUSY);
                    return Resolve(sender, element);
                }
                log.LogExceptionFormat(LogType.Error, comEx, $"Cannot resolve decorator '{Ident}':{comEx.Message}");
                return false;
            }
            catch (Exception ex)
            {
                log.LogExceptionFormat(LogType.Error, ex, $"Cannot resolve decorator2 '{Ident}':{ex.Message}");
                return false;
            }
        }

        /// <summary> Invoke the decorator</summary>
        /// <param name="sender">Range that ask for a decoration</param>
        /// <param name="contextItem">Binding bindingContextPart of the decoration request</param>
        /// <returns>True if decorator is resolved</returns>
        public override bool Resolve(object sender, IBindingContextItem contextItem)
        {
            ExcelInterop.Range concernedRange = sender as ExcelInterop.Range;
            if (concernedRange == null)
                return false;

            try
            {
                // We delete the previous concernedRange comment 
                ExcelInterop.Comment comment = concernedRange.Comment;
                comment?.Delete();

                // Invoke decorator resolver
                object result = EventCallbacksManager.DecoratorInvoke(Callback, concernedRange, contextItem.DataSource, contextItem.BindingDefinition.Name);
                if (result != null)
                {
                    string commentStr = result as string;
                    if (!string.IsNullOrEmpty(commentStr))
                    {
                        concernedRange.AddComment(commentStr);
                        ExcelInterop.Comment addedComment = concernedRange.Comment;
                        ExcelInterop.Shape shape = addedComment.Shape;
                        ExcelInterop.TextFrame textFrame = shape.TextFrame;
                        textFrame.AutoSize = true;
                    }
                    return commentStr != null;
                }
                return false;
            }
            catch (COMException comEx)
            {
                if (comEx.ErrorCode == ETKExcel.EXCEL_BUSY)
                {
                    Thread.Sleep(ETKExcel.WAITINGTIME_EXCEL_BUSY);
                    return Resolve(sender, contextItem);
                }
                log.LogExceptionFormat(LogType.Error, comEx, $"Cannot resolve decorator '{Ident}':{comEx.Message}");
                return false;
            }
            catch (Exception ex)
            {
                log.LogExceptionFormat(LogType.Error, ex, $"Cannot resolve simple decorator '{Ident}':{ex.Message}");
                return false;
            }
        }
        #endregion

        #region private methods
        private void CheckParameters(EventCallback callback)
        {
            //addConcernedRangeParameter = false;
            //bool error = false;

            //if (callback.IsNotDotNet)
            //    addConcernedRangeParameter = true;
            //else
            //{
            //    ParameterInfo[] parametersInfo = callback.Callback.GetParameters();
            //    if (parametersInfo == null || parametersInfo.Count() > 3 || parametersInfo.Count() < 2)
            //        error = true;

            //    if (!error && parametersInfo.Count() == 2)
            //    {
            //        if (callback.Callback.ReturnType != typeof(DecoratorResult))
            //            error = true;

            //        //if (parametersInfo[0].ParameterType != typeof(object))
            //        //    error = true;
            //        if (parametersInfo[1].ParameterType != typeof(string))
            //            error = true;
            //    }
            //    if (!error && parametersInfo.Count() == 3)
            //    {
            //        addConcernedRangeParameter = true;

            //        if (callback.Callback.ReturnType != typeof(DecoratorResult))
            //            error = true;

            //        if (!parametersInfo[0].ParameterType.Name.Equals("Range"))
            //            error = true;
            //        //if (parametersInfo[1].ParameterType != typeof(object))
            //        //    error = true;
            //        if (parametersInfo[2].ParameterType != typeof(string))
            //            error = true;
            //    }

            //    if (error)
            //    {
            //        throw new Exception("MethodInfo must be 'DecoratorResult MethodName(Range <range to decorate>, object <object bound with the range to decorate>, string <expression bound with the range to decorate>)'"
            //                             + "\r\n'DecoratorResult MethodName(object <object bound with the range to decorate>, string <expression bound with the range to decorate>)'");
            //    }
            //}
        }
        #endregion
    }
}
