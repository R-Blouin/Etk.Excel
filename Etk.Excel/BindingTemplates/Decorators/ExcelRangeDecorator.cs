using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Decorators;
using Etk.BindingTemplates.Definitions.EventCallBacks;
using Etk.Excel.BindingTemplates.Decorators.XmlDefinitions;
using Etk.Tools.Log;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;
using Etk.Excel.Application;

namespace Etk.Excel.BindingTemplates.Decorators
{
    class DecoratorProperty
    {
        public double FrontColor
        { get; }

        public double BackColor
        { get;  }

        public DecoratorProperty(double frontColor, double backColor)
        {
            FrontColor = frontColor;
            BackColor = backColor;
        }
    }

    /// <summary>Excel concernedRange decorator. The decorator styles or colors come from a Excel concernedRange</summary>
    public class ExcelRangeDecorator : Decorator
    {
        private static EventCallbacksManager eventCallbacksManager;
        private static EventCallbacksManager EventCallbacksManager => eventCallbacksManager ??
                                                                      (eventCallbacksManager = CompositionManager.Instance.GetExportedValue<EventCallbacksManager>());

        //private bool useOnlyColors;
        private readonly ILogger log = Logger.Instance;
        //private bool isRangedName;
        private readonly string rangeId ;
        private ExcelInterop.Range decoratorRange;
        //private readonly ExcelInterop.Application excelApplication;
        private bool addConcernedRangeParameter;
        private readonly bool notOnlyColor;
        private List<DecoratorProperty> decoratorProperties;

        #region .ctors and factories
        /// <summary> Constructor</summary>
        private ExcelRangeDecorator(string ident, string description, EventCallback callback, string rangeId, bool notOnlyColor)//, bool useOnlyColors)
                                  : base(ident, description, callback)    
        {
            this.notOnlyColor = notOnlyColor;
            this.rangeId = rangeId;
            //this.useOnlyColors = useOnlyColors;
            // We try to initialyze the concernedRange used to decorate. But, maybe, not all the workbooks are loaded yet
            try
            {
                CheckParameters(callback);
                RevolveDecoratorRange(); 
            }
            catch(Exception ex)
            {
                log.LogFormat(LogType.Warn, $"'ExcelRangeDecorator' constructor:{ex.Message}");
            }
        }

        /// <summary> Factory</summary>
        public static ExcelRangeDecorator CreateInstance(XmlExcelRangeDecorator xmlDecorator)
        {
            if (xmlDecorator == null)
                return null;

            try
            {
                if (string.IsNullOrEmpty(xmlDecorator.Ident))
                    throw new Exception("A decorator cannot be null or empty");

                if (string.IsNullOrEmpty(xmlDecorator.Method))
                    throw new Exception("The method info parameter cannot be null or empty");

                if (string.IsNullOrEmpty(xmlDecorator.Range))
                    throw new Exception("The range ident of a decorator cannot be null or empty");
                EventCallback callback = EventCallbacksManager.RetrieveCallback(null, xmlDecorator.Method);
                ExcelRangeDecorator ret = new ExcelRangeDecorator(xmlDecorator.Ident, xmlDecorator.Description, callback, xmlDecorator.Range, xmlDecorator.NotOnlyColor);//, xmlDecorator.UseOnlyColors);
                return ret;
            }
            catch (Exception ex)
            {
                throw new Exception($"Cannot create decorator '{xmlDecorator.Ident ?? string.Empty}':{ex.Message}");
            }
        }
        #endregion

        #region public methods
        public bool Resolve(object sender, IBindingContextElement element)
        {
            ExcelInterop.Range concernedRange = sender as ExcelInterop.Range;
            if (concernedRange == null)
                return false;

            ExcelInterop.Range concernedRangeFirstCell = null;
            try
            {
                if (decoratorRange == null)
                    RevolveDecoratorRange();

                concernedRangeFirstCell = concernedRange.Cells[1, 1];

                // We delete the previous concernedRange comment 
                ExcelInterop.Comment comment = concernedRangeFirstCell.Comment;
                if(comment != null)
                {
                    comment.Delete();
                    ExcelApplication.ReleaseComObject(comment);
                }

                // Invoke decorator resolver
                object result = EventCallbacksManager.DecoratorInvoke(Callback, addConcernedRangeParameter ? concernedRange : null, element.DataSource, null);
                if (result != null)
                {
                    DecoratorResult decoratorResult = result as DecoratorResult;
                    if (decoratorResult?.Item != null)
                    {
                        if (!string.IsNullOrEmpty(decoratorResult.Comment))
                        {
                            concernedRangeFirstCell.AddComment(decoratorResult.Comment);
                            ExcelInterop.Comment addedComment = concernedRangeFirstCell.Comment;
                            addedComment.Visible = decoratorResult.CommentAlwaysVisible;
                            ExcelInterop.Shape shape = addedComment.Shape;
                            ExcelInterop.TextFrame textFrame = shape.TextFrame;
                            textFrame.AutoSize = true;

                            ExcelApplication.ReleaseComObject(textFrame);
                            ExcelApplication.ReleaseComObject(shape);
                            ExcelApplication.ReleaseComObject(addedComment);
                        }

                        if (notOnlyColor)
                        {
                            decoratorRange[decoratorResult.Item.Value + 1].Copy();
                            concernedRange.PasteSpecial(ExcelInterop.XlPasteType.xlPasteFormats);
                        }
                        else
                        {
                            if (decoratorResult.Item.Value <= decoratorProperties.Count)
                            {
                                ExcelInterop.Interior interior = concernedRange.Interior;
                                ExcelInterop.Font font = concernedRange.Font;

                                font.Color = decoratorProperties[decoratorResult.Item.Value].FrontColor;
                                interior.Color = decoratorProperties[decoratorResult.Item.Value].BackColor;

                                ExcelApplication.ReleaseComObject(font);
                                ExcelApplication.ReleaseComObject(interior);
                            }
                        }
                    }
                    return decoratorResult != null;
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
                log.LogExceptionFormat(LogType.Error, ex, $"Cannot resolve decorator '{Ident}':{ex.Message}");
                return false;
            }
            finally
            {
                ExcelApplication.ReleaseComObject(concernedRangeFirstCell);
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
                if(comment != null)
                {
                    comment.Delete();
                    ExcelApplication.ReleaseComObject(comment);
                }

                if (decoratorRange == null)
                    RevolveDecoratorRange();

                // Invoke decorator resolver
                object result = EventCallbacksManager.DecoratorInvoke(Callback, addConcernedRangeParameter ? concernedRange : null, contextItem.DataSource, contextItem.BindingDefinition.Name);
                // addConcernedRangeParameter == false => the method resolver returns a 'DecoratorResult' we manage below 
                if (result != null)
                {
                    DecoratorResult decoratorResult = result as DecoratorResult;

                    if (!string.IsNullOrEmpty(decoratorResult.Comment))
                    {
                        concernedRange.AddComment(decoratorResult.Comment);
                        ExcelInterop.Comment addedComment = concernedRange.Comment;
                        addedComment.Visible = decoratorResult.CommentAlwaysVisible;
                        ExcelInterop.Shape shape = addedComment.Shape;
                        ExcelInterop.TextFrame textFrame = shape.TextFrame;
                        textFrame.AutoSize = true;

                        ExcelApplication.ReleaseComObject(textFrame);
                        ExcelApplication.ReleaseComObject(shape);
                        ExcelApplication.ReleaseComObject(addedComment);
                    }
                    if (decoratorResult.Item.HasValue)
                    {
                        if (notOnlyColor)
                        {
                            decoratorRange[decoratorResult.Item.Value + 1].Copy();
                            concernedRange.PasteSpecial(ExcelInterop.XlPasteType.xlPasteFormats);
                        }
                        else
                        {
                            if (decoratorResult.Item.Value <= decoratorProperties.Count)
                            {
                                ExcelInterop.Interior interior = concernedRange.Interior;
                                ExcelInterop.Font font = concernedRange.Font;

                                font.Color = decoratorProperties[decoratorResult.Item.Value].FrontColor;
                                interior.Color = decoratorProperties[decoratorResult.Item.Value].BackColor;

                                ExcelApplication.ReleaseComObject(font);
                                ExcelApplication.ReleaseComObject(interior);
                            }
                        }
                    }
                    return decoratorResult != null;
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
                log.LogExceptionFormat(LogType.Error, ex, $"Cannot resolve decorator '{Ident}':{ex.Message}");
                return false;
            }
        }
        #endregion

        #region private methods
        private void CheckParameters(EventCallback callback)
        {
            addConcernedRangeParameter = false;
            bool error = false;

            if (callback.IsNotDotNet)
                addConcernedRangeParameter = true;
            else
            {
                ParameterInfo[] parametersInfo = callback.Callback.GetParameters();
                if (parametersInfo == null || parametersInfo.Count() > 3 || parametersInfo.Count() < 2)
                    error = true;

                if (!error && parametersInfo.Count() == 2)
                {
                    if (callback.Callback.ReturnType != typeof(DecoratorResult))
                        error = true;

                    //if (parametersInfo[0].ParameterType != typeof(object))
                    //    error = true;
                    if (parametersInfo[1].ParameterType != typeof(string))
                        error = true;
                }
                if (!error && parametersInfo.Count() == 3)
                {
                    addConcernedRangeParameter = true;

                    if (callback.Callback.ReturnType != typeof(DecoratorResult))
                        error = true;

                    if (!parametersInfo[0].ParameterType.Name.Equals("Range"))
                        error = true;
                    //if (parametersInfo[1].ParameterType != typeof(object))
                    //    error = true;
                    if (parametersInfo[2].ParameterType != typeof(string))
                        error = true;
                }

                if (error)
                {
                    throw new Exception("MethodInfo must be 'DecoratorResult MethodName(Range <range to decorate>, object <object bound with the range to decorate>, string <expression bound with the range to decorate>)'"
                                         + "\r\n'DecoratorResult MethodName(object <object bound with the range to decorate>, string <expression bound with the range to decorate>)'");
                }
            }
        }

        private void RevolveDecoratorRange()
        {
            try
            {
                decoratorRange = ETKExcel.ExcelApplication.Application.get_Range(rangeId);
                if (decoratorRange != null)
                {
                    decoratorProperties = new List<DecoratorProperty>();
                    foreach (ExcelInterop.Range cell in decoratorRange.Cells)
                    {
                        ExcelInterop.Interior interior = cell.Interior;
                        ExcelInterop.Font font = cell.Font;

                        decoratorProperties.Add(new DecoratorProperty((double)font.Color, (double)interior.Color));

                        ExcelApplication.ReleaseComObject(font);
                        ExcelApplication.ReleaseComObject(interior);
                        ExcelApplication.ReleaseComObject(cell);
                    }
                }
            }
            catch (Exception ex)
            { 
                throw new Exception($"Cannot resolve Decorator range '{rangeId ?? string.Empty}':{ex.Message}");
            }
        }
        #endregion
    }
}
