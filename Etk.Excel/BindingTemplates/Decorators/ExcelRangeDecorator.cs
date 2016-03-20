namespace Etk.Excel.BindingTemplates.Decorators
{
    using Etk.BindingTemplates.Context;
    using Etk.BindingTemplates.Definitions.Decorators;
    using Etk.Excel.BindingTemplates.Decorators.XmlDefinitions;
    using Etk.Excel.UI.Log;
    using Etk.Excel.UI.Reflection;
    using Microsoft.Office.Interop.Excel;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;

    class DecoratorProperty
    {
        public double FrontColor
        { get; private set; }

        public double BackColor
        { get; private set; }

        public DecoratorProperty(double frontColor, double backColor)
        {
            FrontColor = frontColor;
            BackColor = backColor;
        }
    }

    /// <summary>Excel concernedRange decorator. The decorator styles or colors come from a Excel concernedRange</summary>
    public class ExcelRangeDecorator : Decorator
    {
        //private bool useOnlyColors;
        private ILogger log = Logger.Instance;
        //private bool isRangedName;
        private string rangeId ;
        private Range decoratorRange;
        private Application excelApplication;
        private bool addConcernedRangeParameter;
        private bool notOnlyColor;
        private List<DecoratorProperty> decoratorProperties = null;

        #region .ctors and factories
        /// <summary> Constructor</summary>
        private ExcelRangeDecorator(Application excelApplication, string ident, string description, MethodInfo toInvoke, string rangeId, bool notOnlyColor)//, bool useOnlyColors)
                                  : base(ident, description, toInvoke)    
        {
            this.notOnlyColor = notOnlyColor;
            this.rangeId = rangeId;
            this.excelApplication = excelApplication;
            //this.useOnlyColors = useOnlyColors;
            // We try to initialyze the concernedRange used to decorate. But, maybe, not all the workbooks are loaded yet
            try
            {
                CheckParameters(toInvoke);
                RevolveDecoratorRange(); 
            }
            catch(Exception ex)
            {
                log.LogFormat(LogType.Warn, string.Format("'ExcelRangeDecorator' constructor:{0}", ex.Message));
            }
        }

        /// <summary> Factory</summary>
        public static ExcelRangeDecorator CreateInstance(Application excelApplication, XmlExcelRangeDecorator xmlDecorator)
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

                MethodInfo methodInfo = TypeHelpers.GetMethod(null, xmlDecorator.Method);
                ExcelRangeDecorator ret = new ExcelRangeDecorator(excelApplication, xmlDecorator.Ident, xmlDecorator.Description, methodInfo, xmlDecorator.Range, xmlDecorator.NotOnlyColor);//, xmlDecorator.UseOnlyColors);
                return ret;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Cannot create decorator '{0}':{1}", xmlDecorator.Ident ?? string.Empty, ex.Message), ex);
            }
        }
        #endregion

        #region public methods
        public bool Resolve(object sender, IBindingContextElement element)
        {
            Range concernedRange = sender as Range;
            if (concernedRange == null)
                return false;

            try
            {
                if (decoratorRange == null)
                    RevolveDecoratorRange();

                Range concernedRangeFirstCell = concernedRange.Cells[1, 1];

                // We delete the previous concernedRange comment 
                Comment comment = concernedRangeFirstCell.Comment;
                if (comment != null)
                    comment.Delete();

                // Prepare parameters
                object[] parameters;
                if (addConcernedRangeParameter)
                    parameters = new object[] { concernedRange, element.DataSource, null };
                else
                    parameters = new object[] { element.DataSource, null };

                // Invoke decorator resolver
                object result = ToInvoke.Invoke(ToInvoke.IsStatic ? null : element.DataSource, parameters);
                if (result != null)
                {
                    DecoratorResult decoratorResult = result as DecoratorResult;
                    if (decoratorResult.Item.HasValue)
                    {
                        if (!string.IsNullOrEmpty(decoratorResult.Comment))
                        {
                            concernedRangeFirstCell.AddComment(decoratorResult.Comment);
                            Comment addedComment = concernedRangeFirstCell.Comment;
                            addedComment.Visible = decoratorResult.CommentAlwaysVisible;
                            Shape shape = addedComment.Shape;
                            TextFrame textFrame = shape.TextFrame;
                            textFrame.AutoSize = true;
                        }

                        if (notOnlyColor)
                        {
                            decoratorRange[decoratorResult.Item.Value + 1].Copy();
                            concernedRange.PasteSpecial(XlPasteType.xlPasteFormats);
                        }
                        else
                        {
                            if (decoratorResult.Item.Value <= decoratorProperties.Count)
                            {
                                Interior interior = concernedRange.Interior;
                                Font font = concernedRange.Font;

                                font.Color = decoratorProperties[decoratorResult.Item.Value].FrontColor;
                                interior.Color = decoratorProperties[decoratorResult.Item.Value].BackColor;
                            }
                        }
                    }
                    return decoratorResult != null;
                }
                return false;
            }
            catch (Exception ex)
            {
                log.LogExceptionFormat(LogType.Error, ex, string.Format("Cannot resolve decorator '{0}':{1}", Ident, ex.Message));
                return false;
            }
        }

        /// <summary> Invoke the decorator</summary>
        /// <param name="sender">Range that ask for a decoration</param>
        /// <param name="contextItem">Binding bindingContextPart of the decoration request</param>
        /// <returns>True if decorator is resolved</returns>
        public override bool Resolve(object sender, IBindingContextItem contextItem)
        {
            Range concernedRange = sender as Range;
            if (concernedRange == null)
                return false;

            try
            {
                // We delete the previous concernedRange comment 
                Comment comment = concernedRange.Comment;
                if (comment != null)
                    comment.Delete();

                if (decoratorRange == null)
                    RevolveDecoratorRange();

                // Prepare parameters
                object[] parameters;
                if(addConcernedRangeParameter)
                    parameters = new object[] { concernedRange, contextItem.DataSource, contextItem.BindingDefinition.Name };
                else
                    parameters = new object[] { contextItem.DataSource, contextItem.BindingDefinition.Name };

                // Invoke decorator resolver
                object result = ToInvoke.Invoke(ToInvoke.IsStatic ? null : contextItem.DataSource, parameters);

                // addConcernedRangeParameter == true => the method resolver managed the style of the concerned targetRange 
                // if (addConcernedRangeParameter)
                //    return (bool) result;

                // addConcernedRangeParameter == false => the method resolver returns a 'DecoratorResult' we manage below 
                if (result != null)
                {
                    DecoratorResult decoratorResult = result as DecoratorResult;

                    if (!string.IsNullOrEmpty(decoratorResult.Comment))
                    {
                        concernedRange.AddComment(decoratorResult.Comment);
                        Comment addedComment = concernedRange.Comment;
                        addedComment.Visible = decoratorResult.CommentAlwaysVisible;
                        Shape shape = addedComment.Shape;
                        TextFrame textFrame = shape.TextFrame;
                        textFrame.AutoSize = true;
                    }
                    if (decoratorResult.Item.HasValue)
                    {
                        if (notOnlyColor)
                        {
                            decoratorRange[decoratorResult.Item.Value + 1].Copy();
                            concernedRange.PasteSpecial(XlPasteType.xlPasteFormats);
                        }
                        else
                        {
                            if (decoratorResult.Item.Value <= decoratorProperties.Count)
                            {
                                Interior interior = concernedRange.Interior;
                                Font font = concernedRange.Font;

                                font.Color = decoratorProperties[decoratorResult.Item.Value].FrontColor;
                                interior.Color = decoratorProperties[decoratorResult.Item.Value].BackColor;
                            }
                        }
                    }
                    return decoratorResult != null;
                }
                return false;
            }
            catch (Exception ex)
            {
                log.LogExceptionFormat(LogType.Error, ex, string.Format("Cannot resolve decorator '{0}':{1}", Ident, ex.Message));
                return false;
            }
        }
        #endregion

        #region private methods
        private bool CheckParameters(MethodInfo methodInfo)
        {
            addConcernedRangeParameter = false;
            bool error = false;
            ParameterInfo[] parametersInfo = methodInfo.GetParameters();
            if (parametersInfo == null || parametersInfo.Count() > 3 || parametersInfo.Count() < 2)
                error = true;

            if (!error && parametersInfo.Count() == 2)
            {
                if (methodInfo.ReturnType != typeof(DecoratorResult))
                    error = true;

                //if (parametersInfo[0].ParameterType != typeof(object))
                //    error = true;
                if (parametersInfo[1].ParameterType != typeof(string))
                    error = true;
            }
            if (!error && parametersInfo.Count() == 3)
            {
                addConcernedRangeParameter = true;

                if (methodInfo.ReturnType != typeof(DecoratorResult))
                    error = true;

                if (! parametersInfo[0].ParameterType.Name.Equals("Range"))
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
            return addConcernedRangeParameter;
        }

        private void RevolveDecoratorRange()
        {
            try
            {
                decoratorRange = excelApplication.get_Range(rangeId);
                if (decoratorRange != null)
                {
                    decoratorProperties = new List<DecoratorProperty>();
                    foreach (Range cell in decoratorRange.Cells)
                    {
                        Interior interior = cell.Interior;
                        Font font = cell.Font;
                        decoratorProperties.Add(new DecoratorProperty((double)font.Color, (double)interior.Color));
                    }
                }
            }
            catch (Exception ex)
            { 
                throw new Exception(string.Format("Cannot resolve Decorator range '{0}':{1}", rangeId ?? string.Empty, ex.Message));
            }
        }
        #endregion
    }
}
