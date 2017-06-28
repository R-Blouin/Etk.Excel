using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Etk.BindingTemplates.Definitions.Decorators;
using Etk.BindingTemplates.Definitions.EventCallBacks;
using Etk.Tools.Reflection;
using Etk.BindingTemplates.Definitions.Templates;

namespace Etk.BindingTemplates.Definitions.Binding
{
    public enum ShowHideMode
    {
        None,
        StartShown,
        StartHidden    
    }

    public class BindingDefinitionDescription
    {
        private static DecoratorsManager decoratorsManager;
        private static DecoratorsManager DecoratorsManager
        {
            get
            {
                if (decoratorsManager == null)
                    decoratorsManager = CompositionManager.Instance.GetExportedValue<DecoratorsManager>();
                return decoratorsManager;
            }
        }

        private static EventCallbacksManager eventCallbacksManager;
        private static EventCallbacksManager EventCallbacksManager
        {
            get
            {
                if (eventCallbacksManager == null)
                    eventCallbacksManager = CompositionManager.Instance.GetExportedValue<EventCallbacksManager>();
                return eventCallbacksManager;
            }
        }

        public string Name
        { get; set; }

        public bool IsConst
        { get; set; }

        public bool IsReadOnly
        { get; set; }

        public string BindingExpression
        { get;  set; }

        public string Description
        { get; set; }

        public Decorator Decorator
        { get; private set; }

        public MethodInfo OnSelection
        { get; private set; }

        public MethodInfo OnLeftDoubleClick
        { get; private set; }

        public bool IsMultiLine
        { get; private set; }

        public double MultiLineFactor
        { get; private set; }

        public MethodInfo MultiLineFactorResolver
        { get; private set; }

        public ShowHideMode ShowHideMode
        { get; private set; }

        public int ShowHideValue
        { get; private set; }

        public string Formula
        { get; private set; }

        #region .ctors and factories
        public BindingDefinitionDescription()
        {}

        private BindingDefinitionDescription(ITemplateDefinition templateDefinition, string bindingExpression, bool isConst, List<string> options)
        {
            BindingExpression = bindingExpression;
            IsConst = isConst;
            if (options != null)
            {
                IsReadOnly = options.Contains("R");
                options.Remove("R");

                if (options.Count > 0)
                {
                    foreach (string option in options)
                    {
                        // Formula
                        if (option.StartsWith("F="))
                        {
                            Formula = option.Substring(2);
                            continue;
                        }

                        // Description
                        if (option.StartsWith("D="))
                        {
                            Description = option.Substring(2);
                            continue;
                        }
                        // Name
                        if (option.StartsWith("N="))
                        {
                            Name = option.Substring(2);
                            continue;
                        }
                        // Decorator
                        if (option.StartsWith("DEC="))
                        {
                            string decoratorIdent = option.Substring(4);
                            Decorator = DecoratorsManager.GetDecorator(decoratorIdent);
                            continue;
                        }
                        // On Selection 
                        if (option.StartsWith("S="))
                        {
                            string methodInfoIdent = option.Substring(2);
                            OnSelection = RetrieveMethodInfo(templateDefinition, option, methodInfoIdent);
                            continue;
                        }
                        // On double left click
                        if (option.StartsWith("LDC="))
                        {
                            string methodInfoIdent = option.Substring(4);
                            OnLeftDoubleClick = RetrieveMethodInfo(templateDefinition, option, methodInfoIdent);
                            continue;
                        }
                        // MultiLine based on the number of line in the bound property
                        if (option.StartsWith("M="))
                        {
                            IsMultiLine = true;
                            string factor = option.Substring(2);
                            double multiLineFactor;
                            if (! string.IsNullOrEmpty(factor) && double.TryParse(factor, out multiLineFactor))
                                MultiLineFactor = multiLineFactor;
                            else
                                MultiLineFactor = 1.5;
                            continue;
                        }
                        // MultiLine where the number of line is determinated by a callback invocation
                        if (option.StartsWith("ME="))
                        {
                            string multiLineFactorResolver = option.Substring(3);
                            if (!string.IsNullOrEmpty(multiLineFactorResolver))
                            {
                                try
                                {
                                    MultiLineFactorResolver = RetrieveMethodInfo(templateDefinition, null, multiLineFactorResolver);
                                    if (MultiLineFactorResolver != null)
                                    {
                                        int parametersCpt = MultiLineFactorResolver.GetParameters().Length;
                                        if (MultiLineFactorResolver.ReturnType != typeof(int) || 
                                            parametersCpt > 1 ||
                                            (parametersCpt == 1 && !(MultiLineFactorResolver.GetParameters()[0].ParameterType.IsAssignableFrom(typeof(object)))))
                                            throw new Exception("The function prototype must be defined as 'int <Function Name>([param]) with 'param' inheriting from 'system.object'");
                                        IsMultiLine = true;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    throw new Exception(string.Format("Cannot resolve the 'ME' attribute for the binding definition '{0}'", bindingExpression), ex);
                                }
                            }
                        }
                        // On double left click, Show/Hide the x following/preceding colums/rows. Start hidden  
                        if (option.StartsWith("SV="))
                        {
                            string numberOfConcernedColumns = option.Substring(3);
                            int wrk; 
                            if (string.IsNullOrEmpty(numberOfConcernedColumns) && int.TryParse(numberOfConcernedColumns, out wrk))
                            {
                                ShowHideMode = ShowHideMode.StartShown;
                                ShowHideValue = wrk;
                                continue;
                            }
                            throw new Exception("The 'Show/Hide' prototype must be defined as 'SV=<int>' whre '<int>' is a integer");
                        }
                        // On double left click, Show/Hide the x following/preceding colums. Start shown
                        if (option.StartsWith("SH="))
                        {
                            string numberOfConcernedColumns = option.Substring(3);
                            int wrk;
                            if (string.IsNullOrEmpty(numberOfConcernedColumns) && int.TryParse(numberOfConcernedColumns, out wrk))
                            {
                                ShowHideMode = ShowHideMode.StartHidden;
                                ShowHideValue = wrk;
                                continue;
                            }
                            throw new Exception("The 'Show/Hide' columns prototype must be defined as 'SH=<int>' whre '<int>' is a integer");
                        }
                    }
                }
            }
        }

        private MethodInfo RetrieveMethodInfo(ITemplateDefinition templateDefinition, string option, string methodInfoIdent)
        {
            MethodInfo methodInfo = null;
            try
            {
                string[] parts = methodInfoIdent.Split(',');
                if (parts.Count() == 1)
                {
                    EventCallback callBack = EventCallbacksManager.GetCallback(methodInfoIdent);
                    if (callBack != null)
                        methodInfo = callBack.Callback;
                }
                if (parts.Count() == 3)
                {
                    Type inType = null;
                    string methodInfoName = methodInfoIdent;
                    if (string.IsNullOrEmpty(parts[0]) && string.IsNullOrEmpty(parts[1]))
                    {
                        if(templateDefinition != null && templateDefinition.MainBindingDefinition != null && templateDefinition.MainBindingDefinition.BindingType != null)
                        {
                            inType = templateDefinition.MainBindingDefinition.BindingType;
                            methodInfoName = parts[2];
                        }
                    }
                    methodInfo = TypeHelpers.GetMethod(inType, methodInfoName);
                }

                if (methodInfo == null)
                    throw new Exception(string.Format("Cannot find the callback '{0}'", methodInfoIdent));

                return methodInfo;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Property option '{0}'. {1}", option, ex.Message));
            }
        }

        public static BindingDefinitionDescription CreateBindingDescription(ITemplateDefinition templateDefinition, string toAnalyze, string trimmedToAnalyze)
        {
            BindingDefinitionDescription ret = null;
            string bindingExpression;
            List<string> options = null;
            if (!string.IsNullOrEmpty(trimmedToAnalyze))
            {
                bool isConstante = false;
                if (!trimmedToAnalyze.StartsWith("{") || trimmedToAnalyze.StartsWith("["))
                {
                    isConstante = true;
                    if (trimmedToAnalyze.StartsWith("["))
                    {
                        if (!trimmedToAnalyze.EndsWith("]"))
                            throw new BindingTemplateException(string.Format("Cannot create constante BindingDefinition from '{0}': cannot find the closing ']'", toAnalyze));
                        bindingExpression = trimmedToAnalyze.Substring(1, trimmedToAnalyze.Length - 2);

                        int optionsEnd  = bindingExpression.IndexOf(':');
                        if (optionsEnd != -1)
                        {
                            string optionsString = bindingExpression.Substring(0, optionsEnd);
                            string[] optionsArray = optionsString.Split(';');
                            options = optionsArray.Where(p => !string.IsNullOrEmpty(p)).Select(p => p.Trim()).ToList();
                            bindingExpression = bindingExpression.Substring(optionsEnd + 1);
                        }
                    }
                    else
                        bindingExpression = toAnalyze;
                }
                else
                {
                    if (!trimmedToAnalyze.EndsWith("}"))
                        throw new BindingTemplateException(string.Format("Cannot create BindingDefinition from '{0}': cannot find the closing '}'", toAnalyze));
                    bindingExpression = trimmedToAnalyze.Substring(1, trimmedToAnalyze.Length - 2);

                    int postSep = bindingExpression.LastIndexOf(":");
                    if (postSep != -1)
                    {
                        string optionsString = bindingExpression.Substring(0, postSep);
                        string[] optionsArray = optionsString.Split(';');
                        //string[] optionsArray = optionsString.Split(new string[] { "::" }, StringSplitOptions.None);
                        options = optionsArray.Where(p => !string.IsNullOrEmpty(p)).Select(p => p.Trim()).ToList();
                        bindingExpression = bindingExpression.Substring(postSep + 1);
                    }
                }

                ret = new BindingDefinitionDescription(templateDefinition, bindingExpression, isConstante, options);
            }
            return ret;
        }
        #endregion
    }
}
