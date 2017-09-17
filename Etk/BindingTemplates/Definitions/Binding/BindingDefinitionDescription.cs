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

        public EventCallback OnSelection
        { get; private set; }

        public EventCallback OnLeftDoubleClick
        { get; private set; }

        public bool IsMultiLine
        { get; private set; }

        public double MultiLineFactor
        { get; private set; }

        public EventCallback MultiLineFactorResolver
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
                            string methodInfoName = option.Substring(2);
                            OnSelection = RetrieveMethodInfo(templateDefinition, option, methodInfoName);
                            continue;
                        }
                        // On double left click
                        if (option.StartsWith("LDC="))
                        {
                            string methodInfoName = option.Substring(4);
                            OnLeftDoubleClick = RetrieveMethodInfo(templateDefinition, option, methodInfoName);
                            continue;
                        }
                        // MultiLine based on the number of line passed as parameter
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
                            string methodInfoName = option.Substring(3);
                            if (!string.IsNullOrEmpty(methodInfoName))
                            {
                                try
                                {
                                    MultiLineFactorResolver = RetrieveMethodInfo(templateDefinition, null, methodInfoName);
                                    if (MultiLineFactorResolver != null)
                                    {
                                        if (! MultiLineFactorResolver.IsNotDotNet)
                                        {
                                            int parametersCpt = MultiLineFactorResolver.Callback.GetParameters().Length;
                                            if (MultiLineFactorResolver.Callback.ReturnType != typeof(int) || parametersCpt > 1 || (parametersCpt == 1 && !(MultiLineFactorResolver.Callback.GetParameters()[0].ParameterType.IsAssignableFrom(typeof(object)))))
                                                throw new Exception("The function prototype must be defined as 'int <Function Name>([param]) with 'param' inheriting from 'system.object'");
                                        }
                                        IsMultiLine = true;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    throw new Exception($"Cannot resolve the 'ME' attribute for the binding definition '{bindingExpression}'", ex);
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
                        // On double left click, Show/Hide the x following/preceding colums/rows. Start shown
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
                            throw new Exception("The 'Show/Hide' columns prototype must be defined as 'SH=<int>' where '<int>' is a integer");
                        }
                    }
                }
            }
        }

        protected EventCallback RetrieveMethodInfo(ITemplateDefinition templateDefinition, string option, string callbackName)
        {
            try
            {
                return EventCallbacksManager.RetrieveCallback(templateDefinition, callbackName);
            }
            catch (Exception ex)
            {
                throw new Exception($"Property option '{option}'. {ex.Message}");
            }
        }

        public static BindingDefinitionDescription CreateBindingDescription(ITemplateDefinition templateDefinition, string toAnalyze, string trimmedToAnalyze)
        {
            BindingDefinitionDescription ret = null;
            List<string> options = null;
            if (!string.IsNullOrEmpty(trimmedToAnalyze))
            {
                string bindingExpression;
                bool isConstante = false;
                // Constante
                if (!trimmedToAnalyze.StartsWith("{") || trimmedToAnalyze.StartsWith("["))
                {
                    isConstante = true;
                    if (trimmedToAnalyze.StartsWith("["))
                    {
                        if (!trimmedToAnalyze.EndsWith("]"))
                            throw new BindingTemplateException($"Cannot create constante BindingDefinition from '{toAnalyze}': cannot find the closing ']'");
                        bindingExpression = trimmedToAnalyze.Substring(1, trimmedToAnalyze.Length - 2);

                        int postSep = bindingExpression.LastIndexOf("::");
                        if (postSep != -1)
                        {
                            string optionsString = bindingExpression.Substring(0, postSep);
                            string[] optionsArray = optionsString.Split(';');
                            options = optionsArray.Where(p => !string.IsNullOrEmpty(p)).Select(p => p.Trim()).ToList();
                            bindingExpression = bindingExpression.Substring(postSep + 2);
                        }
                    }
                    else
                        bindingExpression = toAnalyze;
                }
                // No Constante
                else
                {
                    if (!trimmedToAnalyze.EndsWith("}"))
                        throw new BindingTemplateException($"Cannot create BindingDefinition from '{toAnalyze}': cannot find the closing '}}'");
                    bindingExpression = trimmedToAnalyze.Substring(1, trimmedToAnalyze.Length - 2);

                    int postSep = bindingExpression.LastIndexOf("::");
                    if (postSep != -1)
                    {
                        string optionsString = bindingExpression.Substring(0, postSep);
                        string[] optionsArray = optionsString.Split(';');
                        options = optionsArray.Where(p => !string.IsNullOrEmpty(p)).Select(p => p.Trim()).ToList();
                        bindingExpression = bindingExpression.Substring(postSep + 2);
                    }
                    else if (bindingExpression.StartsWith("=")) // USe for Formula not bind with the model
                    {
                        options = new List<string>(new []{ "F" + bindingExpression });
                        bindingExpression = string.Empty;
                    }
                }

                ret = new BindingDefinitionDescription(templateDefinition, bindingExpression, isConstante, options);
            }
            return ret;
        }
        #endregion
    }
}
