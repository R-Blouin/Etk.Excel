using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Etk.BindingTemplates.Definitions.Decorators;
using Etk.BindingTemplates.Definitions.EventCallBacks;
using Etk.Tools.Reflection;

namespace Etk.BindingTemplates.Definitions.Binding
{
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

        public MethodInfo MultiLineFactorResolver
        { get; private set; }

        #region .ctors and factories
        public BindingDefinitionDescription()
        {}

        private BindingDefinitionDescription(string bindingExpression, bool isConst, List<string> options)
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
                            OnSelection = EventCallbacksManager.GetCallback(methodInfoIdent);
                            continue;
                        }
                        // On double left click
                        if (option.StartsWith("LDC="))
                        {
                            string methodInfoIdent = option.Substring(4);
                            OnLeftDoubleClick = EventCallbacksManager.GetCallback(methodInfoIdent);
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
                                    MultiLineFactorResolver = TypeHelpers.GetMethod(null, multiLineFactorResolver);
                                    if (MultiLineFactorResolver != null)
                                    {
                                        int parametersCpt = MultiLineFactorResolver.GetParameters().Length;
                                        if (MultiLineFactorResolver.ReturnType != typeof(int) || 
                                            parametersCpt > 1 ||
                                            (parametersCpt == 1 && !(MultiLineFactorResolver.GetParameters()[0].ParameterType.IsAssignableFrom(typeof(object)))))
                                            throw new Exception("The function prototype must be  be defined as 'int <Function Name>([param]) with 'param' inheriting from 'system.object'");
                                        IsMultiLine = true;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    throw new Exception(string.Format("Cannot resolve the 'ME' attribute for the binding definition '{0}'", bindingExpression), ex);
                                }
                            }
                        }
                    }
                }
            }
        }

        public static BindingDefinitionDescription CreateBindingDescription(string toAnalyze, string trimmedToAnalyze)
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
                            string[] optionsArray = optionsString.Split(',');
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

                    int compoStart = bindingExpression.IndexOf('{');
                    string toAnalyzeOptions = compoStart == -1 ? bindingExpression : bindingExpression.Substring(0, compoStart);

                    string[] parts = toAnalyzeOptions.Split(':');
                    if (parts.Count() > 2)
                        throw new BindingTemplateException(string.Format("Cannot create BindingDefinition from '{0}': options not properly set. Syntax is: {opt1,opt2...:...}", toAnalyze));
                    if (parts.Count() == 2)
                    {
                        string optionsString = parts[0];
                        string[] optionsArray = optionsString.Split(';');
                        options = optionsArray.Where(p => !string.IsNullOrEmpty(p)).Select(p => p.Trim()).ToList();
                        if (compoStart == -1)
                            bindingExpression = parts[1];
                        else
                            bindingExpression = bindingExpression.Substring(compoStart);
                    }
                }

                ret = new BindingDefinitionDescription(bindingExpression, isConstante, options);
            }
            return ret;
        }
        #endregion
    }
}
