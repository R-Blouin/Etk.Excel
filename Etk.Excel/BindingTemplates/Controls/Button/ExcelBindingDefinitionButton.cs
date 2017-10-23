using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.Excel.BindingTemplates.Definitions;
using Etk.Tools.Extensions;
using Etk.Tools.Reflection;
using Etk.BindingTemplates.Definitions.EventCallBacks;

namespace Etk.Excel.BindingTemplates.Controls.Button
{
    class ExcelBindingDefinitionButton : BindingDefinition
    {
        #region attributes and properties
        private static EventCallbacksManager eventCallbacksManager;
        private static EventCallbacksManager EventCallbacksManager => eventCallbacksManager ??
                                                                      (eventCallbacksManager = CompositionManager.Instance.GetExportedValue<EventCallbacksManager>());

        public const string BUTTON_TEMPLATE_PREFIX = "<Button";

        public IBindingDefinition LabelBindingDefinition
        { get; private set; }

        public ExcelButtonDefinition Definition
        { get; private set; }

        public ExcelTemplateDefinition TemplateDefinition
        { get; private set; }

        public EventCallback Command
        { get; private set; }

        public bool OnClickWithRange
        { get; private set; }

        public PropertyInfo EnablePropertyInfo
        { get; private set; }
        #endregion

        #region .ctors and factories
        private ExcelBindingDefinitionButton(BindingDefinitionDescription definitionDescription, ExcelTemplateDefinition templateDefinition, ExcelButtonDefinition definition)
                                            : base(definitionDescription)
        {
            TemplateDefinition = templateDefinition;
            Definition = definition;
            if (! string.IsNullOrEmpty(definition.Label))
            {
                LabelBindingDefinition = BindingDefinitionFactory.CreateInstances(templateDefinition, DefinitionDescription);
                CanNotify = LabelBindingDefinition.CanNotify; 
            }
            RetrieveOnClickMethod();
            RetrieveEnableProperty();
        }

        public static ExcelBindingDefinitionButton CreateInstance(ExcelTemplateDefinition templateDefinition, string definition)
        {
            ExcelBindingDefinitionButton ret = null;
            if (! string.IsNullOrEmpty(definition))
            {
                try
                {
                    ExcelButtonDefinition excelButtonDefinition = definition.Deserialize<ExcelButtonDefinition>();
                    BindingDefinitionDescription definitionDescription = BindingDefinitionDescription.CreateBindingDescription(templateDefinition, excelButtonDefinition.Label, excelButtonDefinition.Label);
                    ret = new ExcelBindingDefinitionButton(definitionDescription, templateDefinition, excelButtonDefinition);
                }
                catch (Exception ex)
                {
                    string message = $"Cannot retrieve the button dataAccessor '{definition.EmptyIfNull()}'. {ex.Message}";
                    throw new EtkException(message);
                }
            }
            return ret;
        }
        #endregion

        #region public methods
        public override IBindingContextItem ContextItemFactory(IBindingContextElement parent)
        {
            return new ExcelContextItemButton(parent, this);
        }

        public override object UpdateDataSource(object dataSource, object data)
        {
            return null;
        }

        public override object ResolveBinding(object dataSource)
        {
            return LabelBindingDefinition == null ? null : LabelBindingDefinition.ResolveBinding(dataSource);
        }

        public override IEnumerable<INotifyPropertyChanged> GetObjectsToNotify(object dataSource)
        {
            return LabelBindingDefinition == null ? null : LabelBindingDefinition.GetObjectsToNotify(dataSource);
        }

        public override bool MustNotify(object dataSource, object source, PropertyChangedEventArgs args)
        {
            return LabelBindingDefinition == null ? false : LabelBindingDefinition.MustNotify(dataSource, source, args);
        }
        #endregion

        #region private methods
        private void RetrieveOnClickMethod()
        {
            if (!string.IsNullOrEmpty(Definition.Command))
            {
                try
                {
                    string onCommand = Definition.Command.Trim();
                    Command =  EventCallbacksManager.RetrieveCallback(TemplateDefinition, onCommand);

                    if(! Command.IsNotDotNet)
                    {
                        ParameterInfo[] parameters = Command.Callback.GetParameters();
                        if (Command.Callback.IsStatic)
                        {
                            if (parameters.Count() > 2)
                                throw new EtkException($"Method dataAccessor must be 'void static {Command.Callback.Name}(object currentObject [, Range <currentObject caller>]'");

                            OnClickWithRange = parameters.Count() == 2;
                        }
                        else
                        {
                            if (parameters.Count() > 1 || (parameters.Count() == 1 && parameters[0].ParameterType != typeof(Microsoft.Office.Interop.Excel.Range)))
                                throw new EtkException($"Method dataAccessor must be 'void {Command.Callback.Name}([Range <currentObject caller>])'");

                            OnClickWithRange = parameters.Count() == 1;
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new EtkException($"Get 'Command' methodInfo information failed:{ex.Message}");
                }
            }
        }

        private void RetrieveEnableProperty()
        {
            if (!string.IsNullOrEmpty(Definition.EnableProp))
            {
                try
                {
                    string enableProp = Definition.EnableProp.Trim();
                    string[] enablePropElements = enableProp.Split(',');
                    if (enablePropElements.Count() != 1 && enablePropElements.Count() != 3)
                        throw new ArgumentException("The 'EnableProp' separator is ',' and it must be composed this way 'Assembly,Type,Property' or, if the property is part of the calling Instance, 'Property'");

                    Type type;
                    string propertyName;
                    if (enablePropElements.Count() == 1)
                    {
                        propertyName = enablePropElements[0].EmptyIfNull().Trim();
                        type = TemplateDefinition.MainBindingDefinition.BindingTypeIsGeneric ? TemplateDefinition.MainBindingDefinition.BindingGenericType : TemplateDefinition.MainBindingDefinition.BindingType;
                    }
                    else
                    {
                        propertyName = enablePropElements[2].EmptyIfNull().Trim();
                        if (!string.IsNullOrEmpty(enablePropElements[0]) && !string.IsNullOrEmpty(enablePropElements[1]))
                            type = TypeHelpers.GetType(enablePropElements[1], enablePropElements[0]);
                        else
                            type = TemplateDefinition.MainBindingDefinition.BindingTypeIsGeneric ? TemplateDefinition.MainBindingDefinition.BindingGenericType : TemplateDefinition.MainBindingDefinition.BindingType;
                    }


                    if (!string.IsNullOrEmpty(enablePropElements[0]) && !string.IsNullOrEmpty(enablePropElements[1]))
                        type = TypeHelpers.GetType(enablePropElements[1], enablePropElements[0]);
                    else
                        type = TemplateDefinition.MainBindingDefinition.BindingTypeIsGeneric ? TemplateDefinition.MainBindingDefinition.BindingGenericType : TemplateDefinition.MainBindingDefinition.BindingType;

                    EnablePropertyInfo = type.GetProperty(enablePropElements[2]);
                    if (EnablePropertyInfo == null)
                        throw new ArgumentException($"Property '{enablePropElements[2]}' not found");
                }
                catch (Exception ex)
                {
                    throw new EtkException($"Get 'EnableProp' property information failed:{ex.Message}");
                }
            }
        }
        #endregion
    }
}
