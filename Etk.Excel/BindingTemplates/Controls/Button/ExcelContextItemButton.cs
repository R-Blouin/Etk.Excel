namespace Etk.Excel.BindingTemplates.Controls.Button
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Reflection;
    using Etk.BindingTemplates.Context;
    using Etk.BindingTemplates.Definitions.Binding;
    using Microsoft.Office.Interop.Excel;

    class ExcelContextItemButton : BindingContextItem, IBindingContextItemCanNotify, IExcelControl
    {
        #region attributes and properties
        private ExcelBindingDefinitionButton excelBindingDefinitionButton;
        private ExcelButton button;
        private IEnumerable<INotifyPropertyChanged> objectsToNotify;

        public Action<IBindingContextItem, object> OnPropertyChangedAction
        { get; set; }

        public object OnPropertyChangedActionArgs
        { get; set; }

        public MethodInfo EnablePropertyGet
        { get; private set; }

        public INotifyPropertyChanged EnableProperty
        { get; private set; }
        #endregion

        #region .ctors
        public ExcelContextItemButton(IBindingContextElement parent, IBindingDefinition bindingDefinition)
                                     : base(parent, bindingDefinition)
        {
            CanNotify = bindingDefinition == null ? false : bindingDefinition.CanNotify;
            excelBindingDefinitionButton = bindingDefinition as ExcelBindingDefinitionButton;

            if (CanNotify)
            {
                objectsToNotify = excelBindingDefinitionButton.GetObjectsToNotify(DataSource);
                if (objectsToNotify != null)
                {
                    foreach (INotifyPropertyChanged obj in objectsToNotify)
                        obj.PropertyChanged += OnPropertyChanged;
                }
            }

            if (excelBindingDefinitionButton.EnablePropertyInfo != null)
            {
                EnablePropertyGet = excelBindingDefinitionButton.EnablePropertyInfo.GetGetMethod();
                if (EnablePropertyGet != null)
                {
                    if (! EnablePropertyGet.IsStatic)
                    {
                        if (ParentElement.DataSource != null && ParentElement.DataSource is INotifyPropertyChanged)
                        {
                            EnableProperty = (INotifyPropertyChanged)ParentElement.DataSource;
                            EnableProperty.PropertyChanged += OnPropertyChanged;
                        }
                    }
                }
            }
        }
        #endregion

        public override void RealDispose()
        {           
            OnPropertyChangedAction = null;

            if (objectsToNotify != null)
            {
                foreach (INotifyPropertyChanged obj in objectsToNotify)
                    obj.PropertyChanged -= OnPropertyChanged;
                objectsToNotify = null;
            }

            if (EnableProperty != null)
            {
                EnableProperty.PropertyChanged -= OnPropertyChanged;
                EnableProperty = null;
            }

            if (button != null)
                button.Dispose();
        }

        public void OnPropertyChanged(object source, PropertyChangedEventArgs args)
        {
            if (objectsToNotify != null && OnPropertyChangedAction != null)
            {
                if (BindingDefinition.MustNotify(this.DataSource, source, args))
                    OnPropertyChangedAction(this, OnPropertyChangedActionArgs);
            }

            if (button != null && EnableProperty != null && args.PropertyName.Equals(excelBindingDefinitionButton.EnablePropertyInfo.Name))
            { 
                if(EnablePropertyGet.IsStatic)
                    button.Enable = (bool) EnablePropertyGet.Invoke(null, null);
                else
                    button.Enable = (bool) EnablePropertyGet.Invoke(source, null);
            }
        }
    
        public void  CreateControl(Range range)
        {
            ExcelBindingDefinitionButton definition = (ExcelBindingDefinitionButton)BindingDefinition;
            button = new ExcelButton(range, definition.Definition);
            bool isStatic = definition.Command == null ? false : definition.Command.IsStatic;
            button.SetCommand(definition.Command, base.ParentElement.DataSource, definition.OnClickWithRange);

            ResolveBinding();

            if (EnableProperty != null)
            {
                if (EnablePropertyGet.IsStatic)
                    button.Enable = (bool)EnablePropertyGet.Invoke(null, null);
                else
                    button.Enable = (bool)EnablePropertyGet.Invoke(base.ParentElement.DataSource, null);
            }
        }

        public override object ResolveBinding()
        {
            object value = null;
            if (excelBindingDefinitionButton != null)
                value = excelBindingDefinitionButton.ResolveBinding(DataSource);
            if (button != null)
                button.Text = value == null ? string.Empty : value.ToString();
            return null;
        }
    }
}
