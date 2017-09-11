using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Binding;
using ExcelInterop = Microsoft.Office.Interop.Excel; 

namespace Etk.Excel.BindingTemplates.Controls.CheckBox
{
    class ExcelContextItemCheckBox : BindingContextItem, IBindingContextItemCanNotify, IExcelControl
    {
        #region attributes and properties
        private readonly ExcelBindingDefinitionCheckBox excelBindingDefinition;
        private ExcelCheckBox checkBox;
        private IEnumerable<INotifyPropertyChanged> objectsToNotify;

        public Action<IBindingContextItem, object> OnPropertyChangedAction
        { get; set; }

        public object OnPropertyChangedActionArgs
        { get; set; }

        public MethodInfo EnablePropertyGet
        { get; }

        public INotifyPropertyChanged EnableProperty
        { get; private set; }
        #endregion

        #region .ctors
        public ExcelContextItemCheckBox(IBindingContextElement parent, IBindingDefinition bindingDefinition)
            : base(parent, bindingDefinition)
        {
            if (bindingDefinition != null)
            {
                CanNotify =  bindingDefinition.CanNotify;
                excelBindingDefinition = bindingDefinition as ExcelBindingDefinitionCheckBox;

                if (excelBindingDefinition != null)
                {
                    objectsToNotify = excelBindingDefinition.GetObjectsToNotify(DataSource);
                    if (objectsToNotify != null)
                    {
                        foreach (INotifyPropertyChanged obj in objectsToNotify)
                            obj.PropertyChanged += OnPropertyChanged;
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

            checkBox?.Dispose();
        }

        public void OnPropertyChanged(object source, PropertyChangedEventArgs args)
        {
            if (objectsToNotify != null && OnPropertyChangedAction != null)
            {
                if (BindingDefinition.MustNotify(this.DataSource, source, args))
                    OnPropertyChangedAction(this, OnPropertyChangedActionArgs);
            }

        }

        public void CreateControl(ExcelInterop.Range range)
        {
            ExcelBindingDefinitionCheckBox definition = (ExcelBindingDefinitionCheckBox)BindingDefinition;
            checkBox = new ExcelCheckBox(range, definition.Definition);

            ResolveBinding();
            checkBox.SetOnClick(() => definition.UpdateDataSource(this.DataSource, checkBox.IsChecked));
        }

        public override object ResolveBinding()
        {
            bool value = false;
            {
                if (excelBindingDefinition != null)
                    value = (bool) excelBindingDefinition.ResolveBinding(DataSource);
            }
            if (checkBox != null)
                checkBox.IsChecked = value;
            return null;
        }
    }
}
