using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Binding;
using ExcelInterop = Microsoft.Office.Interop.Excel; 

namespace Etk.Excel.BindingTemplates.Controls.Picture
{
    class ExcelContextItemPicture : BindingContextItem, IBindingContextItemCanNotify, IExcelControl
    {
        #region attributes and properties
        private readonly ExcelBindingDefinitionPicture excelBindingDefinition;
        //private ExcelContextItemPicture picture;
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
        public ExcelContextItemPicture(IBindingContextElement parent, IBindingDefinition bindingDefinition)
            : base(parent, bindingDefinition)
        {
            if (bindingDefinition != null)
            {
                CanNotify =  bindingDefinition.CanNotify;
                excelBindingDefinition = bindingDefinition as ExcelBindingDefinitionPicture;

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

            //if (picture != null)
            //    picture.Dispose();
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
            //ExcelBindingDefinitionPicture definition = (ExcelBindingDefinitionPicture) BindingDefinition;
            //picture = new ExcelPicture(range, definition);

            ResolveBinding();
        }

        public override object ResolveBinding()
        {
            bool value = false;
            {
                if (excelBindingDefinition != null)
                    value = (bool) excelBindingDefinition.ResolveBinding(DataSource);
            }
            return null;
        }
    }
}
