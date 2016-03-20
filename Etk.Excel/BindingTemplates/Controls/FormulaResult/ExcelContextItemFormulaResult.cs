namespace Etk.Excel.BindingTemplates.Controls.FormulaResult
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Runtime.InteropServices;
    using Etk.BindingTemplates.Context;
    using Etk.BindingTemplates.Definitions.Binding;
    using Microsoft.Office.Interop.Excel;

    class ExcelContextItemFormulaResult : BindingContextItem, IBindingContextItemCanNotify, IExcelControl, ISheetCalculate
    {
        #region properties and attributes
        private IEnumerable<INotifyPropertyChanged> objectsToNotify;
        private ExcelBindingDefinitionFormulaResult excelBindingDefinitionFormulaResult;
        private Worksheet workSheet;
        private Application application;
        private object currentValue;

        public Range Range
        { get; private set; }

        public Action<IBindingContextItem, object> OnPropertyChangedAction
        { get; set; }

        public object OnPropertyChangedActionArgs
        { get; set; }     
        #endregion

        #region .ctors
        public ExcelContextItemFormulaResult(IBindingContextElement parent, IBindingDefinition bindingDefinition)
                                            : base(parent, bindingDefinition)
        {
            excelBindingDefinitionFormulaResult = bindingDefinition as ExcelBindingDefinitionFormulaResult;
            CanNotify = excelBindingDefinitionFormulaResult.CanNotify;
            //NestedContextItem = nestedContextItem;

            if (CanNotify)
            {
                objectsToNotify = bindingDefinition.GetObjectsToNotify(DataSource);
                if (objectsToNotify != null)
                {
                    foreach (INotifyPropertyChanged obj in objectsToNotify)
                        obj.PropertyChanged += OnPropertyChanged;
                }
            }
        }
        #endregion

        public void CreateControl(Range range)
        {
            this.Range = range;
            application = this.Range.Application;
            workSheet = this.Range.Worksheet;
        }

        override public void RealDispose()
        {
            OnPropertyChangedAction = null;
            if (objectsToNotify != null)
            {
                foreach (INotifyPropertyChanged obj in objectsToNotify)
                    obj.PropertyChanged -= OnPropertyChanged;
                objectsToNotify = null;
            }
            if (workSheet != null)
            {
                Marshal.ReleaseComObject(workSheet);
                workSheet = null;
            }
            if (application != null)
            {
                Marshal.ReleaseComObject(application);
                application = null;
            }
            Range = null;
        }

        override public object ResolveBinding()
        {
            if (excelBindingDefinitionFormulaResult.UseFormulaBindingDefinition != null)
                excelBindingDefinitionFormulaResult.UseFormulaBindingDefinition.UpdateDataSource(this.DataSource, (bool)Range.HasFormula);

            if (Range != null && Range.HasFormula)
                return Range.Formula;
            else
                return excelBindingDefinitionFormulaResult.NestedBindingDefinition.ResolveBinding(this.DataSource);
        }

        override public bool UpdateDataSource(object data, out object retValue)
        {
            if (Range.HasFormula)
                retValue = Range.Value2;
            else
                retValue = excelBindingDefinitionFormulaResult.NestedBindingDefinition.UpdateDataSource(this.DataSource, data);
            return true;
        }

        public void OnPropertyChanged(object source, PropertyChangedEventArgs args)
        {
            if (CanNotify && objectsToNotify != null && OnPropertyChangedAction != null)
            {
                if (BindingDefinition.MustNotify(this.DataSource, source, args))
                    OnPropertyChangedAction(this, OnPropertyChangedActionArgs);
            }
        }

        public void OnSheetCalculate()
        {
            if (Range.HasFormula && ! object.Equals(Range.Value2, currentValue))
            {
                if (application.WorksheetFunction.IsError(Range))
                { 
                    Type type = excelBindingDefinitionFormulaResult.NestedBindingDefinition.BindingType;
                    object nullValue = type.IsValueType ? Activator.CreateInstance(type) : null;
                    excelBindingDefinitionFormulaResult.NestedBindingDefinition.UpdateDataSource(this.DataSource, nullValue);
                }
                else
                    excelBindingDefinitionFormulaResult.NestedBindingDefinition.UpdateDataSource(this.DataSource, Range.Value2);
                currentValue = Range.Value2;
            }
        } 
    }
}
