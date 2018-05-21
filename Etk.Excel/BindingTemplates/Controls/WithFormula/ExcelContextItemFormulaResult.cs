using System;
using System.Collections.Generic;
using System.ComponentModel;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Binding;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Etk.Excel.Application;

namespace Etk.Excel.BindingTemplates.Controls.WithFormula
{
    class ExcelContextItemFormulaResult : BindingContextItem, IBindingContextItemCanNotify, IExcelControl, IFormulaCalculation
    {
        #region properties and attributes
        private IEnumerable<INotifyPropertyChanged> objectsToNotify;
        private readonly ExcelBindingDefinitionFormulaResult excelBindingDefinitionFormulaResult;
        private object currentValue;

        public ExcelInterop.Range Range
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

        public void CreateControl(ExcelInterop.Range range)
        {
            Range = range;
        }

        public override void RealDispose()
        {
            OnPropertyChangedAction = null;
            if (objectsToNotify != null)
            {
                foreach (INotifyPropertyChanged obj in objectsToNotify)
                    obj.PropertyChanged -= OnPropertyChanged;
                objectsToNotify = null;
            }
            ExcelApplication.ReleaseComObject(Range);
        }

        public override object ResolveBinding()
        {
            if (excelBindingDefinitionFormulaResult.UseFormulaBindingDefinition != null)
                excelBindingDefinitionFormulaResult.UseFormulaBindingDefinition.UpdateDataSource(DataSource, (bool) Range.HasFormula);

            if (Range != null && Range.HasFormula)
                return Range.Formula;
            return excelBindingDefinitionFormulaResult.NestedBindingDefinition.ResolveBinding(DataSource);
        }

        public override bool UpdateDataSource(object data, out object retValue)
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
                if (BindingDefinition.MustNotify(DataSource, source, args))
                    OnPropertyChangedAction(this, OnPropertyChangedActionArgs);
            }
        }

        public void OnSheetCalculate()
        {
            if (!Range.HasFormula)
                return;

            object resolvedBinding =  excelBindingDefinitionFormulaResult.NestedBindingDefinition.ResolveBinding(DataSource);
            if (Range.FormulaLocal != resolvedBinding || ! object.Equals(Range.Value2, currentValue))
            {
                ExcelInterop.WorksheetFunction worksheetFunction = ETKExcel.ExcelApplication.Application.WorksheetFunction; 
                if (worksheetFunction.IsError(Range))
                { 
                    Type type = excelBindingDefinitionFormulaResult.NestedBindingDefinition.BindingType;
                    if (type != null)
                    {
                        object nullValue = type.IsValueType ? Activator.CreateInstance(type) : null;
                        excelBindingDefinitionFormulaResult.NestedBindingDefinition.UpdateDataSource(DataSource, nullValue);
                    }
                }
                else
                    excelBindingDefinitionFormulaResult.NestedBindingDefinition.UpdateDataSource(DataSource, Range.Value2);
                currentValue = Range.Value2;

                if (worksheetFunction != null)
                    ExcelApplication.ReleaseComObject(worksheetFunction);
            }
        } 
    }
}
