using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Binding;

namespace Etk.Excel.BindingTemplates.Controls.WithFormula
{
    class ExcelContextItemWithFormula : BindingContextItem, IBindingContextItemCanNotify, IExcelControl, IFormulaCalculation
    {
        #region properties and attributes
        private IEnumerable<INotifyPropertyChanged> objectsToNotify;
        private readonly ExcelBindingDefinitionWithFormula excelBindingDefinitionWithFormula;
        private object currentValue;
        private object currentFormula;
        private bool useFormula;

        public Microsoft.Office.Interop.Excel.Range Range
        { get; private set; }

        public Action<IBindingContextItem, object> OnPropertyChangedAction
        { get; set; }

        public object OnPropertyChangedActionArgs
        { get; set; }
        #endregion

        #region .ctors
        public ExcelContextItemWithFormula(IBindingContextElement parent, IBindingDefinition bindingDefinition)
                                          : base(parent, bindingDefinition)
        {
            excelBindingDefinitionWithFormula = bindingDefinition as ExcelBindingDefinitionWithFormula;
            CanNotify = excelBindingDefinitionWithFormula.CanNotify;
            useFormula = excelBindingDefinitionWithFormula.FormulaBindingDefinition != null;

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

        public void CreateControl(Microsoft.Office.Interop.Excel.Range range)
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
            Range = null;
        }

        public override object ResolveBinding()
        {
            if (useFormula)
            {
                if(excelBindingDefinitionWithFormula.FormulaBindingDefinition != null)
                    return "=" + excelBindingDefinitionWithFormula.FormulaBindingDefinition.ResolveBinding(DataSource);
                return Range.HasFormula ? Range.Formula : null;
            }

            if (excelBindingDefinitionWithFormula.TargetBindingDefinition != null)
                return excelBindingDefinitionWithFormula.TargetBindingDefinition.ResolveBinding(DataSource);
            return null;
        }

        public override bool UpdateDataSource(object data, out object retValue)
        {
            retValue = data;
            if (data == null) // If null enter => ResolveBinding the binding
            {
                useFormula = excelBindingDefinitionWithFormula.FormulaBindingDefinition != null;
                if (excelBindingDefinitionWithFormula.FormulaBindingDefinition != null)
                    retValue = "=" + excelBindingDefinitionWithFormula.FormulaBindingDefinition.ResolveBinding(DataSource);
                else if (excelBindingDefinitionWithFormula.TargetBindingDefinition != null)
                    retValue = excelBindingDefinitionWithFormula.TargetBindingDefinition.ResolveBinding(DataSource);
            }
            else
            {
                useFormula = data.ToString().Trim().StartsWith("=");
                if (useFormula)
                {
                    try
                    {
                        Range.Formula = data.ToString();
                        retValue = Range.Formula;
                    }
                    catch
                    { }
                }
                else if(excelBindingDefinitionWithFormula.TargetBindingDefinition != null)
                    retValue = excelBindingDefinitionWithFormula.TargetBindingDefinition.UpdateDataSource(DataSource, data);
            }
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
            if ((Range.HasFormula && Range.Formula != currentFormula) ||  ! object.Equals(Range.Value2, currentValue))
            {
                Microsoft.Office.Interop.Excel.WorksheetFunction worksheetFunction = ETKExcel.ExcelApplication.Application.WorksheetFunction;
                if (Range.HasFormula && worksheetFunction.IsError(Range))
                {
                    if (excelBindingDefinitionWithFormula.TargetBindingDefinition != null)
                    {
                        Type type = excelBindingDefinitionWithFormula.TargetBindingDefinition.BindingType;
                        if (type != null)
                        {
                            object nullValue = type.IsValueType ? Activator.CreateInstance(type) : null;
                            excelBindingDefinitionWithFormula.TargetBindingDefinition.UpdateDataSource(DataSource, nullValue);
                        }
                    }
                }
                else
                {
                    if(excelBindingDefinitionWithFormula.TargetBindingDefinition != null)
                        excelBindingDefinitionWithFormula.TargetBindingDefinition.UpdateDataSource(DataSource, Range.Value2);
                }

                currentValue = Range.Value2;
                currentFormula = Range.HasFormula ? Range.Formula : null;

                if (worksheetFunction != null)
                {
                    Marshal.ReleaseComObject(worksheetFunction);
                    worksheetFunction = null;
                }
            }
        }
    }
}
