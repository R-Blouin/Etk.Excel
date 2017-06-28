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
            object formula = null;
            if (excelBindingDefinitionWithFormula.FormulaBindingDefinition != null)
            {
                formula = excelBindingDefinitionWithFormula.FormulaBindingDefinition.ResolveBinding(DataSource);
                return formula != null ? "=" + formula : null;
            }
            return Range.HasFormula ? Range.Formula : null;
        }

        public override bool UpdateDataSource(object data, out object retValue)
        {
            retValue = null;
            if (excelBindingDefinitionWithFormula.FormulaBindingDefinition != null)
                retValue = excelBindingDefinitionWithFormula.FormulaBindingDefinition.UpdateDataSource(DataSource, data);
            else
                retValue = Range.HasFormula? Range.Formula : null;
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
            if (excelBindingDefinitionWithFormula.TargetBindingDefinition == null || !Range.HasFormula)
                return;

            if ((Range.HasFormula && Range.Formula != currentFormula) ||  ! object.Equals(Range.Value2, currentValue))
            {
                Microsoft.Office.Interop.Excel.WorksheetFunction worksheetFunction = ETKExcel.ExcelApplication.Application.WorksheetFunction;
                if (worksheetFunction.IsError(Range))
                {
                    Type type = excelBindingDefinitionWithFormula.TargetBindingDefinition.BindingType;
                    if (type != null)
                    {
                        object nullValue = type.IsValueType ? Activator.CreateInstance(type) : null;
                        excelBindingDefinitionWithFormula.TargetBindingDefinition.UpdateDataSource(DataSource, nullValue);
                    }
                }
                else
                    excelBindingDefinitionWithFormula.TargetBindingDefinition.UpdateDataSource(DataSource, Range.Value2);
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
