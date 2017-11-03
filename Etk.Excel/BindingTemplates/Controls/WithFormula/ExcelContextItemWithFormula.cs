using System;
using System.Collections.Generic;
using System.ComponentModel;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.Excel.Application;

namespace Etk.Excel.BindingTemplates.Controls.WithFormula
{
    class ExcelContextItemWithFormula : BindingContextItem, IBindingContextItemCanNotify, IExcelControl, IFormulaCalculation
    {
        #region properties and attributes
        private IEnumerable<INotifyPropertyChanged> objectsToNotify;
        private readonly ExcelBindingDefinitionWithFormula excelBindingDefinitionWithFormula;
        private object currentValue;
        private object currentFormula;
        private IBindingContextItem formulaBindingContext = null;

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

            if (excelBindingDefinitionWithFormula.FormulaBindingDefinition != null)
                formulaBindingContext = excelBindingDefinitionWithFormula.FormulaBindingDefinition.ContextItemFactory(parent);

            if (CanNotify)
            {
                objectsToNotify = excelBindingDefinitionWithFormula.GetObjectsToNotify(DataSource);
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
            if(excelBindingDefinitionWithFormula.FormulaBindingDefinition != null)
            {
                string ret = $"={formulaBindingContext.ResolveBinding()}";
                return ret;

            }
            else if (excelBindingDefinitionWithFormula.TargetBindingDefinition != null)
                return excelBindingDefinitionWithFormula.TargetBindingDefinition.ResolveBinding(DataSource);

            //if (excelBindingDefinitionWithFormula.TargetBindingDefinition != null)
            //    return excelBindingDefinitionWithFormula.TargetBindingDefinition.ResolveBinding(DataSource);
            return null;
        }

        public override bool UpdateDataSource(object data, out object retValue)
        {
            excelBindingDefinitionWithFormula.TargetBindingDefinition?.UpdateDataSource(DataSource, data);

            if (data == null) // If null enter => ResolveBinding the binding
                retValue = ResolveBinding();
            else
            {
                if (data.ToString().Trim().StartsWith("="))
                    retValue = data.ToString();
                else
                    retValue = data;
            }
            return true;
        }

        public void OnPropertyChanged(object source, PropertyChangedEventArgs args)
        {
            if (CanNotify && objectsToNotify != null && OnPropertyChangedAction != null)
            {
                if (excelBindingDefinitionWithFormula.MustNotify(DataSource, source, args))
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
                    ExcelApplication.ReleaseComObject(worksheetFunction);
                    worksheetFunction = null;
                }
            }
        }
    }
}
