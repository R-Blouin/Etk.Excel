using System;
using System.Collections.Generic;
using System.ComponentModel;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.Excel.Application;
using ExcelInterop = Microsoft.Office.Interop.Excel;

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

        private ExcelInterop.Range range;

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

        public void CreateControl(ExcelInterop.Range range)
        {
            this.range = range[1, 1];
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
            ExcelApplication.ReleaseComObject(range);
        }

        public override object ResolveBinding()
        {
            //if (excelBindingDefinitionWithFormula.TargetBindingDefinition != null)
            //    excelBindingDefinitionWithFormula.TargetBindingDefinition.ResolveBinding(DataSource);

            if (excelBindingDefinitionWithFormula.FormulaBindingDefinition != null)
            {
                string ret = $"={formulaBindingContext.ResolveBinding()}";
                //if(formulaValue == null)
                //    formulaValue = ret;
                //else if (formulaValue != ret)
                //{
                //    formulaValue = ret;

                //    //using (FreezeExcel freeze = new FreezeExcel(true, true, false, false))
                //    {
                //        XlCalculation calculationMode = ETKExcel.ExcelApplication.Application.Calculation;

                //        //Range.Formula = formulaValue;
                //        //ETKExcel.ExcelApplication.Application.Calculation = XlCalculation.xlCalculationManual;
                //        //(Range.Parent as Worksheet).Calculate();
                //        Range.Calculate();

                //        ETKExcel.ExcelApplication.Application.Calculation = calculationMode;
                //    }

                //}
                return ret;
            }
            return null;
        }

        public void UpdateTarget(object data)
        {
            excelBindingDefinitionWithFormula.TargetBindingDefinition?.UpdateDataSource(DataSource, data);
        }

        public override bool UpdateDataSource(object data, out object retValue)
        {
            UpdateTarget(data);

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
            if ((range.HasFormula && range.Formula != currentFormula) ||  ! object.Equals(range.Value2, currentValue))
            {
                ExcelInterop.WorksheetFunction worksheetFunction = ETKExcel.ExcelApplication.Application.WorksheetFunction;
                if (range.HasFormula && worksheetFunction.IsError(range))
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
                        excelBindingDefinitionWithFormula.TargetBindingDefinition.UpdateDataSource(DataSource, range.Value2);
                }

                currentValue = range.Value2;
                currentFormula = range.HasFormula ? range.Formula : null;

                if (worksheetFunction != null)
                    ExcelApplication.ReleaseComObject(worksheetFunction);
            }
        }
    }
}
