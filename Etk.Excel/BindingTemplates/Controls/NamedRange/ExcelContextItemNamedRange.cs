namespace Etk.Excel.BindingTemplates.Controls.NamedRange
{
    using System.ComponentModel;
    using System.Runtime.InteropServices;
    using Etk.BindingTemplates.Context;
    using Etk.BindingTemplates.Definitions.Binding;
    using Etk.BindingTemplates.Views;
    using Microsoft.Office.Interop.Excel;

    class ExcelContextItemNamedRange : BindingContextItem, IBindingContextItemCanNotify, IExcelControl
    {
        #region properties and attributes
        private ExcelBindingDefinitionNamedRange excelBindingDefinitionNamedRange;
        private Worksheet workSheet;
        private string name;
        private Name rangeName;

        public Range Range
        { get; private set; }

        public IBindingContextItem NestedContextItem
        { get; private set; }

        public System.Action<IBindingContextItem, object> OnPropertyChangedAction
        {
            get 
            { 
                if(NestedContextItem == null || ! (NestedContextItem is IBindingContextItemCanNotify))
                    return null;
                return ((IBindingContextItemCanNotify) NestedContextItem).OnPropertyChangedAction;
            }
            set
            {
                if (NestedContextItem != null && NestedContextItem is IBindingContextItemCanNotify)
                    ((IBindingContextItemCanNotify) NestedContextItem).OnPropertyChangedAction = value;
            }
        }

        public object OnPropertyChangedActionArgs
        {
            get 
            { 
                if(NestedContextItem == null || ! (NestedContextItem is IBindingContextItemCanNotify))
                    return null;
                return ((IBindingContextItemCanNotify) NestedContextItem).OnPropertyChangedActionArgs;
            }
            set
            {
                if (NestedContextItem != null && NestedContextItem is IBindingContextItemCanNotify)
                    ((IBindingContextItemCanNotify) NestedContextItem).OnPropertyChangedActionArgs = value;
            } 
        }
        #endregion

        #region .ctors
        public ExcelContextItemNamedRange(IBindingContextElement parent, string name, IBindingDefinition bindingDefinition, IBindingContextItem nestedContextItem)
                                         : base(parent, bindingDefinition)
        {
            this.name = name;
            excelBindingDefinitionNamedRange = bindingDefinition as ExcelBindingDefinitionNamedRange;
            CanNotify = excelBindingDefinitionNamedRange.CanNotify;
            NestedContextItem = nestedContextItem;
        }
        #endregion

        #region public methods
        public void CreateControl(Range range)
        {
            this.Range = range;
            workSheet = this.Range.Worksheet;
            if (!string.IsNullOrEmpty(name))
            {
                Names names = null;
                try
                {
                    names = workSheet.Names;
                    rangeName = names.Add(name, this.Range);
                }
                catch (COMException ex)
                {
                    throw new EtkException(string.Format("Cannot create named caller '{0}': {1}", name, ex.Message));
                }
                finally
                { 
                    if(names != null)
                        Marshal.ReleaseComObject(names);
                }
            }

            if (NestedContextItem != null && NestedContextItem is IExcelControl)
                ((IExcelControl) NestedContextItem).CreateControl(range);
        }

        override public void RealDispose()
        {
            if (rangeName != null)
                rangeName.Delete();

            if (NestedContextItem != null)
                NestedContextItem.Dispose();

            if (workSheet != null)
            {
                Marshal.ReleaseComObject(workSheet);
                workSheet = null;
            }
            Range = null;
        }

        override public object ResolveBinding()
        {
            return NestedContextItem == null ? null : NestedContextItem.ResolveBinding();
        }

        override public bool UpdateDataSource(object data, out object retValue)
        {
            if(NestedContextItem != null)
                return NestedContextItem.UpdateDataSource(data, out retValue);
            else
            {
                retValue = null;
                return false;
            }
        }

        public void OnPropertyChanged(object source, PropertyChangedEventArgs args)
        {
            if (NestedContextItem != null && NestedContextItem is IBindingContextItemCanNotify)
                ((IBindingContextItemCanNotify) NestedContextItem).OnPropertyChanged(source, args);
        }
        #endregion
    }
}
