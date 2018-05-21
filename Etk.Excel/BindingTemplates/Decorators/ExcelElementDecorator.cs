using System;
using Etk.BindingTemplates.Context;
using Etk.Excel.Application;
using ExcelInterop = Microsoft.Office.Interop.Excel; 

namespace Etk.Excel.BindingTemplates.Decorators
{
    class ExcelElementDecorator : IDisposable
    {
        private readonly ExcelInterop.Range range;
        private readonly ExcelRangeDecorator decorator;
        private readonly IBindingContextElement contextElement;

        public ExcelElementDecorator(ExcelInterop.Range range, ExcelRangeDecorator decorator, IBindingContextElement contextElement)
        {
            this.range = range;
            this.decorator = decorator;
            this.contextElement = contextElement;
        }

        public void Dispose()
        {
            ExcelApplication.ReleaseComObject(range);
        }

        public void Resolve()
        {
            decorator.Resolve(range, contextElement);
        }
    }
}
