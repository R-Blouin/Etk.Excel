using Etk.BindingTemplates.Context;
using ExcelInterop = Microsoft.Office.Interop.Excel; 

namespace Etk.Excel.BindingTemplates.Decorators
{
    class ExcelElementDecorator
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

        public void Resolve()
        {
            decorator.Resolve(range, contextElement);
        }
    }
}
