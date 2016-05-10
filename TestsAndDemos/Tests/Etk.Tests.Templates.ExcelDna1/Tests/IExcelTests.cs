namespace Etk.Tests.Templates.ExcelDna1.Tests
{
    using System.Collections.Generic;
    using Etk.Excel.BindingTemplates.Views;

    interface IExcelTests
    {
        string Description { get; }
        bool InitSuccessful { get; }
        string Exception { get; }
        List<IExcelTest> Tests { get; }

        void Execute();
    }
}
