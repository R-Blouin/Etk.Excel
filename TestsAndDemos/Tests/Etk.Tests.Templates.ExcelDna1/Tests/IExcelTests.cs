namespace Etk.Tests.Templates.ExcelDna1.Tests
{
    using System.Collections.Generic;
    using Etk.Excel.BindingTemplates.Views;

    interface IExcelTests
    {
        string Description { get; }
        bool InitSuccessful { get; }
        string Exception { get; }
        IEnumerable<IExcelTest> Tests { get; }

        void Init();
        void Execute();
    }
}
