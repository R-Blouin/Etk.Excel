namespace Etk.Tests.Templates.ExcelDna1.Tests
{
    using System.Collections.Generic;
    using Etk.Excel.BindingTemplates.Views;

    interface IExcelTestTopic
    {
        string Description { get; }
        bool InitSuccessful { get; }
        string Exception { get; }
        string DestinationSheetName { get; }
        List<IExcelTest> Tests { get; }

        void Execute();
        void InitTestsStatus();
        void ExecuteTests();
    }
}
