namespace Etk.Tests.Templates.ExcelDna1.Tests
{
    using System.Collections.Generic;
    using Etk.Excel.BindingTemplates.Views;

    interface IExcelTestTopic
    {
        int Id { get; }
        string Description { get; }
        bool RenderSuccessful { get; }
        string Exception { get; }
        string DestinationSheetName { get; }
        List<IExcelTest> Tests { get; }

        void Init();
        void Execute();
        void ExecuteTests();
    }
}
