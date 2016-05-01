namespace Etk.Tests.Templates.ExcelDna1.Tests
{
    interface IExcelTest
    {
        string Description{ get; }
        bool Success{ get; }
        bool Done{ get; }
        string Errors { get; }

        void Execute();
    }
}
