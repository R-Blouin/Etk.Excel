namespace Etk.Excel.RequestManagement.Definitions
{
    using Etk.Excel.BindingTemplates.Views;

    class ExcelRequestDefinition
    {
        #region properties
        public string Name
        { get; private set; }

        public string Description
        { get; private set; }

        public ExcelTemplateView View
        { get; private set; }
        #endregion

        public ExcelRequestDefinition(string name, string description, ExcelTemplateView view)
        {
            Name = name;
            Description = description;
            View = view;
        }
    }
}
