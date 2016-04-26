using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Interop;
using Etk.Excel.Application;
using Etk.Excel.BindingTemplates;
using Etk.Excel.BindingTemplates.Views;
using Etk.Excel.ContextualMenus;
using Etk.Excel.Extensions;
using Etk.Excel.RequestManagement.Definitions;
using Etk.Excel.UI.Windows.ModelManagement;
using Etk.Excel.UI.Windows.ModelManagement.ViewModels;

namespace Etk.Excel.RequestManagement
{
    [Export]
    [PartCreationPolicy(CreationPolicy.Shared)]
    class RequestsManager : IRequestManager, IDisposable
    {
        #region attributes and properties
        private IContextualMenu RequestManagementMenu;

        ExcelRequestDefinition test;
        int xOffset = 0;
        int yOffset = 0;  
        #endregion

        #region .ctors
        [ImportingConstructor]
        public RequestsManager([Import] ContextualMenuManager contextualMenuManager)
        {
            try
            {
                // Create the contextual menu instances. 
                Assembly assembly = Assembly.GetExecutingAssembly();
                using (TextReader textReader = new StreamReader(assembly.GetManifestResourceStream("Etk.Excel.Resources.RequestsManagerContextualMenu.xml")))
                {
                    string menuXml = textReader.ReadToEnd();
                    RequestManagementMenu = ContextualMenuFactory.CreateInstances(menuXml).FirstOrDefault();
                }
                // Declare the contextual menus activators. 
                contextualMenuManager.OnContextualMenusRequested += ManageRequestsContexualMenu;
            }
            catch (Exception ex)
            {
                throw new EtkException(string.Format("UDF manager initialization failed:{0}", ex.Message), ex);
            }
        }
        #endregion

        #region public methods
        public void Dispose()
        {
            ETKExcel.ContextualMenuManager.OnContextualMenusRequested -= ManageRequestsContexualMenu;
        }
        #endregion

        #region private methods
        private IEnumerable<IContextualMenu> ManageRequestsContexualMenu(Microsoft.Office.Interop.Excel.Worksheet sheet, Microsoft.Office.Interop.Excel.Range range)
        {
            if (!ETKExcel.ModelDefinitionManager.HasModels)
                return null;

            List<IContextualMenu> menus = new List<IContextualMenu>();
            menus.Add(RequestManagementMenu);
            foreach (IContextualMenu menu in menus)
                (menu as ContextualMenu).SetAction(range); 
            return menus;
        }
        #endregion

        #region Excel Functions
        public static void AddRequest(Microsoft.Office.Interop.Excel.Range caller)
        {
            using (ExcelMainWindow excelWindow = new ExcelMainWindow(caller.Application.Hwnd))
            {
                Microsoft.Office.Interop.Excel.Range firstOutputRange = caller.Offset[0, 1];
                WizardViewModel viewModel = WizardViewModel.CreateInstance(caller, firstOutputRange);
                DynamicRequestManagementWindow window = new DynamicRequestManagementWindow(viewModel);

                WindowInteropHelper windowInteropHelper = new WindowInteropHelper(window);
                windowInteropHelper.Owner = excelWindow.Handle;

                window.DataContext = viewModel;
                window.ShowDialog();
            }
        }

        public static void ModifyRequest()
        { }

        public static void DeleteRequest()
        { }

        public object GDA(Microsoft.Office.Interop.Excel.Range caller, string dataType, object[] parameters)
        {
            if (caller == null)
                return null;

            if (ETKExcel.ExcelApplication.IsInEditMode())
                return "#Edit Mode";

            if(test == null)
            {

                IExcelTemplateView view = ETKExcel.TemplateManager.AddView("Templates Customer", "AllCustomers", caller.Worksheet.Name, caller.Address);
                test = new ExcelRequestDefinition("Test", "Ceci est un test", view as ExcelTemplateView);
                Microsoft.Office.Interop.Excel.Comment comment = caller.Comment;
                if (comment != null)
                    comment.Delete();
                caller.Application.Application.Caller.AddComment(test.Description);
            }

            Microsoft.Office.Interop.Excel.Range firstOutputCell = (caller as Microsoft.Office.Interop.Excel.Range).Offset[++yOffset, ++xOffset];
            (test.View as ExcelTemplateView).FirstOutputCell = firstOutputCell;

            try
            {
                //if (parameters.Length < 2)
                //    return "#N/A";

                ETKExcel.TemplateManager.ClearView(test.View);

                ExcelApplication application = (ETKExcel.TemplateManager as ExcelTemplateManager).ExcelApplication;
                application.PostAsynchronousAction(() =>
                            {
                                try
                                {
                                    application.PostAsynchronousAction(() => (test.View as ExcelTemplateView).FirstOutputCell.Value2 = "#Retrieving Data");
                                    Task task = new Task(() =>
                                    {
                                        //Thread.Sleep(5000);

                                        test.View.SetDataSource(parameters[0]);
                                        application.PostAsynchronousAction(() => test.View.FirstOutputCell.Value2 = string.Empty);
                                        application.PostAsynchronousAction(() => ETKExcel.TemplateManager.Render(test.View));
                                    });
                                    task.Start();
                                }
                                catch (Exception ex)
                                {
                                    string errorMessage = string.Format("#ERR:{0}.{1}", ex.Message, ex.InnerException == null ? string.Empty : ex.InnerException.Message);
                                    application.PostAsynchronousAction(() => (test.View as ExcelTemplateView).FirstOutputCell.Value2 = errorMessage);
                                }
                                                         
                            });
                return test.Name;
            }
            catch (Exception ex)
            {
                return string.Format("#ERR: {0}", ex.Message);
            }                
        }
        #endregion
    }
}
