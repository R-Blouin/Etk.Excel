using System.IO;
using System.Linq;
using System.Reflection;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.BindingTemplates.Definitions.Templates;
using Etk.Excel.BindingTemplates.Views;
using Etk.Excel.ContextualMenus;
using Etk.SortAndFilter;

namespace Etk.Excel.BindingTemplates.SortSearchAndFilter
{
    class SortSearchAndFilterMenuManager
    {
        private IContextualMenu sortSearchAndFilersMenu;

        #region .ctors
        public SortSearchAndFilterMenuManager()
        {
            // Create the contextual menu instances. 
            Assembly assembly = Assembly.GetExecutingAssembly();
            using (TextReader textReader = new StreamReader(assembly.GetManifestResourceStream("Etk.Excel.Resources.ViewSortSearchAndFilterContextualMenu.xml")))
            {
                string menuXml = textReader.ReadToEnd();
                sortSearchAndFilersMenu = ContextualMenuFactory.CreateInstances(menuXml).FirstOrDefault();
            }
        }
        #endregion

        #region public methods
        /// <summary>
        /// Manage the contextual menus
        /// </summary>
        public IContextualMenu GetMenus(ExcelTemplateView view, Microsoft.Office.Interop.Excel.Range range, IBindingContextItem contextItem)
        {
            IBindingDefinition bindingDefinition = contextItem.BindingDefinition;
            if (bindingDefinition == null || !bindingDefinition.IsBoundWithData || bindingDefinition.BindingType == null)
                return null;

            if (! ((TemplateDefinition) contextItem.ParentElement.ParentPart.ParentContext.TemplateDefinition).CanSort)
                return null;

            foreach (IContextualPart menuPart in sortSearchAndFilersMenu.Items)
            {
                ContextualMenuItem menuItem = menuPart as ContextualMenuItem;
                menuItem.SetAction(() => menuItem.MethodInfo.Invoke(null, new object[] { view, contextItem }));
            }
            return sortSearchAndFilersMenu;
        }
        #endregion

        #region public methods
        public static void SortAscending(ExcelTemplateView view, IBindingContextItem contextItem)
        {
            ITemplateDefinition templateDefinition = contextItem.ParentElement.ParentPart.TemplateDefinitionPart.Parent;
            ISorterDefinition sortDefinition = SortDefinitionFactory.CreateInstance(templateDefinition, contextItem.BindingDefinition, false, false);

            ExecuteSort(view, sortDefinition);
        }

        public static void SortDescending(ExcelTemplateView view, IBindingContextItem contextItem)
        {
            ITemplateDefinition templateDefinition = contextItem.ParentElement.ParentPart.TemplateDefinitionPart.Parent;
            ISorterDefinition sortDefinition = SortDefinitionFactory.CreateInstance(templateDefinition, contextItem.BindingDefinition, true, false);

            ExecuteSort(view, sortDefinition);
        }
        #endregion


        #region private methods
        private static void ExecuteSort(ExcelTemplateView view, ISorterDefinition sortDefinition)
        {
            view.SorterDefinition = sortDefinition;

            object currentDataSource = view.GetDataSource();
            ETKExcel.TemplateManager.ClearView(view as ExcelTemplateView);
            // We reinject the datasource to force the filtering
            view.CreateBindingContext(currentDataSource);
            // RenderView the view to see the filering application
            ETKExcel.TemplateManager.Render(view as ExcelTemplateView);
        }

        #endregion
    }
}
