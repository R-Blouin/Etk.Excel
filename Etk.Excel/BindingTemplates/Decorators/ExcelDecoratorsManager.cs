namespace Etk.Excel.BindingTemplates.Decorators
{
    using System;
    using System.ComponentModel.Composition;
    using Etk.BindingTemplates.Definitions.Decorators;
    using Etk.Excel.BindingTemplates.Decorators.XmlDefinitions;
    using Excel = Microsoft.Office.Interop.Excel;

    /// <summary>Manage the Excel decorators for the current Excel instance</summary>
    [Export]
    [PartCreationPolicy(CreationPolicy.Shared)]
    public class ExcelDecoratorsManager 
    {
        private Excel.Application excelApplication;
        private DecoratorsManager decoratorsManager;

        [ImportingConstructor]
        public ExcelDecoratorsManager([Import] Excel.Application application, [Import] DecoratorsManager decoratorsManager)
        {
            this.excelApplication = application;
            this.decoratorsManager = decoratorsManager;
        }

        /// <summary>Register decorators from xml definitions</summary>
        /// <param name="xml">The xml that contains the decorators definitions </param>
        public void RegisterDecoratorsFromXml(string xml)
        {
            try
            {
                XmlExcelDecorators xmlDecorators = XmlExcelDecorators.CreateInstance(xml);
                if (xmlDecorators == null)
                    return;

                //if(xmlDecorators.Decorators != null)
                //{
                //    foreach (XmlExcelDecorator xmlDecorator in xmlDecorators.Decorators)
                //    {
                        
                //    }                
                //}

                if(xmlDecorators.RangeDecorators != null)
                {
                    foreach (XmlExcelRangeDecorator xmlDecorator in xmlDecorators.RangeDecorators)
                    {
                        ExcelRangeDecorator rangeDecorator = ExcelRangeDecorator.CreateInstance(excelApplication, xmlDecorator);
                        decoratorsManager.RegisterDecorator(rangeDecorator); 
                    }                
                }
            }
            catch (Exception ex)
            {
                string message = xml.Length > 350 ? xml.Substring(0, 350) + "..." : xml;
                throw new EtkException(string.Format("Cannot create decorators from xml '{0}':{1}", message, ex.Message), ex);
            }
        }

        /// <summary> Register a decorator for future use</summary>
        /// <param name="decorator">The decorator to register</param>
        public void RegisterDecorator(ExcelRangeDecorator decorator)
        {
            try
            {
                if (decorator != null)
                {
                    decoratorsManager.RegisterDecorator(decorator);
                }
            }
            catch (Exception ex)
            {
                throw new EtkException(string.Format("Cannot register decorator '{0}':{1}", decorator.Ident ?? string.Empty, ex.Message), ex);
            }
        }
    }
}
