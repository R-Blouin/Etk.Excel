using System;
using System.ComponentModel.Composition;
using Etk.BindingTemplates.Definitions.Decorators;
using Etk.BindingTemplates.Definitions.EventCallBacks;
using Etk.Excel.BindingTemplates.Decorators.XmlDefinitions;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using Etk.BindingTemplates.Definitions.Templates;

namespace Etk.Excel.BindingTemplates.Decorators
{
    /// <summary>Manage the Excel decorators for the current Excel instance</summary>
    [Export(typeof(DecoratorsManager))]
    [Export(typeof(ExcelDecoratorsManager))]
    [PartCreationPolicy(CreationPolicy.Shared)]
    public class ExcelDecoratorsManager : DecoratorsManager
    {
        private readonly ExcelInterop.Application excelApplication;

        private static EventCallbacksManager eventCallbacksManager;
        private static EventCallbacksManager EventCallbacksManager => eventCallbacksManager ??
                                                                      (eventCallbacksManager = CompositionManager.Instance.GetExportedValue<EventCallbacksManager>());

        [ImportingConstructor]
        public ExcelDecoratorsManager([Import] ExcelInterop.Application application)
        {
            excelApplication = application;
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

                if (xmlDecorators.RangeDecorators != null)
                {
                    foreach (XmlExcelRangeDecorator xmlDecorator in xmlDecorators.RangeDecorators)
                    {
                        ExcelRangeDecorator rangeDecorator = ExcelRangeDecorator.CreateInstance(excelApplication, xmlDecorator);
                        RegisterDecorator(rangeDecorator); 
                    }                
                }
            }
            catch (Exception ex)
            {
                string message = xml.Length > 350 ? xml.Substring(0, 350) + "..." : xml;
                throw new EtkException($"Cannot create decorators from xml '{message}':{ex.Message}");
            }
        }

        /// <summary> Register a decorator for future use</summary>
        /// <param name="decorator">The decorator to register</param>
        public void RegisterDecorator(ExcelRangeDecorator decorator)
        {
            if (decorator == null)
                return;
            try
            {
                RegisterDecorator(decorator);
            }
            catch (Exception ex)
            {
                throw new EtkException($"Cannot register decorator '{decorator.Ident ?? string.Empty}':{ex.Message}");
            }
        }

        public override Decorator CreateSimpleDecorator(ITemplateDefinition templateDefinition, string callbackName)
        {
            EventCallback callback = EventCallbacksManager.RetrieveCallback(templateDefinition, callbackName);
            return  new ExcelRangeSimpleDecorator(callbackName, callback);
        }
    }
}
