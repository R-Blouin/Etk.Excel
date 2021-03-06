﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Etk.Excel.Application;
using Etk.Excel.BindingTemplates.Views;
using Etk.Tools.Reflection;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.BindingTemplates
{
    class AccessorParametersManager : IDisposable
    {
        private List<ExcelInterop.Range> rangesToListen = new List<ExcelInterop.Range>();

        public ExcelTemplateView View
        { get; private set; }
        
        public IEnumerable<object> Parameters
        { get; private set; }

        public AccessorParametersManager(ExcelTemplateView view, IEnumerable<object> parameters)
        {
            View = view;
            Parameters = parameters;

            if (Parameters != null && Parameters.Any())
            {
                foreach (object param in Parameters)
                {
                    if (param is ExcelInterop.Range)
                        rangesToListen.Add(param as ExcelInterop.Range);
                    else if (param.GetType().GetInterfaces().Any(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IEnumerable<>)))
                    {
                        Type genericType = param.GetType().GetGenericArguments()[0];
                        MethodInfo convertCollection = typeof(TypeConvertor).GetMethod("ConvertCollection").MakeGenericMethod(genericType);
                        IEnumerable<ExcelInterop.Range> ranges = convertCollection.Invoke(null, new object[] { genericType, param }) as IEnumerable<ExcelInterop.Range>;
                        rangesToListen.AddRange(ranges);
                    }
                }

                if (rangesToListen.Count > 0)
                {
                    ExcelInterop.Worksheet sheet = View.ViewSheet;
                    sheet.Change += OnParametersChanged;
                    ExcelApplication.ReleaseComObject(sheet);
                    sheet = null;
                }
            }
        }

        void OnParametersChanged(ExcelInterop.Range target)
        {
            bool parametersChanged = false;

            foreach (ExcelInterop.Range range in rangesToListen)
            {
                if (ETKExcel.ExcelApplication.Application.Intersect(range, target) != null)
                {
                    parametersChanged = true;
                    break;
                }
            }

            if (parametersChanged)
            {
                List<object> parameters = new List<object>();
                foreach (object param in Parameters)
                {
                    if (param == null)
                        parameters.Add(null);
                    else if (param is ExcelInterop.Range)
                        parameters.Add((param as ExcelInterop.Range).Value);
                    else
                        parameters.Add(param);
                }

                ETKExcel.TemplateManager.ClearView(View);
                ExcelApplication application = (ETKExcel.TemplateManager as ExcelTemplateManager).ExcelApplication;
                application.PostAsynchronousAction(() =>{
                                                            (View as ExcelTemplateView).FirstOutputCell.Value2 = "#Retrieving Data";
                                                            Task task = new Task(() =>
                                                                    {
                                                                        object result = View.TemplateDefinition.DataAccessor.Invoke(parameters);
                                                                        View.SetDataSource(result);
                                                                        ExcelApplication application2 = (ETKExcel.TemplateManager as ExcelTemplateManager).ExcelApplication;
                                                                        application2.PostAsynchronousAction(() =>{
                                                                                                                    (View as ExcelTemplateView).FirstOutputCell.Value2 = string.Empty;
                                                                                                                    ETKExcel.TemplateManager.Render(View as ExcelTemplateView);
                                                                                                                 });
                                                                    });
                                                            task.Start();
                                                        });
            }
        }

        public void Dispose()
        {
            if (rangesToListen != null)
            {
                foreach (ExcelInterop.Range range in rangesToListen)
                    ExcelApplication.ReleaseComObject(range);

                rangesToListen.Clear();
                rangesToListen = null;
            }
            ExcelInterop.Worksheet sheet = View.ViewSheet;
            sheet.Change -= OnParametersChanged;
            ExcelApplication.ReleaseComObject(sheet);
        }
    }
}
