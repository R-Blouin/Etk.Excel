using Etk.Excel.BindingTemplates.Views;
using System;
using System.Runtime.InteropServices;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.Addin
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class EtkView
    {
        #region attributs and properties
        [ComVisible(false)]
        public IExcelTemplateView ExcelView { get; }

        public object RenderedRange => ExcelView.RenderedRange;

        public object DataSource
        {
            get { return ExcelView.GetDataSource(); }
            set { ExcelView.SetDataSource(value); }
        }

        public string SearchValue
        { get; set; }

        public object ClearingCell
        {
            get { return ExcelView.ClearingCell; }
            set { ExcelView.ClearingCell = value as ExcelInterop.Range; }
        }

        private Action dataChangedAction;
        private string dataChanged;
        public string OnDataChanged
        {
            get { return dataChanged; }
            set
            {
                dataChanged = value;
                if (dataChangedAction != null)
                    ExcelView.DataChanged -= dataChangedAction;

                if (!string.IsNullOrEmpty(value))
                {
                    dataChangedAction = () => ETKExcel.ExcelApplication.ExecuteVbaMAcro(dataChanged, null);
                    ExcelView.DataChanged += dataChangedAction;
                }
            }
        }

        private Action<bool> beforeRenderingAction;
        private string beforeRendering;
        public string OnBeforeRendering
        {
            get { return beforeRendering; }
            set
            {
                beforeRendering = value;
                if (beforeRenderingAction != null)
                    ExcelView.BeforeRendering -= beforeRenderingAction;

                if (! string.IsNullOrEmpty(value))
                {
                    beforeRenderingAction = (p) => ETKExcel.ExcelApplication.ExecuteVbaMAcro(beforeRendering, new object[] { p });
                    ExcelView.BeforeRendering += beforeRenderingAction;
                }
            }
        }

        private Action<bool> afterRenderingAction;
        private string afterRendering;
        public string OnAfterRendering
        {
            get { return afterRendering; }
            set
            {
                afterRendering = value;
                if (afterRenderingAction != null)
                    ExcelView.AfterRendering -= afterRenderingAction;

                if (!string.IsNullOrEmpty(value))
                {
                    afterRenderingAction = (p) => ETKExcel.ExcelApplication.ExecuteVbaMAcro(afterRendering, new object[] { p });
                    ExcelView.AfterRendering += afterRenderingAction;
                }
            }
        }

        private Action<IExcelTemplateView> viewSheetIsActivatedAction;
        private string viewSheetIsActivated;
        public string OnViewSheetIsActivated
        {
            get { return viewSheetIsActivated; }
            set
            {
                viewSheetIsActivated = value;
                if (viewSheetIsActivatedAction != null)
                    ExcelView.ViewSheetIsActivated -= viewSheetIsActivatedAction;

                if (!string.IsNullOrEmpty(value))
                {
                    viewSheetIsActivatedAction = ExcelView => ETKExcel.ExcelApplication.ExecuteVbaMAcro(viewSheetIsActivated, new object[] { this });
                    ExcelView.ViewSheetIsActivated += viewSheetIsActivatedAction;
                }
            }
        }

        private System.Action<IExcelTemplateView> viewSheetIsDeactivatedAction;
        private string viewSheetIsDeActivated;
        public string OnViewSheetIsDeActivated
        {
            get { return viewSheetIsDeActivated; }
            set
            {
                viewSheetIsDeActivated = value;
                if (viewSheetIsDeactivatedAction != null)
                    ExcelView.ViewSheetIsActivated -= viewSheetIsDeactivatedAction;

                if (!string.IsNullOrEmpty(value))
                {
                    viewSheetIsDeactivatedAction = ExcelView => ETKExcel.ExcelApplication.ExecuteVbaMAcro(viewSheetIsDeActivated, new object[] { this });
                    ExcelView.ViewSheetIsDeactivated += viewSheetIsDeactivatedAction;
                }
            }
        }
        #endregion

        #region .ctors and factories
        private EtkView(IExcelTemplateView excelView)
        {
            ExcelView = excelView;
        }

        public static EtkView CreateInstance(IExcelTemplateView view)
        {
            if (view == null)
                return null;
            return new EtkView(view);
        }
        #endregion

        #region public methods
        public void ExecuteSearch()
        {
            ExcelView.ExecuteSearch();
        }

        public void ExecuteAutoFit()
        {
            ExcelView.ExecuteAutoFit();
        }
        #endregion
    }
}
