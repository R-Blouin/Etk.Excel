using System;
using System.Collections.Generic;
using Etk.BindingTemplates.Views;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace Etk.Excel.BindingTemplates.Views
{
    public interface IExcelTemplateView : ITemplateView
    {
        /// <summary>Get the the sheet that contains the view</summary>
        ExcelInterop.Worksheet ViewSheet { get; }
        /// <summary>Contains the range rendered by the View</summary>
        ExcelInterop.Range RenderedRange { get; }
        /// <summary>Contains the size of the rendered aread</summary>
        RenderedArea RenderedArea { get; }
        /// <summary>Set/get if view must autofit after a complete rendering (default is true)</summary>
        AutoFitMode AutoFit { get; set; }
        /// <summary>Set the cell uses to clear the rendered area of the view</summary>
        ExcelInterop.Range ClearingCell { get; set; }

        /// <summary> Rerender (description, style and data) the current view.
        /// Internally call <see cref="Etk.Excel.BindingTemplates.IExcelTemplateManager.RenderView"/></summary>
        void Render();
        /// <summary> Rerender only the data of the View given as parameter (This implies that no change was made in the structure of the data).
        /// Internally call <see cref="Etk.Excel.BindingTemplates.IExcelTemplateManager.RenderDataOnly"/></summary>
        void RenderDataOnly();
        /// <summary> ClearView the previously rendering View
        /// Internally call <see cref="Etk.Excel.BindingTemplates.IExcelTemplateManager.ClearView"/></summary>
        void ClearView();
        /// <summary>Autofit the view (without taking into account the value of <see cref="AutoFit"/>)</summary>
        void ExecuteAutoFit();

        /// <summary>Event calls when a data bound this the View is changed</summary>
        event Action DataChanged;
        /// <summary>Event calls when the View is about to be rendered (When the rendering is done with 'RenderView' function, the parameter is set to false. When the rendering is done with 'RenderViewDataOnly' function, the parameter is set to true</summary>
        event Action<bool> BeforeRendering;
        /// <summary>Event calls after the rendering of the View (When the rendering is done with 'RenderView' function, the parameter is set to false. When the rendering is done with 'RenderViewDataOnly' function, the parameter is set to true</summary>
        event Action<bool> AfterRendering;
        /// <summary>Event calls when the sheet that contains the view is activated. The Parameter is the concerned View</summary>
        event Action ViewSheetIsActivated;
        /// <summary>Event calls when the sheet that contains the view is desactivated. The Parameter is the concerned View</summary>
        event Action ViewSheetIsDeactivated;

        void SetAccessorParameters(IEnumerable<object> parameters);
    }
}
