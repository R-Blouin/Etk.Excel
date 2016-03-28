using System;
using System.Collections.Generic;
using Etk.BindingTemplates.Views;
using Microsoft.Office.Interop.Excel;

namespace Etk.Excel.BindingTemplates.Views
{
    public interface IExcelTemplateView : ITemplateView
    {
        /// <summary>Contains the range rendered by the View</summary>
        Range RenderedRange { get; }
        /// <summary>Contains the size of the rendered aread</summary>
        RenderedArea RenderedArea { get; }

        /// <summary>Event calls when a data bound this the View is changed</summary>
        event Action<object, object> DataChanged;
        /// <summary>Event calls when the View is about to be rendered (When the rendering is done with 'Render' function, the parameter is set to false. When the rendering is done with 'RenderDataOnly' function, the parameter is set to true</summary>
        event Action<bool> BeforeRendering;
        /// <summary>Event calls after the rendering of the View (When the rendering is done with 'Render' function, the parameter is set to false. When the rendering is done with 'RenderDataOnly' function, the parameter is set to true</summary>
        event Action<bool> AfterRendering;
        /// <summary>Event calls when the sheet that owned the View is activate. The Parameter contains the concerned View</summary>
        event Action<IExcelTemplateView> SheetActivation;

        //RenderingArea RenderedArea { get;  }
        //bool AutoFit {get; set;}

        /// <summary>Set the cell uses to clear the rendered area of the view</summary>
        Range ClearingCell { get; set; }

        void SetAccessorParameters(IEnumerable<object> parameters);
    }
}
