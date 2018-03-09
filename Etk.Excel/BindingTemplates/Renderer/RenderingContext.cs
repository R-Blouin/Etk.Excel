using System.Collections.Generic;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Templates;

namespace Etk.Excel.BindingTemplates.Renderer
{
    class RenderingContext
    {
        public int InitPos { get; private set; }
        public List<IBindingContextItem> ContextItems { get; private set; }
        public IBindingContextElement ContextElement { get; private set; }

        public LinkedTemplateDefinition LinkedTemplateDefinition { get; set; }

        public int CurrentHeight { get; set; }
        public int CurrentWidth { get; set; }
        public int PosCurrentLink { get; set; }
        public int PosPreviousLink { get; set; }
        public int LinkedViewRenderedOffset { get; set; }
        public int RefPos { get; set; }
        public bool RowColAdded { get; set; }

        public RenderingContext(IBindingContextElement contextElement, int pos)
        {
            ContextElement = contextElement;
            InitPos = pos;
            ContextItems = new List<IBindingContextItem>();
        }
    }
}
