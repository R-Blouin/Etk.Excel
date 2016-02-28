namespace Etk.BindingTemplates.Definitions.SortSearchAndFilter
{
    using Etk.BindingTemplates.Context.SortSearchAndFilter;
    using Etk.BindingTemplates.Definitions.Binding;
    using Etk.BindingTemplates.Views;

    public abstract class BindingSearchDefinition : IDefinitionPart
    {
        #region attribuets and properties
        public BindingPartType PartType
        {
            get { return BindingPartType.SearchDefinition; }
        }

        /// <summary>Watermark</summary>
        public string Watermark
        { get; protected set; }
        #endregion

        #region .ctors
        protected BindingSearchDefinition(string watermark)
        {
            Watermark = watermark;
        }
        #endregion

        #region public methods
        abstract public BindingSearchContextItem CreateContextItem(ITemplateView view);
        #endregion
    }
}
