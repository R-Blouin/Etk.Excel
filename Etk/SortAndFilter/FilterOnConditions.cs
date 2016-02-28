namespace Etk.SortAndFilter
{
    using Etk.BindingTemplates.Definitions.Binding;
    using Etk.BindingTemplates.Definitions.Templates;

    public class FilterOnConditions : IFilterDefinition
    {
        public ITemplateDefinition TemplateDefinition
        { get; private set; }

        public IBindingDefinition DefinitionToFilter
        { get; private set; }

        public string FilterExpression
        { get; protected set; }

        public string Condition
        { get; private set; }

        public bool CaseSensitive
        { get; private set; }

        public FilterDefinitionRelation Relation
        { get; private set; }

        #region .ctors
        public FilterOnConditions(ITemplateDefinition templateDefinition, IBindingDefinition definition, string condition, FilterDefinitionRelation relation, bool caseSensitive)
        {
            TemplateDefinition = templateDefinition;
            DefinitionToFilter = definition;
            Condition = condition;
            Relation = relation;
            CaseSensitive = caseSensitive;
        }
        #endregion
    }
}