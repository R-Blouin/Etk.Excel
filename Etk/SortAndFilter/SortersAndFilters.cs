using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using Etk.BindingTemplates.Definitions.Templates;

namespace Etk.SortAndFilter
{
    public class SortersAndFilters<T> : ISortersAndFilters
    {
        #region attributes and properties
        private Func<T, bool> filterMethod;

        public ITemplateDefinition TemplateDefinition
        { get; private set; }

        public List<IFilterDefinition> Filters
        { get; private set; }

        public List<ISorterDefinition> Sorters
        { get;  }

        public Type ResultType => typeof(T);

        public bool IsActive => Filters != null && Filters.Any() || Sorters != null && Sorters.Any();

        #endregion

        #region .ctors
        public SortersAndFilters(ITemplateDefinition templateDefinition, IEnumerable<IFilterDefinition> filters, IEnumerable<ISorterDefinition> sorters)
        {
            TemplateDefinition = templateDefinition;
            Filters = filters?.ToList();
            Sorters = sorters?.ToList();
            SetFilterMethod();
        }
        #endregion

        #region public methods
        public object Execute(IEnumerable<object> param)
        {
            IEnumerable<T> ret = param.Cast<T>().ToList();
            if (ret != null)
            {
                if (filterMethod != null)
                    ret = ret.Where(e => filterMethod(e));

                if (Sorters != null)
                {
                    foreach (ISorterDefinition sorter in Sorters)
                        ret = sorter.Sort(ret) as IEnumerable<T>;
                }
            }
            return ret;
        }

        public void Add(IFilterDefinition filterElement)
        {
            if (Filters == null)
                Filters = new List<IFilterDefinition>();
            if (!Filters.Contains(filterElement))
            {
                Filters.Add(filterElement);
                SetFilterMethod();
            }
        }

        public void Remove(IFilterDefinition filterElement)
        {
            if (Filters != null)
            {
                Filters.Remove(filterElement);
                SetFilterMethod();
            }
        }
        #endregion

        #region private methods
        private void SetFilterMethod()
        {
            if (Filters != null)
            {
                string[] filters = Filters.Where(f => !string.IsNullOrEmpty(f.FilterExpression))
                                          .Select(f => $"({f.FilterExpression})")
                                          .ToArray();
                if (filters.Any())
                {
                    string expressionString = string.Join(" AND ", filters);
                    Expression<Func<T, bool>> expression = System.Linq.Dynamic.DynamicExpression.ParseLambda<T, bool>(expressionString, null);
                    filterMethod = expression.Compile();
                }
            }
        }
        #endregion
    }
}