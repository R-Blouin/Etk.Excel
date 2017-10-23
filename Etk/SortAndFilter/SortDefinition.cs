using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.BindingTemplates.Definitions.Templates;

namespace Etk.SortAndFilter
{
    public class SortDefinition<T, TT> : ISorterDefinition
	{
		#region attributes and properties
        public ITemplateDefinition TemplateDefinition
        { get; }
        
        public IBindingDefinition BindingDefinition
		{ get;  }

		public bool Descending
		{ get;  }

		public bool CaseSensitive
		{ get; }

		public Func<T, TT> SortMethod
		{ get; private set; }

		public Type ResultType => typeof(T);

	    #endregion

		#region .ctors and factory
        public SortDefinition(ITemplateDefinition templateDefinition, IBindingDefinition bindingDefinition, bool descending, bool caseSensitive)
		{
            TemplateDefinition = templateDefinition;
			BindingDefinition = bindingDefinition;
			Descending = descending;
			CaseSensitive = caseSensitive;

			SetExpression();
		}
		#endregion

		#region private methods
		private void SetExpression()
		{
			ParameterExpression param = System.Linq.Expressions.Expression.Parameter(typeof(T), "e");
			Expression<Func<T, TT>> expression = Expression.Lambda<Func<T, TT>>(Expression.Property(param, BindingDefinition.Name), param);
			SortMethod = expression.Compile();
		}
		#endregion

        #region public methods
        public object Sort(object source)
        {
            IEnumerable<T> sourceT = source as IEnumerable<T>;
            IOrderedEnumerable<T> ret;
            if (Descending)
                ret = sourceT.OrderByDescending(SortMethod);
            else
                ret = sourceT.OrderBy(SortMethod);
            return ret;
        }
        #endregion
    }
}
