using System;
using System.Collections.Generic;
using System.Reflection;
using Etk.BindingTemplates.Definitions.Templates;

namespace Etk.SortAndFilter
{
    public class SortersAndFilterersFactory
    {
        public static ISortersAndFilters CreateInstance(ITemplateDefinition templateDefinition, IEnumerable<IFilterDefinition> filters, IEnumerable<ISorterDefinition> sorters)
        {
            if (templateDefinition == null)
                return null;

            MethodInfo createLambdaExpression = typeof(SortersAndFilterersFactory).GetMethod("CreateInstance", BindingFlags.NonPublic | BindingFlags.Static).MakeGenericMethod(templateDefinition.BindingType.BindType);
            return createLambdaExpression.Invoke(null, new object[] {templateDefinition, filters, sorters}) as ISortersAndFilters;
        }

        private static ISortersAndFilters CreateInstance<T>(ITemplateDefinition templateDefinition, IEnumerable<IFilterDefinition> filters, IEnumerable<ISorterDefinition> sorters)
        {
            Type type = typeof(SortersAndFilters<>).MakeGenericType(new System.Type[] { typeof(T) });
            return Activator.CreateInstance(type, new object[] {templateDefinition, filters, sorters }) as ISortersAndFilters;
        }
    }
}
