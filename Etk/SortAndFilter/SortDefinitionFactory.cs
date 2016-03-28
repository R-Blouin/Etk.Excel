using System;
using System.Reflection;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.BindingTemplates.Definitions.Templates;

namespace Etk.SortAndFilter
{
    public static class SortDefinitionFactory
    {
        public static ISorterDefinition CreateInstance(ITemplateDefinition templateDefinition, IBindingDefinition bindingDefinition, bool descending, bool caseSensitive)
        {
            if (templateDefinition == null || bindingDefinition == null)
                return null;

            MethodInfo createLambdaExpression = typeof(SortDefinitionFactory).GetMethod("CreateInstance", BindingFlags.NonPublic | BindingFlags.Static).MakeGenericMethod(templateDefinition.BindingType.BindType, bindingDefinition.BindingType);
            return createLambdaExpression.Invoke(null, new object[] { templateDefinition, bindingDefinition, descending, caseSensitive }) as ISorterDefinition;
        }

        private static ISorterDefinition CreateInstance<T, TT>(ITemplateDefinition templateDefinition, IBindingDefinition bindingDefinition, bool descending, bool caseSensitive)
        {
            Type type = typeof(SortDefinition<,>).MakeGenericType(new System.Type[] { typeof(T), typeof(TT) });
            return Activator.CreateInstance(type, new object[] { templateDefinition, bindingDefinition, descending, caseSensitive }) as ISorterDefinition;
        }
    }
}
