using System;
using System.Reflection;
using Etk.BindingTemplates.Context;
using Etk.BindingTemplates.Definitions.Binding;
using Etk.Tools.Reflection;

namespace Etk.BindingTemplates.Definitions.Templates
{
    public class LinkedTemplateDefinition : ILinkedTemplateDefinition 
    {
        public BindingPartType PartType => BindingPartType.LinkedTemplateDefinition;

        public ITemplateDefinition Parent
        { get; private set; }

        public ITemplateDefinition TemplateDefinition
        {get; private set;}

        public IBindingDefinition BindingDefinition
        { get; private set; }

        public LinkedTemplatePositioning Positioning
        { get; private set; }

        /// <summary> Method to invoke to determinate the min nomber of occurences the link templates must occupied</summary>
        public MethodInfo MinOccurencesMethod
        { get; private set; }

        /// <summary> Method to invoke to force the number of occurences the link templates must have</summary>
        public MethodInfo NumberOfOccurencesMethod
        { get; private set; }

        #region .ctors
        public LinkedTemplateDefinition(ITemplateDefinition parent, ITemplateDefinition templateDefinition, TemplateLink linkDefinition)
        {
            try
            {
                Parent = parent;
                TemplateDefinition = templateDefinition;
                Positioning = linkDefinition.Positioning;
                if (!string.IsNullOrEmpty(linkDefinition.With))
                {
                    Type type = parent.MainBindingDefinition != null ? parent.MainBindingDefinition.BindingType : null;
                    BindingDefinitionDescription definitionDescription = new BindingDefinitionDescription()
                    {
                        BindingExpression = linkDefinition.With,
                        Description = linkDefinition.Description,
                        Name = linkDefinition.Name,
                        IsReadOnly = true
                    };
                    BindingDefinition = BindingDefinitionFactory.CreateInstance(type, definitionDescription);
                }

                if (!string.IsNullOrEmpty(linkDefinition.MinOccurencesMethod))
                {
                    try
                    {
                        if (templateDefinition.Header != null && ((TemplateDefinitionPart)templateDefinition.Header).HasLinkedTemplates
                           || templateDefinition.Body != null && ((TemplateDefinitionPart)templateDefinition.Body).HasLinkedTemplates
                           || templateDefinition.Footer != null && ((TemplateDefinitionPart)templateDefinition.Footer).HasLinkedTemplates)
                            throw new Exception("'MinOccurences' is not supported with templates linked with other templates");

                        Type type = TemplateDefinition.MainBindingDefinition?.BindingType;
                        MinOccurencesMethod = TypeHelpers.GetMethod(type, linkDefinition.MinOccurencesMethod);
                        if (MinOccurencesMethod.GetParameters().Length > 2)
                            throw new Exception("The 'MinOccurences' resolver method signature must be 'int <MethodName>([instance of element of the collection that owned the link declaration])'");
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Cannot retrieve the 'MinOccurences' resolver method:{ex.Message}");
                    }
                }
                if (!string.IsNullOrEmpty(linkDefinition.NumberOfOccurencesMethod))
                {
                    try
                    {
                        if (templateDefinition.Header != null && ((TemplateDefinitionPart)templateDefinition.Header).HasLinkedTemplates
                            || templateDefinition.Body != null && ((TemplateDefinitionPart)templateDefinition.Body).HasLinkedTemplates
                            || templateDefinition.Footer != null && ((TemplateDefinitionPart)templateDefinition.Footer).HasLinkedTemplates)
                            throw new Exception("'NumberOfOccurences' is not supported with templates linked with other templates");

                        Type type = TemplateDefinition.MainBindingDefinition?.BindingType;
                        NumberOfOccurencesMethod = TypeHelpers.GetMethod(type, linkDefinition.NumberOfOccurencesMethod);
                        if (NumberOfOccurencesMethod.GetParameters().Length > 2)
                            throw new Exception("The 'NumberOfOccurences' resolver method signature must be 'int <MethodName>([instance of element of the collection that owned the link declaration])'");
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Cannot retrieve the 'NumberOfOccurences' resolver method:{ex.Message}");
                    }
                }

            }
            catch (Exception ex)
            {
                string message = $"Cannot resolve linked template definition 'To={linkDefinition.To}, With='{linkDefinition.With}'";
                throw new Exception(message, ex);
            }
        }
        #endregion

        #region public methods
        public object ResolveBinding(object dataSource)
        {
            if (dataSource == null)
                return null;

            if (BindingDefinition == null)
                return dataSource;

            if (BindingDefinition.IsOptional)
                BindingDefinition = (BindingDefinition as BindingDefinitionOptional).CreateRealBindingDefinition(dataSource.GetType());
       
            return BindingDefinition.ResolveBinding(dataSource);
        }

        #region public methods
        public static int ResolveNumberOfOccurences(MethodInfo method, IBindingContextElement parentElement)
        {
            object invokeTarget = method.IsStatic ? null : parentElement.DataSource;
            int nbrParameters = method.GetParameters().Length;

            object[] parameters = null;
            switch (nbrParameters)
            {
                //case 2:
                //    parameters = new object[] { currentElement.DataSource, parentElement.DataSource };
                //break;
                case 1:
                    parameters = new object[] { parentElement.DataSource };
                    break;
            }

            return (int)method.Invoke(invokeTarget, parameters);
        }

        #endregion
        #endregion
    }
}
