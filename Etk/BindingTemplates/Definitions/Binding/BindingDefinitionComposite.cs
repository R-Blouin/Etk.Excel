namespace Etk.BindingTemplates.Definitions.Binding
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using Etk.Excel.UI.Log;

    class BindingDefinitionComposite : BindingDefinition
    {
        #region attributes and properties
        static private string pattern = "(?<={)(.*?)(?=})";
        private ILogger log = Logger.Instance;

        private List<IBindingDefinition> nestedDefinitions;
        public ReadOnlyCollection<IBindingDefinition> NestedDefinitions
        { get { return new ReadOnlyCollection<IBindingDefinition>(nestedDefinitions); } }

        public string BindingFormat
        { get; private set; }

        private List<IBindingDefinition> canBeNotifiedNestedDefinitions;
        #endregion

        #region .ctors
        private BindingDefinitionComposite(BindingDefinitionDescription bindingDefinitionDescription)
            : base(bindingDefinitionDescription)
        {}
        #endregion

        #region override 'BindingDefinition' methods
        override public object ResolveBinding(object dataSource)
        {
            try
            {
                if (dataSource != null)
                {
                    object[] results = new object[nestedDefinitions.Count];
                    for (int cpt = 0; cpt < nestedDefinitions.Count; cpt++)
                    {
                        object obj = nestedDefinitions[cpt].ResolveBinding(dataSource);
                        results[cpt] = obj ?? string.Empty;
                    }

                    return string.Format(BindingFormat, results);
                }
                return null;
            }
            catch (Exception ex)
            {
                string message = string.Format("Can't Resolve the 'Binding' for the BindingExpression '{0}'. {1}", BindingExpression, ex.Message);
                throw new BindingTemplateException(message, ex);
            }
        }

        /// <summary>
        /// Update the datasource with the binding value.
        /// If the BindingDefinition to update is readonly, then return the currently loaded value
        /// Else return the value passed as a parameter. 
        /// </summary>
        /// <param name="contextItem">the data IBindingContextItem to update.</param>
        /// <param name="data">the data to update the datasource.</param>
        /// <returns></returns>
        override public object UpdateDataSource(object dataSource, object data)
        {
            try
            {
                if (dataSource == null)
                    return null;

                if (! IsReadOnly)
                    nestedDefinitions[0].UpdateDataSource(dataSource, data);
                return ResolveBinding(dataSource);
            }
            catch (Exception ex)
            {
                log.LogFormat(LogType.Warn, "Cannot Resolve 'UpdateDataSource' for the BindingExpression '{0}'. {1}", BindingExpression, ex.Message);
                return ResolveBinding(dataSource);
            }
        }

        override public IEnumerable<INotifyPropertyChanged> GetObjectsToNotify(object dataSource)
        {
            List<INotifyPropertyChanged> notifyPropertyChangedList = new List<INotifyPropertyChanged>();
            foreach (IBindingDefinition definition in canBeNotifiedNestedDefinitions)
            {
                IEnumerable<INotifyPropertyChanged> results = definition.GetObjectsToNotify(dataSource);
                if (results != null)
                    notifyPropertyChangedList.AddRange(results);
            }
            return notifyPropertyChangedList;
        }

        override public bool MustNotify(object dataSource, object source, PropertyChangedEventArgs args)
        {
            foreach (IBindingDefinition definition in canBeNotifiedNestedDefinitions)
            {
                if (definition.MustNotify(dataSource, source, args))
                    return true;
            }
            return false;
        }
        #endregion

        #region static public methods
        static public BindingDefinitionComposite CreateInstance(Type type, BindingDefinitionDescription definitionDescription)
        {
            try
            {
                if (string.IsNullOrEmpty(definitionDescription.Name))
                {
                    definitionDescription.Name = definitionDescription.BindingExpression.Replace('.', '_');
                    MatchCollection ret = BindingDefinition.ValidCharExtract.Matches(definitionDescription.Name);
                    StringBuilder sb = new StringBuilder();
                    foreach (Match m in ret)
                        sb.Append(m.Value);
                    definitionDescription.Name = sb.ToString();
                }

                definitionDescription.BindingExpression = definitionDescription.BindingExpression.Substring(1);
                definitionDescription.BindingExpression = definitionDescription.BindingExpression.Substring(0, definitionDescription.BindingExpression.Length - 1);

                BindingDefinitionComposite definition = null;
                string bindingFormat = definitionDescription.BindingExpression; 
                List<string> results = new List<string>();
                MatchCollection matches = Regex.Matches(bindingFormat, pattern);

                int cpt = -1;
                foreach (Match match in matches)
                {
                    string[] elements = match.Value.Split(':');
                    if (string.IsNullOrEmpty(elements[0]))
                        bindingFormat = bindingFormat.Replace(string.Format("{{{0}}}", match.Value), string.Empty);
                    else
                    {
                        int pos = results.FindIndex(s => s.Equals(elements[0]));
                        if (pos == -1)
                        {
                            results.Add(elements[0]);
                            pos = ++cpt;
                        }
                        else
                            pos = cpt;
                        string format = string.Format("{{{0}}}", match.Value);
                        bindingFormat = bindingFormat.Replace(format, string.Format("{{{0}}}", pos));
                    }
                }
                if (results.Count > 0)
                {
                    List<BindingDefinitionDescription> definitionDescriptions = new List<BindingDefinitionDescription>();
                    if (results.Count == 1)
                        definitionDescriptions.Add(new BindingDefinitionDescription() { BindingExpression = results[0], IsReadOnly = definitionDescription.IsReadOnly});
                    else
                        definitionDescriptions.AddRange(results.Select(s => new BindingDefinitionDescription() { BindingExpression = s, IsReadOnly = false}));

                    List<IBindingDefinition> nestedDefinitions = BindingDefinitionFactory.CreateInstances(type, definitionDescriptions);

                    if(nestedDefinitions.FirstOrDefault(nd => nd.IsACollection) != null)
                        throw new BindingTemplateException("The nested 'BindingDefinition' of a 'Composite BindingDefinition' cannot be a collection");

                    // If more than one nested definition, then force the Binding definition to be ReadOnly
                    definitionDescription.IsReadOnly = nestedDefinitions.Count > 1 ? true : nestedDefinitions[0].IsReadOnly;
                    definition = new BindingDefinitionComposite(definitionDescription){nestedDefinitions = nestedDefinitions,
                                                                                       BindingFormat = bindingFormat,
                                                                                       BindingType = typeof(string)};

                    definition.canBeNotifiedNestedDefinitions = definition.nestedDefinitions.Where(d => d.CanNotify).ToList();
                    definition.CanNotify = definition.canBeNotifiedNestedDefinitions.Count > 0;
                }
                return definition;
            }
            catch (Exception ex)
            {
                string message = string.Format("Cannot create the 'Composite BindingDefinition' '{0}'. {1}", definitionDescription.BindingExpression, ex.Message);
                throw new BindingTemplateException(message, ex);
            }
        }
        #endregion
    }
}
