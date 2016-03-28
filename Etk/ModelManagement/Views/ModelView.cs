using System;
using System.Collections.Generic;
using System.Linq;
using Etk.ModelManagement.Definitions.XmlDefinition;

namespace Etk.ModelManagement.Views
{
    public class ModelView : IModelView
    {
        #region constantes
        const char SUB_PROP_START = '{';
        const char SUB_PROP_END = '}';
        const char PROPERTIES_SEP = ';';
        #endregion

        #region properties and attributes
        private string propertiesList;
        private string accessorName;

        ///// <summary> Implements <see cref="IModelView.Ident"/> </summary> 
        //public Guid Ident
        //{ get; private set; }

        /// <summary> Implements <see cref="IModelView.Name"/> </summary> 
        public string Name
        { get; private set; }

        /// <summary> Implements <see cref="IModelView.Description"/> </summary> 
        public string Description
        { get; private set; }

        private ModelType parent;
        /// <summary> Implements <see cref="IModelView.ParentElement"/> </summary> 
        public IModelType Parent
        { get { return parent; } }

        /// <summary> Implements <see cref="IModelView.IsDefault"/> </summary> 
        public bool IsDefault
        { get; private set; }

        /// <summary> Implements <see cref="IModelView.Accessor"/> </summary> 
        public IModelAccessor Accessor
        { get; private set; }

        /// <summary> Implements <see cref="IModelView.Parts"/> </summary> 
        public IEnumerable<IModelViewPart> Parts
        { get { return ChangeableParts; } }

        public List<IModelViewPart> ChangeableParts
        { get; private set; }
        #endregion

        #region .ctors and factories
        public ModelView(IModelType parent, string name, string description)
        {
            this.parent = parent as ModelType;
            Name = name;
            Description = description;
            ChangeableParts = new List<IModelViewPart>();
        }

        public ModelView(IModelType parent, XmlModelView xmlView)
        {
            Name = string.IsNullOrEmpty(xmlView.Name) ? parent.Name : xmlView.Name;
            Description = string.IsNullOrEmpty(xmlView.Description) ? parent.Description : xmlView.Description;
            propertiesList = xmlView.Properties;
            accessorName = xmlView.Accessor;

            IsDefault = xmlView.IsDefault;
            this.parent = parent as ModelType;

            ChangeableParts = new List<IModelViewPart>();
        }

        public ModelView(IModelType parent, string propertiesList)
        {
            this.parent = parent as ModelType;
            this.propertiesList = propertiesList;
            ChangeableParts = new List<IModelViewPart>();
            ResolveDependencies();
        }
        #endregion

        #region public methods
        /// <summary> Resoolve the view : analyze its contain and create the parts</summary>
        public void ResolveDependencies()
        {
            // Get properties definition
            string toAnalyze = propertiesList.Trim();
            while (!string.IsNullOrEmpty(toAnalyze))
            {
                string beforeSubProperties = null;
                string afterSubProperties = null;
                string subProperties = null;
                int startSubProperties, endSubproperties;
                FindSubProperties(toAnalyze, out startSubProperties, out endSubproperties);
                if (startSubProperties == -1)
                    beforeSubProperties = toAnalyze;
                else
                {
                    if (endSubproperties == -1)
                        throw new EtkException(string.Format("'Cannot create View. '{0}' doesn't contain '{1}'", toAnalyze, SUB_PROP_END));

                    subProperties = toAnalyze.Substring(startSubProperties + 1, endSubproperties - startSubProperties - 1);
                    beforeSubProperties = toAnalyze.Substring(0, startSubProperties);
                    afterSubProperties = toAnalyze.Substring(endSubproperties + 1);
                }

                AnalyzeProperties(this, beforeSubProperties, subProperties);
                toAnalyze = afterSubProperties;
            }
        }
        #endregion

        #region private methods
        private void AnalyzeProperties(ModelView modelView, string toAnalyze, string subProperties)
        {
            string[] properties = toAnalyze.Split(PROPERTIES_SEP).Where(s => ! string.IsNullOrWhiteSpace(s)).ToArray();
            int lastIndex = properties.Count() - 1;
            for (int i = 0; i < properties.Count(); i++)
            {
                string part = properties[i].Trim();
                if (!string.IsNullOrEmpty(part))
                {
                    if (i == lastIndex && !string.IsNullOrWhiteSpace(subProperties))
                    {
                        ModelType nestedModelType = modelView.Parent.Parent.GetModelType(part) as ModelType;
                        if (nestedModelType == null)
                            throw new EtkException(string.Format("Cannot find the model type '{0}'", part));

                        ModelView nestedModelView = new ModelView(nestedModelType, subProperties);
                        ChangeableParts.Add(nestedModelView);
                    }
                    else
                    {
                        try
                        {
                            ModelViewProperty property = ModelViewProperty.CreateInstance(modelView.Parent, part);
                            if (property != null)
                                modelView.ChangeableParts.Add(property);
                        }
                        catch (Exception ex)
                        {
                            throw new EtkException(string.Format("Cannot find property '{0}':{1}", part, ex.Message));
                        }
                    }
                }
            }
        }

        private static void FindSubProperties(string toAnalyze, out int start, out int end)
        {
            start = toAnalyze.IndexOf(SUB_PROP_START);
            end = toAnalyze.IndexOf(SUB_PROP_END);
            if (start != -1 && end != -1)
            {
                int findNextStartBracket = toAnalyze.IndexOf(SUB_PROP_START, start + 1);
                while (findNextStartBracket != -1 && findNextStartBracket < end)
                {
                    findNextStartBracket = toAnalyze.IndexOf(SUB_PROP_START, findNextStartBracket + 1);
                    end = toAnalyze.IndexOf(SUB_PROP_END, end + 1);
                }
            }
        }
        #endregion
    }
}
