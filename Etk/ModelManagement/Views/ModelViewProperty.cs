namespace Etk.ModelManagement.Views
{
    class ModelViewProperty : IModelViewProperty
    {
        public IModelProperty ModelProperty
        { get; private set; }

        public string Name
        { get { return ModelProperty.Name; } }

        public string Header
        { get; private set; }

        public IModelType Parent
        { get; private set; }

        //public bool IsComposed
        //{ get; private set; }

        #region .ctors and factories
        private ModelViewProperty(IModelType parent, IModelProperty modelProperty)
        {
            Parent = parent;
            ModelProperty = modelProperty;
        }

        public static ModelViewProperty CreateInstance(IModelType parent, string name)
        {
            if (parent == null)
                return null;

            if (string.IsNullOrWhiteSpace(name))
                return null;

            name = name.Trim();
            IModelProperty modelProperty = parent.GetProperty(name);
            if (modelProperty == null)
                throw new EtkException(string.Format("Cannot find property '{0}' for model type {1}", name, parent.Name));

            return new ModelViewProperty(parent, modelProperty);
        }
        #endregion

        #region public methods
        #endregion

        #region private methods
        #endregion
    }
}
