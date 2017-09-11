namespace Etk.ModelManagement.Views
{
    class ModelViewProperty : IModelViewProperty
    {
        public IModelProperty ModelProperty
        { get; }

        public string Name => ModelProperty.Name;

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
                throw new EtkException($"Cannot find property '{name}' for model type {parent.Name}");

            return new ModelViewProperty(parent, modelProperty);
        }
        #endregion

        #region public methods
        #endregion

        #region private methods
        #endregion
    }
}
