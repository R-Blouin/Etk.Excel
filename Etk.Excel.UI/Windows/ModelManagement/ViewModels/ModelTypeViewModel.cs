using Etk.Excel.MvvmBase;
using Etk.ModelManagement;

namespace Etk.Excel.UI.Windows.ModelManagement.ViewModels
{
    public class ModelTypeViewModel : ViewModelBase
    {
        #region attributes and properties
        public IModelType ModelType
        { get; private set; }

        //private List<IModelPropertyViewModel> properties;
        //public IEnumerable<IModelPropertyViewModel> Properties
        //{
        //    get { return properties; }
        //}

        //private IEnumerable<IModelViewProperty> selectedProperties;
        //public IEnumerable<IModelViewProperty> SelectedProperties
        //{
        //    get { return selectedProperties; }
        //    set
        //    {
        //        selectedProperties = value;
        //        OnPropertyChanged("SelectedProperties");
        //    }
        //}

        //private List<AccessorParameter> selectedAccessorParameters;
        //public List<AccessorParameter> SelectedAccessorParameters
        //{
        //    get { return selectedAccessorParameters; }
        //}

        //public ListCollectionView SelectedParametersCollectionView
        //{
        //    get
        //    {
        //        ListCollectionView ret = new ListCollectionView(SelectedAccessorParameters);
        //        ret.GroupDescriptions.Add(new PropertyGroupDescription("ParameterInfos.Name"));
        //        return ret;
        //    }
        //}
        #endregion

        #region .ctors and factory
        //public ModelTypeViewModel(Dictionary<string, IModelType> modelTypeByName, IModelType modelType)
        public ModelTypeViewModel(IModelType modelType)
        {
            ModelType = modelType;

            //properties = new List<IModelPropertyViewModel>();
            //foreach (IModelProperty property in modelType.GetProperties())
            //{
            //    IModelPropertyViewModel viewModel = null;
            //    if (property is ModelBoundProperty)
            //        viewModel = new ModelPropertyViewModel(property as ModelBoundProperty, null);
            //    else
            //        viewModel = new ModelLinkPropertyViewModel(property as ModelLinkProperty);

            //    if (viewModel != null)
            //        properties.Add(viewModel);
            //}

            //selectedAccessorParameters = PrepareSelectedAccessorParameters(Accessors.FirstOrDefault()).ToList();
        }
        #endregion

        #region private methods
        //private IEnumerable<AccessorParameter> PrepareSelectedAccessorParameters(IModelAccessor accessor)
        //{
        //    List<AccessorParameter> ret = new List<AccessorParameter>();
        //    if (accessor != null && accessor.DataAccessor.ParametersInfo != null)
        //    {
        //        Dictionary<string, List<ParameterInfo>> parametersInfoByName = new Dictionary<string, List<ParameterInfo>>();
        //        foreach (ParameterInfo parameterInfo in accessor.DataAccessor.ParametersInfo)
        //        {
        //            List<ParameterInfo> infos;
        //            if (!parametersInfoByName.TryGetValue(parameterInfo.Name, out infos))
        //            {
        //                infos = new List<ParameterInfo>();
        //                parametersInfoByName[parameterInfo.Name] = infos;
        //            }
        //            infos.Add(parameterInfo);
        //        }

        //        foreach (KeyValuePair<string, List<ParameterInfo>> kvp in parametersInfoByName)
        //        {
        //            AccessorParameter accessorParameter = new AccessorParameter(kvp.Key, kvp.Value.FirstOrDefault());
        //            ret.Add(accessorParameter);
        //        }
        //    }
        //    return ret.OrderBy(p => p.Name);
        //}
        #endregion
    }
}
