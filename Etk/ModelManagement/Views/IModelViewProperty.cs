namespace Etk.ModelManagement.Views
{
    public interface IModelViewProperty : IModelViewPart
    {
        IModelProperty ModelProperty { get;}
        string Name { get; }
        //bool IsComposed { get; }

        //void ResolveDependencies();
    }
}
