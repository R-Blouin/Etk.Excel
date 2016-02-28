namespace Etk.ModelManagement
{
    public interface  IModelProperty
    {
        IModelType Parent { get; }
        string Name { get; set; }
        string Description { get; set; }

        bool IsACollection { get; }
    }
}
