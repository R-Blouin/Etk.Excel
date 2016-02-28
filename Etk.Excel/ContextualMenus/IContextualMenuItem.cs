namespace Etk.Excel.ContextualMenus
{
    using System.Reflection;

    interface IContextualMenuItem: IContextualPart
    {
        string Caption { get; }
        bool BeginGroup { get; }
        int FaceId {get;}
        MethodInfo MethodInfo { get; }
    }
}
