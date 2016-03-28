using System.Reflection;

namespace Etk.Excel.ContextualMenus
{
    interface IContextualMenuItem: IContextualPart
    {
        string Caption { get; }
        bool BeginGroup { get; }
        int FaceId {get;}
        MethodInfo MethodInfo { get; }
    }
}
