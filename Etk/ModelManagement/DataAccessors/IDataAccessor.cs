namespace Etk.ModelManagement.DataAccessors
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using System.Collections.ObjectModel;

    public interface IDataAccessor
    {
        Type ReturnType {get;}

        MethodInfo MethodInfo {get;}
        List<ParameterInfo> ParametersInfo { get; }
        DataAccessorInstanceType InstanceType { get; }
        object Invoke(List<object> parameters);
    }
}
