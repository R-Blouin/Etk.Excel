using System;
using System.Collections.Generic;
using System.Reflection;

namespace Etk.ModelManagement.DataAccessors
{
    public interface IDataAccessor
    {
        Type ReturnType {get;}

        MethodInfo MethodInfo {get;}
        List<ParameterInfo> ParametersInfo { get; }
        DataAccessorInstanceType InstanceType { get; }
        object Invoke(List<object> parameters);
    }
}
