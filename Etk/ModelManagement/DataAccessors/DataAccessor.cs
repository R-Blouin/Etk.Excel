using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Etk.Tools.Reflection;

namespace Etk.ModelManagement.DataAccessors
{
    public class DataAccessor : IDataAccessor
    {
        private object callingInstance;

        public Type ReturnType         
        { get ; private set; }

        public MethodInfo MethodInfo
        { get; private set; }

        public List<ParameterInfo> ParametersInfo
        { get; private set; }

        public DataAccessorInstanceType InstanceType
        { get; private set; }

        #region static public methods
        public object Invoke(List<object> parameters)
        {
            if (parameters != null && parameters.Count > ParametersInfo.Count())
                throw new ArgumentException("Too many parameters in input");
            if (parameters != null && parameters.Count > 0 && ParametersInfo.Count() == 0)
                throw new ArgumentException("Too many parameters in input");

            if (ParametersInfo == null || ParametersInfo.Count() == 0)
                return MethodInfo.Invoke(callingInstance, null);

            if (parameters == null || parameters.Count == 0)
                return MethodInfo.Invoke(callingInstance, new object[ParametersInfo.Count()]);

            object[] realParameters = new object[ParametersInfo.Count()];
            for (int i = 0; i < parameters.Count(); i++)
                realParameters[i] = TypeConvertor.ConvertObject(ParametersInfo[i].ParameterType, parameters[i]);

            return MethodInfo.Invoke(callingInstance, realParameters);
        }
        #endregion

        #region static public methods
        public static IDataAccessor CreateInstance(string bindingMethod, DataAccessorInstanceType dataAccessorInstanceType, string instanceName)
        {
            try
            {
                if (dataAccessorInstanceType == DataAccessorInstanceType.Singleton && string.IsNullOrEmpty(instanceName))
                    throw new EtkException("For 'Singleton', 'InstanceName' is mandatory");

                MethodInfo methodInfo = null;
                object callingInstance = null;
                switch (dataAccessorInstanceType)
                {
                    case DataAccessorInstanceType.Singleton:
                    case DataAccessorInstanceType.Static:
                        methodInfo = TypeHelpers.GetMethod(null, bindingMethod);

                        if (dataAccessorInstanceType == DataAccessorInstanceType.Singleton)
                        {
                            PropertyInfo pi = methodInfo.DeclaringType.GetProperties(BindingFlags.Static).FirstOrDefault(p => p.Name.Equals(instanceName));
                            if (pi != null)
                                callingInstance = pi.GetGetMethod().Invoke(null, null);
                            else
                            {
                                MethodInfo mi = methodInfo.DeclaringType.GetMethods(BindingFlags.Static).FirstOrDefault(m => m.Name.Equals(instanceName));
                                if (mi != null)
                                    callingInstance = mi.Invoke(null, null);
                                else
                                    throw new Exception("'InstanceName' not found");
                            }
                        }
                        break;
                }

                DataAccessor dataAccessor = new DataAccessor();
                dataAccessor.callingInstance = callingInstance;
                dataAccessor.InstanceType = dataAccessorInstanceType;
                dataAccessor.MethodInfo = methodInfo;
                ParameterInfo[] parameters = methodInfo.GetParameters();
                dataAccessor.ParametersInfo = new List<ParameterInfo>(parameters == null ? new List<ParameterInfo>() : new List<ParameterInfo>(parameters));
                dataAccessor.ReturnType = methodInfo.ReturnType;

                return dataAccessor;
            }
            catch (Exception ex)
            {
                throw new EtkException(string.Format("Connat create Data Accessor for '{0}':{1}", bindingMethod, ex.Message));
            }
        }
        #endregion

        #region private static Methods
        public static DataAccessorInstanceType AccessorInstanceTypeFrom(string from)
        {
            if (string.IsNullOrEmpty(from))
                return DataAccessorInstanceType.Static;

            switch (from.ToUpper())
            {
                case "STATIC":
                    return DataAccessorInstanceType.Static;

                case "SINGLETON":
                    return DataAccessorInstanceType.Singleton;

                default:
                    throw new EtkException(string.Format("DataAccessorInstanceType instance '{0}' is not valid", from));
            }
        }
        #endregion
    }
}
