using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Etk.Tools.Extensions;

namespace Etk.Tools.Reflection
{
    /// <summary> Internal use</summary>
    public static class TypeHelpers
    {
        private static readonly object syncObj = new object();
        private static readonly Dictionary<Type, FieldInfo[]> fieldsByType = new Dictionary<Type, FieldInfo[]>();

        /// <summary>Return a type given a string having 'Type,Assembly' as pattern</summary>
        /// <param name="typeName">The string containing the type definition</param>
        /// <returns>The type if it's found. If not, an 'EtkException' exception</returns>
        public static Type GetType(string typeName)
        {
            if (string.IsNullOrEmpty(typeName))
                throw new EtkException("Type name cannot be null or empty");

            string[] bindingElements = typeName.Split(',');
            if (bindingElements.Count() > 2)
                throw new EtkException("The 'Type' search string must be 'Type,Assembly' or 'Type'");

            if(bindingElements.Count() == 2)
                return GetType(bindingElements[1], bindingElements[0]);
            else
                return GetTypeInternal(bindingElements[0]);
        }

        /// <summary>Return a type given an assembly and a type name</summary>
        /// <param name="assemblyName">The name of the assembly containing the type</param>
        /// <param name="typeName">The name of the type to find in the given assembly</param>
        /// <returns>The type if it's found. If not, an 'EtkException' exception</returns>
        public static Type GetType(string assemblyName, string typeName)
        {
            assemblyName = assemblyName.EmptyIfNull().Trim();
            typeName = typeName.EmptyIfNull().Trim();

            if (string.IsNullOrEmpty(typeName))
                throw new EtkException("'typeName' cannot be null or empty");

            Type ret = null;
            if (string.IsNullOrEmpty(assemblyName))
            {
                foreach (Assembly currentAssembly in AppDomain.CurrentDomain.GetAssemblies())
                {
                    ret = GetTypeInternal(currentAssembly, typeName);
                    if (ret != null)
                        break;
                }
            }
            else
            {
                Assembly assembly = Assembly.Load(assemblyName.Trim());
                ret = GetTypeInternal(assembly, typeName);
            }

            if (ret == null)
                throw new EtkException("Cannot find the requested Type");
            return ret;
        }

        /// <summary>Return a <see cref="System.Reflection.MethodInfo"/> given a type (not mandatory) and a method name</summary>
        /// <param name="inType">A <see cref="System.Type"/> (not mandatory) </param>
        /// <param name="methodName">If 'inType' supplied, contains name of the method to find, if not:  must be composed this way 'Type,[Assembly],Method'</param>
        /// <returns>The <see cref="System.Reflection.MethodInfo"/> if it's found. If not, an 'EtkException' exception</returns>
        public static MethodInfo GetMethod(Type inType, string methodName)
        {
            if (string.IsNullOrEmpty(methodName))
                throw new EtkException("Method name cannot be null or empty");

            string[] methodNameElements = methodName.Split(',');
            if (methodNameElements.Count() != 1 && methodNameElements.Count() != 3 && 
                (methodNameElements.Count() == 1 && inType == null))
                throw new EtkException("The 'method' separator is ',' and it must be composed this way 'Type,[Assembly],Method' or, if the type is supplied, 'Method'");

            Type type;
            if (methodNameElements.Count() == 1)
            {
                type = inType;
                methodName = methodNameElements[0].EmptyIfNull().Trim();
            }
            else
            {
                methodName = methodNameElements[2].EmptyIfNull().Trim();
                type = TypeHelpers.GetType(methodNameElements[1], methodNameElements[0]);
            }

            MethodInfo ret = type.GetMethod(methodName);
            if (ret == null)
                throw new EtkException("Cannot find the method");
            return ret;
        }


        /// <summary> Create an object from an instance of its <see cref="System.Type"/> base</summary>
        /// <typeparam name="T1">The <see cref="System.Type"/> of the base instance</typeparam>
        /// <typeparam name="T2">The <see cref="System.Type"/> of the object to return</typeparam>
        /// <param name="baseInstance">The intance use to create an instance of type 'T2' ('T2' must inherit from 'T1'</param>
        /// <returns>If 'baseInstance' is not null and 'T2' inherits from 'T1', a new instance of 'T2' based on 'baseInstance'. If not null</returns>
        public static T2 CreateFromBase<T1, T2>(T1 baseInstance) where T1 : class
                                                                 where T2 : class
        {
            T2 ret = null;

            if (baseInstance != null && typeof(T1).IsAssignableFrom(typeof(T2)))
            {
                ret = Activator.CreateInstance(typeof(T2), true) as T2;

                FieldInfo[] fieldsInfo;
                lock (syncObj)
                {
                    if (!fieldsByType.TryGetValue(typeof(T1), out fieldsInfo))
                    {
                        fieldsInfo = typeof(T1).GetFields(BindingFlags.NonPublic);
                        fieldsByType[typeof(T1)] = fieldsInfo;
                    }
                }

                if (fieldsInfo != null)
                {
                    foreach (FieldInfo fi in fieldsInfo)
                        fi.SetValue(ret, fi.GetValue(baseInstance));
                }
            }
            return ret;
        }

        private static Type GetTypeInternal(string typeName)
        {
            Type ret = null;
            foreach (Assembly assembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                ret = GetTypeInternal(assembly, typeName);
                if (ret != null)
                    break;
            }
            if (ret == null)
                throw new EtkException("Cannot find the requested Type");
            return ret;
        }


        private static Type GetTypeInternal(Assembly assembly, string typeName)
        {
            Type ret = null;
            if (assembly != null && !string.IsNullOrEmpty(typeName))
            {
                bool containPoint = typeName.Contains('.');
                foreach (Type type in assembly.GetTypes())
                {
                    if (containPoint)
                    {
                        if (typeName.Equals(type.FullName))
                        {
                            ret = type;
                            break;
                        }
                    }
                    else
                    {
                        if (typeName.Equals(type.Name))
                        {
                            ret = type;
                            break;
                        }
                    }
                }
            }
            return ret;
        }
    }
}
