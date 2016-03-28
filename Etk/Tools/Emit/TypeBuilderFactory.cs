using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;

namespace Etk.Tools.Emit
{
    public class TypeBuilderFactory
    {
        private AssemblyBuilder assemblyBuilder = null;
        private ModuleBuilder moduleBuilder = null;

        #region .ctors
        public TypeBuilderFactory(string name)
        {
            try
            {
                if (string.IsNullOrEmpty(name))
                    throw new ArgumentNullException("'name' parameter cannot be null or empty");

                AssemblyName assemblyName = new AssemblyName(name);
                assemblyBuilder = AppDomain.CurrentDomain.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
                moduleBuilder = assemblyBuilder.DefineDynamicModule("Builder");
            }
            catch (Exception ex)
            {
                throw new EtkException(string.Format("'TypeBuilderFactory' creation failed:{0}", ex.Message)); 
            }
        }
        #endregion

        #region public methods
        public Type CreateType(string typeName, IEnumerable<EmitProperty> properties)
        {
            try
            {
                if (string.IsNullOrEmpty(typeName))
                    throw new ArgumentNullException("'typeName' parameter cannot be null or empty");

                if (properties == null || properties.Count() == 0)
                    throw new ArgumentNullException("'properties' parameter cannot be null or empty");

                TypeBuilder typeBuilder = CreateTypeBuilder(typeName);
                ConstructorBuilder constructor = typeBuilder.DefineDefaultConstructor(MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.RTSpecialName);

                foreach (EmitProperty property in properties)
                    CreateProperty(typeBuilder, property.PropertyName, property.PropertyType);

                return typeBuilder.CreateType();
            }
            catch (Exception ex)
            {
                throw new EtkException(string.Format("'TypeBuilderFactory.CreateType' failed:{0}", ex.Message));
            }
        }
        #endregion

        #region private methods
        private TypeBuilder CreateTypeBuilder(string typeName)
        {
            TypeBuilder typeBuilder = moduleBuilder.DefineType(typeName,
                                                               TypeAttributes.Public | TypeAttributes.Class | TypeAttributes.AutoClass | 
                                                               TypeAttributes.AnsiClass | TypeAttributes.BeforeFieldInit | TypeAttributes.AutoLayout,
                                                               null);
            return typeBuilder;
        }

        private static void CreateProperty(TypeBuilder typeBuilder, string propertyName, Type propertyType)
        {
            FieldBuilder fieldBuilder = typeBuilder.DefineField(propertyName + "_", propertyType, FieldAttributes.Private);

            PropertyBuilder propertyBuilder = typeBuilder.DefineProperty(propertyName, PropertyAttributes.HasDefault, propertyType, null);
            MethodBuilder getBuilder = typeBuilder.DefineMethod("get_" + propertyName, MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.HideBySig, propertyType, Type.EmptyTypes);
            ILGenerator getGenerator = getBuilder.GetILGenerator();

            getGenerator.Emit(OpCodes.Ldarg_0);
            getGenerator.Emit(OpCodes.Ldfld, fieldBuilder);
            getGenerator.Emit(OpCodes.Ret);

            MethodBuilder setBuilder = typeBuilder.DefineMethod("set_" + propertyName,
                                                                 MethodAttributes.Public | MethodAttributes.SpecialName | MethodAttributes.HideBySig,
                                                                 null, 
                                                                 new[] { propertyType });

            ILGenerator setGenerator = setBuilder.GetILGenerator();
            Label modifyProperty = setGenerator.DefineLabel();
            Label exitSet = setGenerator.DefineLabel();

            setGenerator.MarkLabel(modifyProperty);
            setGenerator.Emit(OpCodes.Ldarg_0);
            setGenerator.Emit(OpCodes.Ldarg_1);
            setGenerator.Emit(OpCodes.Stfld, fieldBuilder);

            setGenerator.Emit(OpCodes.Nop);
            setGenerator.MarkLabel(exitSet);
            setGenerator.Emit(OpCodes.Ret);

            propertyBuilder.SetGetMethod(getBuilder);
            propertyBuilder.SetSetMethod(setBuilder);
        }
        #endregion
    }
}
