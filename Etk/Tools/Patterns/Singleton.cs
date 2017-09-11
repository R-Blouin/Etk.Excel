using System;

namespace Etk.Tools.Patterns
{
    /// <summary>Create a lazy singleton of type T </summary>
    public abstract class Singleton<T>  where T : class
    {
        private static readonly Lazy<T> instance = new Lazy<T>(() => Activator.CreateInstance(typeof(T), true) as T);

        public static T Instance => instance.Value;
    }

    /// <summary>Create a lazy singleton of type T1 inheriting from T2 </summary>
    public abstract class Singleton<T1, T2> where T1 : class
                                            where T2 : class
    {
        private static readonly Lazy<T1> instance = new Lazy<T1>( () => {
                                                                            try
                                                                            {
                                                                                if(! typeof(T1).IsAssignableFrom(typeof(T2)))
                                                                                    throw new EtkException($"'{typeof(T2).Name}' is not assignable from '{typeof(T1).Name}'");
                                                                                T2 t2 = Activator.CreateInstance(typeof(T2), true) as T2;
                                                                                return t2 as T1;
                                                                            }
                                                                            catch(Exception ex)
                                                                            {
                                                                                throw new EtkException($"Singleton creation failed: {ex.Message}");
                                                                            }
                                                                        });

        public static T1 Instance => instance.Value;
    }
}
