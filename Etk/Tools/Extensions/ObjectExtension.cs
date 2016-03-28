using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace Etk.Tools.Extensions
{
    public static class ObjectExtension
    {
        /// <summary>     
        /// Perform a deep clone of an object.
        /// The object to clone must be serializable.
        /// </summary>     
        /// <typeparam name="T">The type of the object to clone.</typeparam>     
        /// <param name="source">The instance to copy.</param>     
        /// <returns>The cloned object.</returns>     
        public static T DeepClone<T>(T instance)
        {
            try
            {
                if (!typeof(T).IsSerializable)
                    throw new EtkException("The UnderlyingType to clone must be serializable.", false);

                // Don't serialize a null object,  return the type default value.
                if (Object.ReferenceEquals(instance, null))
                    return default(T);

                using (Stream stream = new MemoryStream())
                {
                    IFormatter formatter = new BinaryFormatter();
                    formatter.Serialize(stream, instance);
                    stream.Seek(0, SeekOrigin.Begin);
                    return (T)formatter.Deserialize(stream);
                }
            }
            catch (Exception ex)
            {
                throw new EtkException(string.Format("DeepClone failed for UnderlyingType '{0}'.{1}", typeof(T).Name, ex.Message));
            }
        }
    }
}
