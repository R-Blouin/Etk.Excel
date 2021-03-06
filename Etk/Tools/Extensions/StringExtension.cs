﻿using System;
using System.IO;
using System.Xml.Serialization;

namespace Etk.Tools.Extensions
{
    /// <summary>
    /// Extension class for the class System.String.
    /// </summary>
    public static class StringExtension
    {
        /// <summary>If the input string is null return string.Empty</summary>
        /// <param name="str">Instance to check.</param>
        /// <returns>If the input string if null, return string.Empty, if not, returns the input string.</returns>
        public static string EmptyIfNull(this string input)
        {
            return string.IsNullOrEmpty(input) ? string.Empty : input; 
        }

        /// <summary>Try to deserialyze the string to type 'T'</summary>
        /// <param name="str">string (xml) to deserialyze.</param>
        /// <returns> If the input string is null or empty, return 'T' default.
        /// Exception if the deserialyzation failed else a new instance of 'T'.</returns>
        public static T Deserialize<T>(this string input)
        {
            if (string.IsNullOrEmpty(input))
                return default(T);
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(T));
                using (StringReader sr = new StringReader(input))
                {
                    return (T)serializer.Deserialize(sr);
                }
            }
            catch (Exception ex)
            {
                throw new EtkException($"Deserialize from xml '{input}' to UnderlyingType '{typeof(T).Name}' failed: {ex.Message} {(ex.InnerException == null ? string.Empty : ex.InnerException.Message)}");
            }
        }
    }
}
