using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace Excel.EasyExport
{
    public static class EasyExtensions
    {
        /// <summary>
        /// Collection of numeric types.
        /// </summary>
        private static readonly List<Type> NumericTypes = new List<Type>
        {
            typeof(decimal),
            typeof(byte), typeof(sbyte),
            typeof(short), typeof(ushort),
            typeof(int), typeof(uint),
            typeof(long), typeof(ulong),
            typeof(float), typeof(double)
        };

        /// <summary>
        /// Check if the given type is a numeric type.
        /// </summary>
        /// <param name="type">The type to be checked.</param>
        /// <returns><c>true</c> if it's numeric; otherwise <c>false</c>.</returns>
        public static bool IsNumeric(this Type type)
        {
            return NumericTypes.Contains(type);
        }

        public static object GetPropertyValue<T>(this T model, string fieldName)
        {
            if (string.IsNullOrEmpty(fieldName))
                return null;

            Type type = typeof(T);

            PropertyInfo property = type.GetProperty(fieldName);

            if (property == null)
                return null;

            return property.GetValue(model, null);

        }
    }
}
