using System;
using System.Reflection;

namespace Helper
{
    public static class ReflactionHelper
    {
        public static Type GetPropertyType(PropertyInfo prop)
        {
            var propertyType = prop.PropertyType;
            return Nullable.GetUnderlyingType(propertyType) ?? propertyType;
        }
    }
}
