using System;
using System.Linq;
using System.Reflection;

namespace ExcelFormulaExpressionParser.Compatibility
{
    public static class TypeExtensions
    {
        public static MethodInfo FindMethod(this Type type, string name, Type[] types)
        {
#if NETSTANDARD
            return type.GetTypeInfo().GetDeclaredMethods(name).Single(mi => mi.GetParameters().Select(p => p.ParameterType).SequenceEqual(types));
#else
            return type.GetMethod(name, types);
#endif
        }
    }
}