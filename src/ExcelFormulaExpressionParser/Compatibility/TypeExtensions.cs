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

        public static bool IsNumeric(this Type type)
        {
            return GetNumericTypeKind(type) != 0;
        }

        static int GetNumericTypeKind(Type type)
        {
#if !(NETSTANDARD)
            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Char:
                case TypeCode.Single:
                case TypeCode.Double:
                case TypeCode.Decimal:
                    return 1;
                case TypeCode.SByte:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                    return 2;
                case TypeCode.Byte:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                    return 3;
                default:
                    return 0;
            }
#else
            if (type.GetTypeInfo().IsEnum)
                return 0;
            if (type == typeof(Char) || type == typeof(Single) || type == typeof(Double) || type == typeof(Decimal))
                return 1;
            if (type == typeof(SByte) || type == typeof(Int16) || type == typeof(Int32) || type == typeof(Int64))
                return 2;
            if (type == typeof(Byte) || type == typeof(UInt16) || type == typeof(UInt32) || type == typeof(UInt64))
                return 3;

            return 0;
#endif
        }
    }
}