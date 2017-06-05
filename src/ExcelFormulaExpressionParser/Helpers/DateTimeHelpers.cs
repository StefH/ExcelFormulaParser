using System;

namespace ExcelFormulaExpressionParser.Helpers
{
    public static class DateTimeHelpers
    {
        public static DateTime FromOADate(double date)
        {
#if NETCORE || NETSTANDARD
            return new DateTime(DoubleDateToTicks(date), DateTimeKind.Unspecified);
#else
            return DateTime.FromOADate(date);
#endif
        }

        public static double ToOADate(DateTime date)
        {
#if NETCORE || NETSTANDARD
            return (date - new DateTime(1900, 1, 1)).TotalMilliseconds / (24 * 60 * 60 * 1000);
#else
            return date.ToOADate();
#endif
        }

        private static long DoubleDateToTicks(double value)
        {
            if (value >= 2958466.0 || value <= -657435.0)
            {
                throw new ArgumentException("Not a valid value");
            }

            long num1 = (long)(value * 86400000.0 + (value >= 0.0 ? 0.5 : -0.5));
            if (num1 < 0L)
                num1 -= num1 % 86400000L * 2L;

            long num2 = num1 + 59926435200000L;
            if (num2 < 0L || num2 >= 315537897600000L)
            {
                throw new ArgumentException("Not a valid value");
            }

            return num2 * 10000L;
        }
    }
}