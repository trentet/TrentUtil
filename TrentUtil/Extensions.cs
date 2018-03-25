using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using System.Threading;
using ExcelDataReader;

namespace TrentUtil
{
    public static class Extensions
    {
        /// <summary>
        /// Used with the Excel class
        /// </summary>
        /// <param name="self"></param>
        /// <returns></returns>
        public static ExcelDataTableConfiguration GetConfiguration(this IExcelDataReader self)
        {
            return new ExcelDataTableConfiguration { UseHeaderRow = true };
        }

        /// <summary>
        /// Randomly generates a Double between a min and max value
        /// </summary>
        /// <param name="random"></param>
        /// <param name="minValue"></param>
        /// <param name="maxValue"></param>
        /// <returns></returns>
        public static double NextDouble(this Random random, double minValue, double maxValue)
        {
            return random.NextDouble() * (maxValue - minValue) + minValue;
        }

        /// <summary>
        /// Gets the description associated with an enum value
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string GetEnumDescription(Enum value)
        {
            FieldInfo fi = value.GetType().GetField(value.ToString());

            DescriptionAttribute[] attributes =
                (DescriptionAttribute[])fi.GetCustomAttributes(
                typeof(DescriptionAttribute), false);

            if (attributes != null &&
                attributes.Length > 0)
                return attributes[0].Description;
            else
                return value.ToString();
        }

        public static class ThreadSafeRandom
        {
            [ThreadStatic] private static Random Local;

            public static Random ThisThreadsRandom
            {
                get { return Local ?? (Local = new Random(unchecked(Environment.TickCount * 31 + Thread.CurrentThread.ManagedThreadId))); }
            }
        }

        /// <summary>
        /// Shuffles a list
        /// </summary>
        /// <typeparam name="T">Type of the list</typeparam>
        /// <param name="list">List to be shuffled</param>
        public static void Shuffle<T>(this IList<T> list)
        {
            int n = list.Count;
            while (n > 1)
            {
                n--;
                int k = ThreadSafeRandom.ThisThreadsRandom.Next(n + 1);
                T value = list[k];
                list[k] = list[n];
                list[n] = value;
            }
        }
    }
}
