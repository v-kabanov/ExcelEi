// /**********************************************************************************************
// Author:  Vasily Kabanov
// Created  2017-02-14
// Comment  
// **********************************************************************************************/
// 

using System;

namespace ExcelEi.Read
{
    public static class Conversion
    {
        /// <summary>
        ///     Convert cell value to desired type, including nullable structs.
        ///     When converting blank string to nullable struct (e.g. ' ' to int?) null is returned.
        ///     When attempted conversion fails exception is passed through.
        /// </summary>
        /// <typeparam name="T">
        ///     The type to convert to.
        /// </typeparam>
        /// <returns>
        ///     The <paramref name="value"/> converted to <typeparamref name="T"/>.
        /// </returns>
        /// <remarks>
        ///     If input is string, parsing is performed for output types of DateTime and TimeSpan, which if fails throws <see cref="FormatException"/>.
        ///     Another special case for output types of DateTime and TimeSpan is when input is double, in which case <see cref="DateTime.FromOADate"/>
        ///     is used for conversion. This special case does not work through other types convertible to double (e.g. integer or string with number).
        ///     In all other cases 'direct' conversion <see cref="Convert.ChangeType(object, Type)"/> is performed.
        /// </remarks>
        /// <exception cref="FormatException">
        ///     <paramref name="value"/> is string and its format is invalid for conversion (parsing fails)
        /// </exception>
        /// <exception cref="InvalidCastException">
        ///     <paramref name="value"/> is not string and direct conversion fails
        /// </exception>
        public static T GetTypedExcelValue<T>(object value)
        {
            if (value == null)
                return default(T);

            var fromType = value.GetType();
            var toType = typeof(T);
            var toNullableUnderlyingType = (toType.IsGenericType && toType.GetGenericTypeDefinition() == typeof(Nullable<>))
                ? Nullable.GetUnderlyingType(toType)
                : null;

            if (fromType == toType || fromType == toNullableUnderlyingType)
                return (T)value;

            // if converting to nullable struct and input is blank string, return null
            if (toNullableUnderlyingType != null && fromType == typeof(string) && ((string)value).Trim() == string.Empty)
                return default(T);

            toType = toNullableUnderlyingType ?? toType;

            if (toType == typeof(DateTime))
            {
                if (value is double)
                    return (T)(object)(DateTime.FromOADate((double)value));

                if (fromType == typeof(string))
                    return (T)(object)DateTime.Parse(value.ToString());

                if (fromType == typeof(TimeSpan))
                    return ((T)(object)(new DateTime(((TimeSpan)value).Ticks)));
            }
            else if (toType == typeof(TimeSpan))
            {
                if (value is double)
                    return (T)(object)(new TimeSpan(DateTime.FromOADate((double)value).Ticks));

                if (fromType == typeof(string))
                    return (T)(object)TimeSpan.Parse(value.ToString());

                if (fromType == typeof(DateTime))
                    return ((T)(object)(new TimeSpan(((DateTime)value).Ticks)));
            }

            return (T)Convert.ChangeType(value, toType);
        }
    }
}