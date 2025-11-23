using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace PnP.Framework.Extensions
{
    /// <summary>
    /// Provide general purpose extension methods
    /// </summary>
    public static class ObjectExtensions
    {

        /// <summary>
        /// Set an object field or property and returns if the value was changed.
        /// </summary>
        /// <typeparam name="TObject">Type of the target object</typeparam>
        /// <typeparam name="T">T of the property</typeparam>
        /// <param name="target">target object </param>
        /// <param name="propertyToSet">Expression to the property or field of the object</param>
        /// <param name="valueToSet">new value to set</param>
        /// <param name="allowNull">continue with set operation is null value is specified</param>
        /// <param name="allowEmpty">continue with set operation is null or empty value is specified</param>
        /// <returns><c>true</c> if the value has changed, otherwise <c>false</c></returns>
        public static bool Set<TObject, T>(this TObject target, Expression<Func<TObject, T>> propertyToSet, T valueToSet, bool allowNull = true, bool allowEmpty = true)
        {
            // Taken from https://stackoverflow.com/a/29092675/588868
            var members = new List<MemberInfo>();

            var exp = propertyToSet.Body;

            if (!allowNull && valueToSet == null)
            {
                return false;
            }

            if (!allowEmpty && (valueToSet is string) && (valueToSet == null || string.IsNullOrEmpty(valueToSet as string)))
            {
                return false;
            }

            while (exp != null)
            {
                var mi = exp as MemberExpression;

                if (mi != null)
                {
                    members.Add(mi.Member);
                    exp = mi.Expression;
                }
                else
                {
                    var pe = exp as ParameterExpression;

                    if (pe == null)
                    {
                        // We support only a ParameterExpression at the base
                        throw new NotSupportedException();
                    }

                    break;
                }
            }

            if (members.Count == 0)
            {
                // We need at least a getter
                throw new NotSupportedException();
            }

            // Now we must walk the getters (excluding the last).
            object targetObject = target;

            // We have to walk the getters from last (most inner) to second
            // (the first one is the one we have to use as a setter)
            for (int i = members.Count - 1; i >= 1; i--)
            {
                var pi = members[i] as PropertyInfo;

                if (pi != null)
                {
                    targetObject = pi.GetValue(targetObject);
                }
                else
                {
                    var fi = (FieldInfo)members[i];
                    targetObject = fi.GetValue(targetObject);
                }
            }

            // The first one is the getter we treat as a setter
            {
                var pi = members[0] as PropertyInfo;

                if (pi != null)
                {
                    var current = (T)pi.GetValue(targetObject);
                    if (!EqualityComparer<T>.Default.Equals(current, valueToSet))
                    {
                        pi.SetValue(targetObject, valueToSet);
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    var fi = (FieldInfo)members[0];
                    var current = (T)fi.GetValue(targetObject);
                    if (!EqualityComparer<T>.Default.Equals(current, valueToSet))
                    {
                        fi.SetValue(targetObject, valueToSet);
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
        }

        /// <summary>
        /// Nullify a string when it's an empty one
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string NullIfEmpty(this string value)
        {
            return string.IsNullOrEmpty(value) ? null : value;
        }

        /// <summary>
        /// Retrieves the value of a public, instance property 
        /// </summary>
        /// <param name="source">The source object</param>
        /// <param name="propertyName">The property name, case insensitive</param>
        /// <returns>The property value, if any</returns>
        public static Object GetPublicInstancePropertyValue(this object source, string propertyName)
        {
            return (source?.GetType()?.GetProperty(propertyName,
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.IgnoreCase)?
                .GetValue(source));
        }

        /// <summary>
        /// Retrieves a public, instance property 
        /// </summary>
        /// <param name="source">The source object</param>
        /// <param name="propertyName">The property name, case insensitive</param>
        /// <returns>The property, if any</returns>
        public static PropertyInfo GetPublicInstanceProperty(this object source, string propertyName)
        {
            return (source?.GetType()?.GetProperty(propertyName,
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.IgnoreCase));
        }

        /// <summary>
        /// Sets the value of a public, instance property 
        /// </summary>
        /// <param name="source">The source object</param>
        /// <param name="propertyName">The property name, case insensitive</param>
        /// <param name="value">The value to set</param>
        public static void SetPublicInstancePropertyValue(this object source, string propertyName, object value)
        {
            source?.GetType()?.GetProperty(propertyName,
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.Public |
                System.Reflection.BindingFlags.IgnoreCase)?
                .SetValue(source, value);
        }

        /// <summary>
        /// Compares two values for equality with special handling for SharePoint field types and null/empty string normalization
        /// </summary>
        /// <param name="a">First value to compare</param>
        /// <param name="b">Second value to compare</param>
        /// <param name="treatEmptyStringAsNull">Whether to treat empty/whitespace strings as null</param>
        /// <returns>True if values are considered equal</returns>
        public static bool ValuesEqual(object a, object b, bool treatEmptyStringAsNull = false)
        {
            if (ReferenceEquals(a, b))
            {
                return true;
            }

            if (a is null || b is null)
            {
                if (treatEmptyStringAsNull)
                {
                    if ((a is null && IsEmptyStringLike(b)) || (b is null && IsEmptyStringLike(a)))
                    {
                        return true;
                    }
                }
                return false;
            }

            a = Normalize(a, treatEmptyStringAsNull);
            b = Normalize(b, treatEmptyStringAsNull);

            if (a is IStructuralEquatable seA && b is IStructuralEquatable seB)
            {
                return StructuralComparisons.StructuralEqualityComparer.Equals(seA, seB);
            }

            return Equals(a, b);
        }

        /// <summary>
        /// Checks whether the provided object is an empty or whitespace string
        /// </summary>
        /// <param name="x">The object to check</param>
        /// <returns>True if the object is a string that is empty or consists only of whitespace</returns>
        private static bool IsEmptyStringLike(object x)
        {
            return x is string s && string.IsNullOrWhiteSpace(s);
        }

        /// <summary>
        /// Normalizes a value for comparison
        /// </summary>
        /// <param name="v">The value to normalize</param>
        /// <param name="treatEmptyStringAsNull">Whether to treat empty/whitespace strings as null</param>
        /// <returns>The normalized value</returns>
        private static object Normalize(object v, bool treatEmptyStringAsNull)
        {
            if (v == null)
            {
                return null;
            }

            switch (v)
            {
                case string s:
                    return (treatEmptyStringAsNull && string.IsNullOrWhiteSpace(s)) ? null : s;

                case bool b:
                    return b;

                case byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal:
                    return Convert.ToDecimal(v, CultureInfo.InvariantCulture);

                case DateTime dt:
                    var utc = (dt.Kind == DateTimeKind.Utc ? dt : DateTime.SpecifyKind(dt, DateTimeKind.Unspecified)).ToUniversalTime();
                    return new DateTime(utc.Year, utc.Month, utc.Day, utc.Hour, utc.Minute, utc.Second, DateTimeKind.Utc);

                case FieldUserValue u:
                    return u.LookupId;

                case FieldLookupValue l:
                    return l.LookupId;

                case IEnumerable<FieldLookupValue> multiLookup:
                    return multiLookup.Select(x => x?.LookupId ?? 0).OrderBy(x => x).ToArray();

                case TaxonomyFieldValue tx:
                    return tx.TermGuid?.Trim().ToLowerInvariant();

                case TaxonomyFieldValueCollection txc:
                    return txc
                        .Where(x => x != null && !string.IsNullOrEmpty(x.TermGuid))
                        .Select(x => x.TermGuid.Trim().ToLowerInvariant())
                        .OrderBy(g => g)
                        .ToArray();

                case FieldGeolocationValue geo:
                    return new ValueTuple<double, double, double, double>(
                        Math.Round(geo.Latitude, 6),
                        Math.Round(geo.Longitude, 6),
                        Math.Round(geo.Altitude, 2),
                        Math.Round(geo.Measure, 2));

                case FieldUrlValue url:
                    return new ValueTuple<string, string>(
                        url.Url?.Trim() ?? string.Empty,
                        url.Description?.Trim() ?? string.Empty);

                case string[] ss:
                    return ss
                        .Select(x => treatEmptyStringAsNull && string.IsNullOrWhiteSpace(x) ? null : x)
                        .OrderBy(x => x, StringComparer.Ordinal)
                        .ToArray();

                case IEnumerable enumerable when v is not string:
                {
                    var list = new List<object>();
                    foreach (var e in enumerable)
                    {
                        list.Add(Normalize(e, treatEmptyStringAsNull));
                    }
                    return list.ToArray();
                }

                default:
                    return v;
            }
        }
    }
}
