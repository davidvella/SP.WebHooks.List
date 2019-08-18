using System;
using System.Collections.Generic;

namespace SharePoint.WebHooks.Common.Extensions
{
    internal static class IListExtension
    {
        /// <summary>
        /// Splits the list into chucks from the <paramref name="collection"/>.
        /// </summary>
        /// <typeparam name="T"><see cref="Type"/> of elements in the <paramref name="collection"/>.</typeparam>
        /// <param name="collection">The collection from which elements are to be retrieved.</param>
        /// <param name="size"> The size of the individual chunks></param>
        internal static IEnumerable<IList<T>> SplitList<T>(IList<T> collection, int size)
        {
            for (var i = 0; i < collection.Count; i += size)
            {
                yield return collection.GetRange(i, Math.Min(size, collection.Count - i));
            }
        }

        /// <summary>
        /// Returns elements in the specified range from the <paramref name="collection"/>.
        /// </summary>
        /// <typeparam name="T"><see cref="Type"/> of elements in the <paramref name="collection"/>.</typeparam>
        /// <param name="collection">The collection from which elements are to be retrieved.</param>
        /// <param name="index">The 0-based index position in the <paramref name="collection"/> from which elements are to be retrieved.</param>
        /// <param name="count">The number of elements to be retrieved from the <paramref name="collection"/> starting at the <paramref name="index"/>.</param>
        /// <returns>An <see cref="IList{T}"/> object.</returns>
        internal static IList<T> GetRange<T>(this IList<T> collection, int index, int count)
        {
            List<T> result = new List<T>();

            for (int i = index; i < index + count; i++)
                result.Add(collection[i]);

            return result;
        }
    }
}
