using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HNRoadFormatConverter
{
    // 自定义 Chunk 扩展方法，兼容 .NET Framework 4.x
    public static class EnumerableExtensions
    {
        public static IEnumerable<IEnumerable<T>> Chunk<T>(this IEnumerable<T> source, int chunkSize)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (chunkSize <= 0) throw new ArgumentException("Chunk size must be positive.", nameof(chunkSize));

            var list = source.ToList();
            for (int i = 0; i < list.Count; i += chunkSize)
            {
                yield return list.Skip(i).Take(chunkSize);
            }
        }
    }
}
