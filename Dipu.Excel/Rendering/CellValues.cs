using System.Collections.Generic;
using System.Linq;

namespace Dipu.Excel.Rendering
{
    public static class CellValues
    {
        public static IEnumerable<IList<IList<CellValue>>> Chunks(this IEnumerable<CellValue> @this)
        {
            return @this
                .GroupBy(v => v.Row)
                .SelectMany(v =>
                {
                    return v.OrderBy(w => w.Column).Aggregate(new List<IList<CellValue>> {new List<CellValue>()},
                        (acc, w) =>
                        {
                            var list = acc[acc.Count - 1];
                            if (list.Count == 0 || list[list.Count - 1].Column + 1 == w.Column)
                            {
                                list.Add(w);
                            }
                            else
                            {
                                list = new List<CellValue> {w};
                                acc.Add(list);
                            }

                            return acc;
                        });
                })
                .GroupBy(v => v[0].Column)
                .SelectMany(v =>
                {
                    return v.OrderBy(w => w[0].Row).Aggregate(
                        new List<IList<IList<CellValue>>> {new List<IList<CellValue>>()},
                        (acc, w) =>
                        {
                            var list = acc[acc.Count - 1];
                            if (list.Count == 0 || list[list.Count - 1].Count == w.Count)
                            {
                                list.Add(w);
                            }
                            else
                            {
                                list = new List<IList<CellValue>> {w};
                                acc.Add(list);
                            }

                            return acc;
                        });
                });
        }
    }
}