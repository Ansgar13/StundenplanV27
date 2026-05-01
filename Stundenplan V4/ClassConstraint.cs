using Google.OrTools.Sat;
using System.Collections.Generic;
using System.Linq;

namespace Stundenplan_V2
{
    public static class ClassConstraint
    {
        public static void Add(
            CpModel model,
            BoolVar[,] x,
            List<UnterrichtsBlock> blocks,
            int S)
        {
            int B = blocks.Count;

            for (int s = 0; s < S; s++)
            {
                var map = new Dictionary<string, List<int>>();

                for (int b = 0; b < B; b++)
                {
                    foreach (var k in blocks[b].Teile.SelectMany(t => t.Klassen).Distinct())
                    {
                        if (!map.ContainsKey(k))
                            map[k] = new List<int>();

                        map[k].Add(b);
                    }
                }

                foreach (var kv in map)
                    model.Add(LinearExpr.Sum(kv.Value.Select(b => x[b, s])) <= 1);
            }
        }
    }
}