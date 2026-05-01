using Google.OrTools.Sat;
using System.Collections.Generic;
using System.Linq;

namespace Stundenplan_V2
{
    public static class RoomConstraint
    {
        public static void Add(
            CpModel model,
            BoolVar[,] x,
            List<UnterrichtsBlock> blocks,
            Dictionary<string, int> fachraumLimit,
            int S)
        {
            for (int s = 0; s < S; s++)
            {
                foreach (var fg in fachraumLimit)
                {
                    var vars = blocks
                        .Select((b, i) => new { b, i })
                        .Where(xb => xb.b.Teile.Any(t => t.FachGruppe == fg.Key))
                        .Select(xb => x[xb.i, s]);

                    model.Add(LinearExpr.Sum(vars) <= fg.Value);
                }
            }
        }
    }
}