using Google.OrTools.Sat;
using System.Collections.Generic;

namespace Stundenplan_V2
{
    public static class TimeConstraint
    {
        public static void AddBlockedSlots(
            CpModel model,
            BoolVar[,] x,
            List<UnterrichtsBlock> blocks,
            List<ZeitSlot> slots,
            int B,
            int S)
        {
            for (int b = 0; b < B; b++)
            {
                for (int s = 0; s < S; s++)
                {
                    foreach (var t in blocks[b].Teile)
                    {
                        // Lehrer gesperrt
                        if (slots[s].LehrerWunsch.TryGetValue(t.Lehrer, out int lw) && lw == -3)
                            model.Add(x[b, s] == 0);

                        // Klassen gesperrt
                        foreach (var k in t.Klassen)
                        {
                            if (slots[s].KlassenWunsch.TryGetValue(k, out int kw) && kw == -3)
                                model.Add(x[b, s] == 0);
                        }
                    }
                }
            }
        }
    }
}