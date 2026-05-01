using Google.OrTools.Sat;
using System.Collections.Generic;
using System.Linq;

namespace Stundenplan_V2
{
    public class BewertungsResultat
    {
        public int Quality;
        public int Early;
        public int Late;
        public int BadUnits;
        public List<string> Details = new();
    }

    public static class PlanBewertung
    {
        // -------------------------------------------------
        // Bewertung eines fertigen Plans
        // Gewichte aus Parameter-Tabelle (B6-B8)
        // -------------------------------------------------
        public static BewertungsResultat Berechne(
            int[,] belegung,
            List<UnterrichtsBlock> blocks,
            List<ZeitSlot> slots,
            int gewichtFrüh = 1,
            int gewichtSpät = 5,
            int gewichtPäd = 5)
        {
            var result = new BewertungsResultat();

            // -------------------------------------------------
            // Doppelstunden zählen
            // -------------------------------------------------
            for (int b = 0; b < blocks.Count; b++)
            {
                for (int s = 0; s < slots.Count - 1; s++)
                {
                    if (slots[s].WTag == slots[s + 1].WTag &&
                        slots[s].Stunde + 1 == slots[s + 1].Stunde)
                    {
                        if (belegung[b, s] == 1 && belegung[b, s + 1] == 1)
                        {
                            if (slots[s].Stunde <= 5)
                                result.Early++;
                            else
                                result.Late++;
                        }
                    }
                }
            }

            // -------------------------------------------------
            // späte pädagogische Einheiten
            // -------------------------------------------------
            var latePerUnit = new Dictionary<string, int>();
            var unitUnr = new Dictionary<string, int>();

            for (int b = 0; b < blocks.Count; b++)
            {
                var block = blocks[b];

                for (int s = 0; s < slots.Count; s++)
                {
                    if (belegung[b, s] != 1)
                        continue;

                    var countedClasses = new HashSet<string>();

                    foreach (var teil in block.Teile)
                    {
                        foreach (var k in teil.Klassen)
                        {
                            if (countedClasses.Contains(k))
                                continue;

                            countedClasses.Add(k);

                            string key = k + "|" + block.Zeilentext;

                            if (!latePerUnit.ContainsKey(key))
                            {
                                latePerUnit[key] = 0;
                                unitUnr[key] = block.UNr;
                            }

                            if (slots[s].Stunde >= 6)
                                latePerUnit[key]++;
                        }
                    }
                }
            }

            foreach (var kv in latePerUnit)
            {
                if (kv.Value >= 2)
                {
                    var parts = kv.Key.Split('|');
                    string klasse = parts[0];
                    string zeilentext = parts[1];
                    int unr = unitUnr[kv.Key];

                    result.BadUnits++;
                    result.Details.Add($"{klasse} | UNr {unr} | {zeilentext}");
                }
            }

            // -------------------------------------------------
            // Qualitätsfunktion mit konfigurierbaren Gewichten
            // -------------------------------------------------
            result.Quality =
                result.Early * gewichtFrüh
                - result.Late * gewichtSpät
                - result.BadUnits * gewichtPäd;

            return result;
        }

        // -------------------------------------------------
        // Solver-Version der späten pädagogischen Einheiten
        // -------------------------------------------------
        public static List<BoolVar> SolverSpaetePaedEinheiten(
            CpModel model,
            BoolVar[,] x,
            List<UnterrichtsBlock> blocks,
            List<ZeitSlot> slots)
        {
            var badVars = new List<BoolVar>();
            var paedEinheiten = new Dictionary<string, List<int>>();

            for (int b = 0; b < blocks.Count; b++)
            {
                var block = blocks[b];
                var seenClasses = new HashSet<string>();

                foreach (var t in block.Teile)
                {
                    foreach (var k in t.Klassen)
                    {
                        if (seenClasses.Contains(k))
                            continue;

                        seenClasses.Add(k);

                        string key = k + "|" + block.Zeilentext;

                        if (!paedEinheiten.ContainsKey(key))
                            paedEinheiten[key] = new List<int>();

                        paedEinheiten[key].Add(b);
                    }
                }
            }

            foreach (var kv in paedEinheiten)
            {
                var blockIds = kv.Value;
                var lateVars = new List<IntVar>();

                foreach (int b in blockIds)
                    for (int s = 0; s < slots.Count; s++)
                        if (slots[s].Stunde >= 6)
                            lateVars.Add(x[b, s]);

                if (lateVars.Count == 0)
                    continue;

                IntVar lateCount = model.NewIntVar(0, lateVars.Count, $"late_{kv.Key}");
                model.Add(lateCount == LinearExpr.Sum(lateVars));

                BoolVar bad = model.NewBoolVar($"bad_{kv.Key}");
                model.Add(lateCount >= 2).OnlyEnforceIf(bad);
                model.Add(lateCount <= 1).OnlyEnforceIf(bad.Not());

                badVars.Add(bad);
            }

            return badVars;
        }
    }
}
