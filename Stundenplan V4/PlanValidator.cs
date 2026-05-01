using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Stundenplan_V2
{
    public static class PlanValidator
    {
        public record Verletzung(
            string Kategorie,
            string Tag,
            int Stunde,
            int UNr,
            string Lehrer,
            string Fach,
            string Details);

        public static List<Verletzung> Prüfe(
            int[,] belegung,
            List<UnterrichtsBlock> blocks,
            List<ZeitSlot> slots,
            List<(int stundeVor, int stundeNach)> grossePausen)
        {
            int B = blocks.Count;
            int S = slots.Count;
            var verletzungen = new List<Verletzung>();

            // Hilfsfunktionen
            string TagStunde(int s) => $"{slots[s].WTag} Std{slots[s].Stunde}";

            // Belegung: block → liste der Slot-Indizes
            var blockSlots = new Dictionary<int, List<int>>();
            for (int b = 0; b < B; b++)
            {
                blockSlots[b] = new List<int>();
                for (int s = 0; s < S; s++)
                    if (belegung[b, s] == 1)
                        blockSlots[b].Add(s);
            }

            // =====================================================
            // 1. WOCHENSTUNDEN: Block hat falsche Anzahl Slots
            // =====================================================
            for (int b = 0; b < B; b++)
            {
                int istWst = blockSlots[b].Count;
                int sollWst = blocks[b].Wst;
                if (istWst != sollWst)
                    verletzungen.Add(new Verletzung(
                        "Wochenstunden",
                        "", 0, blocks[b].UNr,
                        string.Join(", ", blocks[b].Teile.Select(t => t.Lehrer)),
                        blocks[b].Zeilentext,
                        $"Soll={sollWst}, Ist={istWst}"));
            }

            // =====================================================
            // 2. LEHRER-KONFLIKT: gleicher Lehrer in zwei Blöcken im gleichen Slot
            // =====================================================
            for (int s = 0; s < S; s++)
            {
                var lehrerInSlot = new Dictionary<string, List<int>>();
                for (int b = 0; b < B; b++)
                {
                    if (belegung[b, s] != 1) continue;
                    foreach (var t in blocks[b].Teile)
                    {
                        if (!lehrerInSlot.ContainsKey(t.Lehrer))
                            lehrerInSlot[t.Lehrer] = new List<int>();
                        lehrerInSlot[t.Lehrer].Add(b);
                    }
                }
                foreach (var kv in lehrerInSlot.Where(x => x.Value.Count > 1))
                    verletzungen.Add(new Verletzung(
                        "Lehrer-Konflikt",
                        slots[s].WTag, slots[s].Stunde,
                        0, kv.Key, "",
                        $"Blöcke: {string.Join(", ", kv.Value.Select(b => $"UNr{blocks[b].UNr}"))}"));
            }

            // =====================================================
            // 3. KLASSEN-KONFLIKT: gleiche Klasse in zwei Blöcken mit VERSCHIEDENER UNr im gleichen Slot
            // =====================================================
            for (int s = 0; s < S; s++)
            {
                var klassenInSlot = new Dictionary<string, List<int>>();
                for (int b = 0; b < B; b++)
                {
                    if (belegung[b, s] != 1) continue;
                    foreach (var t in blocks[b].Teile)
                        foreach (var k in t.Klassen)
                        {
                            if (!klassenInSlot.ContainsKey(k))
                                klassenInSlot[k] = new List<int>();
                            klassenInSlot[k].Add(b);
                        }
                }
                foreach (var kv in klassenInSlot.Where(x => x.Value.Count > 1))
                {
                    // Nur melden wenn verschiedene UNrn beteiligt sind
                    var unrn = kv.Value.Select(b => blocks[b].UNr).Distinct().ToList();
                    if (unrn.Count <= 1) continue;

                    verletzungen.Add(new Verletzung(
                        "Klassen-Konflikt",
                        slots[s].WTag, slots[s].Stunde,
                        0, "", kv.Key,
                        $"Blöcke: {string.Join(", ", kv.Value.Select(b => $"UNr{blocks[b].UNr}"))}"));
                }
            }

            // =====================================================
            // 4. ZEITWUNSCH-VERLETZUNG: Block in gesperrtem Slot (-3)
            // =====================================================
            for (int b = 0; b < B; b++)
            {
                foreach (int s in blockSlots[b])
                {
                    foreach (var t in blocks[b].Teile)
                    {
                        // Lehrer-Sperre
                        if (slots[s].LehrerWunsch.TryGetValue(t.Lehrer, out int lw) && lw == -3)
                            verletzungen.Add(new Verletzung(
                                "Zeitwunsch Lehrer",
                                slots[s].WTag, slots[s].Stunde,
                                blocks[b].UNr, t.Lehrer, blocks[b].Zeilentext,
                                $"Lehrer {t.Lehrer} hat -3 Sperre"));

                        // Klassen-Sperre
                        foreach (var k in t.Klassen)
                            if (slots[s].KlassenWunsch.TryGetValue(k, out int kw) && kw == -3)
                                verletzungen.Add(new Verletzung(
                                    "Zeitwunsch Klasse",
                                    slots[s].WTag, slots[s].Stunde,
                                    blocks[b].UNr, t.Lehrer, k,
                                    $"Klasse {k} hat -3 Sperre"));
                    }
                }
            }

            // =====================================================
            // 5. DOPPELSTUNDEN: minD/maxD verletzt
            // =====================================================
            for (int b = 0; b < B; b++)
            {
                int minD = blocks[b].Teile.Max(t => t.MinDoppel);
                int maxD = blocks[b].Teile.Max(t => t.MaxDoppel);
                if (minD == 0 && maxD == 0) continue;

                // Zähle tatsächliche Doppelstunden
                int doppelCount = 0;
                var slotsSorted = blockSlots[b].OrderBy(s => s).ToList();
                for (int i = 0; i < slotsSorted.Count - 1; i++)
                {
                    int s1 = slotsSorted[i];
                    int s2 = slotsSorted[i + 1];
                    if (slots[s1].WTag == slots[s2].WTag &&
                        slots[s1].Stunde + 1 == slots[s2].Stunde)
                        doppelCount++;
                }

                if (doppelCount < minD)
                    verletzungen.Add(new Verletzung(
                        "Doppelstunden",
                        "", 0, blocks[b].UNr,
                        string.Join(", ", blocks[b].Teile.Select(t => t.Lehrer)),
                        blocks[b].Zeilentext,
                        $"minD={minD}, maxD={maxD}, tatsächlich={doppelCount}"));
                else if (doppelCount > maxD)
                    verletzungen.Add(new Verletzung(
                        "Doppelstunden",
                        "", 0, blocks[b].UNr,
                        string.Join(", ", blocks[b].Teile.Select(t => t.Lehrer)),
                        blocks[b].Zeilentext,
                        $"minD={minD}, maxD={maxD}, tatsächlich={doppelCount}"));
            }

            // =====================================================
            // 6. PAUSEN-VERLETZUNG: Doppelstunde über große Pause ohne (E)
            // =====================================================
            if (grossePausen != null && grossePausen.Count > 0)
            {
                for (int b = 0; b < B; b++)
                {
                    if (blocks[b].DoppelÜberPauseErlaubt) continue;

                    var slotsSorted = blockSlots[b].OrderBy(s => s).ToList();
                    for (int i = 0; i < slotsSorted.Count - 1; i++)
                    {
                        int s1 = slotsSorted[i];
                        int s2 = slotsSorted[i + 1];
                        if (slots[s1].WTag != slots[s2].WTag) continue;
                        if (slots[s1].Stunde + 1 != slots[s2].Stunde) continue;

                        bool istPause = grossePausen.Any(p =>
                            p.stundeVor == slots[s1].Stunde &&
                            p.stundeNach == slots[s2].Stunde);

                        if (istPause)
                            verletzungen.Add(new Verletzung(
                                "Pausen-Verletzung",
                                slots[s1].WTag, slots[s1].Stunde,
                                blocks[b].UNr,
                                string.Join(", ", blocks[b].Teile.Select(t => t.Lehrer)),
                                blocks[b].Zeilentext,
                                $"Doppelstunde über Pause {slots[s1].Stunde}→{slots[s2].Stunde}"));
                    }
                }
            }

            // =====================================================
            // 7. TAGESREGEL: Block ohne Dopp an mehr als 1 Tag
            //                Block mit Dopp an mehr als 2 Stunden pro Tag
            // =====================================================
            for (int b = 0; b < B; b++)
            {
                int maxD = blocks[b].Teile.Max(t => t.MaxDoppel);

                var proTag = blockSlots[b]
                    .GroupBy(s => slots[s].WTag)
                    .ToDictionary(g => g.Key, g => g.Count());

                foreach (var kv in proTag)
                {
                    int limit = maxD > 0 ? 2 : 1;
                    if (kv.Value > limit)
                        verletzungen.Add(new Verletzung(
                            "Tagesregel",
                            kv.Key, 0, blocks[b].UNr,
                            string.Join(", ", blocks[b].Teile.Select(t => t.Lehrer)),
                            blocks[b].Zeilentext,
                            $"{kv.Value} Stunden an {kv.Key} (max {limit})"));
                }
            }

            // =====================================================
            // 8. FACHRAUM-LIMIT: zu viele Blöcke einer Fachgruppe gleichzeitig
            // =====================================================
            // (wird über fachraumLimit-Dictionary geprüft – hier vereinfacht)

            return verletzungen;
        }

        public static void SchreibeTabelle(
            string excelPfad,
            List<Verletzung> verletzungen)
        {
            using var wb = new XLWorkbook(excelPfad);

            const string sheetName = "Verletzungen";
            if (wb.Worksheets.Any(ws => ws.Name == sheetName))
                wb.Worksheet(sheetName).Delete();

            var sheet = wb.Worksheets.Add(sheetName);

            // Header
            var headers = new[] { "Kategorie", "Tag", "Stunde", "UNr", "Lehrer/Klasse", "Fach/ZeilenText", "Details" };
            for (int i = 0; i < headers.Length; i++)
            {
                sheet.Cell(1, i + 1).Value = headers[i];
                sheet.Cell(1, i + 1).Style.Font.Bold = true;
                sheet.Cell(1, i + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
            }

            if (verletzungen.Count == 0)
            {
                sheet.Cell(2, 1).Value = "✓ Keine Verletzungen gefunden";
                sheet.Cell(2, 1).Style.Fill.BackgroundColor = XLColor.LightGreen;
                sheet.Cell(2, 1).Style.Font.Bold = true;
                sheet.Range(2, 1, 2, headers.Length).Merge();
            }
            else
            {
                // Farben pro Kategorie
                var farben = new Dictionary<string, XLColor>
                {
                    ["Wochenstunden"]    = XLColor.LightPink,
                    ["Lehrer-Konflikt"] = XLColor.OrangeRed,
                    ["Klassen-Konflikt"]= XLColor.Orange,
                    ["Zeitwunsch Lehrer"]= XLColor.LightYellow,
                    ["Zeitwunsch Klasse"]= XLColor.LightYellow,
                    ["Doppelstunden"]   = XLColor.LightBlue,
                    ["Pausen-Verletzung"]= XLColor.Plum,
                    ["Tagesregel"]      = XLColor.LightSalmon,
                };

                for (int i = 0; i < verletzungen.Count; i++)
                {
                    var v = verletzungen[i];
                    int zeile = i + 2;
                    var farbe = farben.TryGetValue(v.Kategorie, out var f) ? f : XLColor.White;

                    sheet.Cell(zeile, 1).Value = v.Kategorie;
                    sheet.Cell(zeile, 2).Value = v.Tag;
                    sheet.Cell(zeile, 3).Value = v.Stunde > 0 ? v.Stunde.ToString() : "";
                    sheet.Cell(zeile, 4).Value = v.UNr > 0 ? v.UNr.ToString() : "";
                    sheet.Cell(zeile, 5).Value = v.Lehrer;
                    sheet.Cell(zeile, 6).Value = v.Fach;
                    sheet.Cell(zeile, 7).Value = v.Details;

                    for (int c = 1; c <= headers.Length; c++)
                        sheet.Cell(zeile, c).Style.Fill.BackgroundColor = farbe;
                }

                // Zusammenfassung oben
                var gruppen = verletzungen
                    .GroupBy(v => v.Kategorie)
                    .OrderByDescending(g => g.Count());
                int sumZeile = verletzungen.Count + 3;
                sheet.Cell(sumZeile, 1).Value = $"Gesamt: {verletzungen.Count} Verletzungen";
                sheet.Cell(sumZeile, 1).Style.Font.Bold = true;
                int row = sumZeile + 1;
                foreach (var g in gruppen)
                {
                    sheet.Cell(row, 1).Value = g.Key;
                    sheet.Cell(row, 2).Value = g.Count();
                    row++;
                }
            }

            sheet.Columns().AdjustToContents();
            wb.Save();
        }
    }
}
