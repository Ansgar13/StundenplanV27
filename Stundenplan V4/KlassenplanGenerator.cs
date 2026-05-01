using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;

namespace Stundenplan_V2
{
    public static class KlassenplanGenerator
    {
        public static void Erzeuge(
            string excelPfad,
            List<UnterrichtsBlock> unterrichtListe,
            List<ZeitSlot> zeitRaster,
            string suffix,
            HashSet<string> klassenFilter = null)
        {
            using var workbook = new XLWorkbook(excelPfad);

            string sheetName = "Klassenpläne_" + suffix;

            if (workbook.Worksheets.Any(ws => ws.Name == sheetName))
                workbook.Worksheet(sheetName).Delete();

            var sheet = workbook.Worksheets.Add(sheetName);

            sheet.Column(1).Width = 12;
            for (int i = 2; i <= 6; i++)
                sheet.Column(i).Width = 20;

            var tage = zeitRaster.Select(z => z.WTag).Distinct().ToList();
            var stunden = zeitRaster.Select(z => z.Stunde).Distinct().OrderBy(x => x).ToList();

            var alleKlassen = unterrichtListe
                .SelectMany(b => b.Teile)
                .SelectMany(t => t.Klassen)
                .Distinct()
                .Where(k => klassenFilter == null || klassenFilter.Contains(k))
                .OrderBy(x => x)
                .ToList();

            var blockLookup = unterrichtListe.ToDictionary(b => b.UNr);

            string Key(string lehrer, string fach, IEnumerable<string> klassen) =>
                lehrer + "|" + fach + "|" + string.Join(",", klassen.OrderBy(x => x));

            // spaeteDoppel: Key → minimale Stunde der zugehörigen Einzelstunde
            // (die dritte Stunde der pädagogischen Einheit)
            var spaeteDoppel = new Dictionary<string, int>();

            for (int i = 0; i < zeitRaster.Count - 1; i++)
            {
                var s1 = zeitRaster[i];
                var s2 = zeitRaster[i + 1];

                if (s1.WTag != s2.WTag) continue;
                if (s1.Stunde + 1 != s2.Stunde) continue;

                foreach (var u1 in s1.BelegteUNrn)
                {
                    var b1 = blockLookup[u1];

                    foreach (var t1 in b1.Teile)
                    {
                        foreach (var u2 in s2.BelegteUNrn)
                        {
                            var b2 = blockLookup[u2];

                            foreach (var t2 in b2.Teile)
                            {
                                if (t1.Lehrer == t2.Lehrer &&
                                    t1.Fach == t2.Fach &&
                                    t1.Klassen.OrderBy(x => x)
                                      .SequenceEqual(t2.Klassen.OrderBy(x => x)))
                                {
                                    if (s1.Stunde >= 5)
                                    {
                                        string k = Key(t1.Lehrer, t1.Fach, t1.Klassen);

                                        // Suche Einzelstunde desselben Lehrers/Fachs
                                        // an einem beliebigen Tag (nicht diese Doppelstunde)
                                        int einzelStunde = int.MaxValue;
                                        foreach (var sX in zeitRaster
                                            .Where(z => !(z.WTag == s1.WTag &&
                                                         (z.Stunde == s1.Stunde ||
                                                          z.Stunde == s2.Stunde))))
                                        {
                                            foreach (var uX in sX.BelegteUNrn)
                                            {
                                                var bX = blockLookup[uX];
                                                if (bX.Teile.Any(tX =>
                                                    tX.Lehrer == t1.Lehrer &&
                                                    tX.Fach == t1.Fach &&
                                                    tX.Klassen.OrderBy(x => x)
                                                      .SequenceEqual(t1.Klassen.OrderBy(x => x))))
                                                {
                                                    if (sX.Stunde < einzelStunde)
                                                        einzelStunde = sX.Stunde;
                                                }
                                            }
                                        }

                                        // Speichere die minimale Einzelstunde für diesen Key
                                        if (!spaeteDoppel.ContainsKey(k) ||
                                            einzelStunde < spaeteDoppel[k])
                                            spaeteDoppel[k] = einzelStunde;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            int startRow = 1;

            foreach (var klasse in alleKlassen)
            {
                int planStart = startRow;

                sheet.Cell(startRow++, 1).Value = klasse;
                sheet.Cell(startRow - 1, 1).Style.Font.Bold = true;

                sheet.Cell(startRow, 1).Value = "Stunde";
                sheet.Cell(startRow, 1).Style.Font.Bold = true;

                for (int t = 0; t < tage.Count; t++)
                {
                    var c = sheet.Cell(startRow, t + 2);
                    c.Value = tage[t];
                    c.Style.Font.Bold = true;
                    c.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }

                startRow++;

                foreach (var stunde in stunden)
                {
                    sheet.Row(startRow).Height = 45;
                    sheet.Cell(startRow, 1).Value = stunde;

                    for (int t = 0; t < tage.Count; t++)
                    {
                        string tag = tage[t];

                        var slot = zeitRaster
                            .FirstOrDefault(z => z.WTag == tag && z.Stunde == stunde);

                        if (slot == null) continue;

                        var cell = sheet.Cell(startRow, t + 2);

                        // Alle passenden Teile für diese Klasse in diesem Slot sammeln
                        var passendeTeileMitBlock = new List<(TeilUnterricht teil, UnterrichtsBlock block)>();

                        foreach (var u in slot.BelegteUNrn)
                        {
                            var block = blockLookup[u];
                            foreach (var teil in block.Teile)
                            {
                                if (!teil.Klassen.Contains(klasse))
                                    continue;
                                passendeTeileMitBlock.Add((teil, block));
                            }
                        }

                        if (passendeTeileMitBlock.Count == 0) goto NächsteZelle;

                        {
                            // Lehrer und Fächer zusammenbauen
                            string lehrer = string.Join(", ", passendeTeileMitBlock.Select(x => x.teil.Lehrer));
                            string fach   = string.Join(", ", passendeTeileMitBlock.Select(x => x.teil.Fach));
                            var erstBlock = passendeTeileMitBlock[0].block;
                            var erstTeil  = passendeTeileMitBlock[0].teil;
                            string key    = Key(erstTeil.Lehrer, erstTeil.Fach, erstTeil.Klassen);

                            cell.Clear();
                            var rt = cell.GetRichText();
                            rt.AddText(lehrer + "\n");
                            rt.AddText(fach + "\n");
                            rt.AddText($"UNr {erstBlock.UNr}    ");
                            var zt = rt.AddText(erstBlock.Zeilentext ?? "");
                            zt.Bold = true;
                            zt.FontSize = 13;

                            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            cell.Style.Alignment.Vertical   = XLAlignmentVerticalValues.Center;
                            cell.Style.Alignment.WrapText   = true;

                            // Doppelstunden-Erkennung anhand des ersten Teils
                            bool istDoppel = false;

                            var next = zeitRaster
                                .FirstOrDefault(z => z.WTag == tag && z.Stunde == stunde + 1);

                            if (next != null)
                            {
                                foreach (var u2 in next.BelegteUNrn)
                                {
                                    var b2 = blockLookup[u2];
                                    if (b2.Teile.Any(t2 =>
                                        t2.Lehrer == erstTeil.Lehrer &&
                                        t2.Fach   == erstTeil.Fach &&
                                        t2.Klassen.OrderBy(x => x)
                                           .SequenceEqual(erstTeil.Klassen.OrderBy(x => x))))
                                        istDoppel = true;
                                }
                            }

                            var prev = zeitRaster
                                .FirstOrDefault(z => z.WTag == tag && z.Stunde == stunde - 1);

                            if (prev != null)
                            {
                                foreach (var u2 in prev.BelegteUNrn)
                                {
                                    var b2 = blockLookup[u2];
                                    if (b2.Teile.Any(t2 =>
                                        t2.Lehrer == erstTeil.Lehrer &&
                                        t2.Fach   == erstTeil.Fach &&
                                        t2.Klassen.OrderBy(x => x)
                                           .SequenceEqual(erstTeil.Klassen.OrderBy(x => x))))
                                        istDoppel = true;
                                }
                            }

                            bool spaeteEinzel =
                                stunde >= 5 &&
                                !istDoppel &&
                                spaeteDoppel.ContainsKey(key);

                            bool istSpätePädEinheit =
                                spaeteDoppel.TryGetValue(key, out int einzelStd) &&
                                einzelStd >= 4 &&
                                einzelStd != int.MaxValue;

                            var fachTrim = erstTeil.Fach.Trim().ToUpper();

                            if (fachTrim.EndsWith("L1"))
                                cell.Style.Fill.BackgroundColor = XLColor.LightBlue;
                            else if (fachTrim.EndsWith("L2"))
                                cell.Style.Fill.BackgroundColor = XLColor.CornflowerBlue;
                            else if (spaeteEinzel && istSpätePädEinheit)
                                cell.Style.Fill.BackgroundColor = XLColor.Orange;
                            else if (istDoppel && stunde >= 5 && istSpätePädEinheit)
                                cell.Style.Fill.BackgroundColor = XLColor.Orange;
                            else if (istDoppel)
                                cell.Style.Fill.BackgroundColor = XLColor.Yellow;
                        }

                        NächsteZelle:

                        if (slot.KlassenWunsch.ContainsKey(klasse))
                            FärbeZelle(cell, slot.KlassenWunsch[klasse]);

                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        cell.Style.Alignment.WrapText = true;
                    }

                    startRow++;
                }

                var r = sheet.Range(planStart, 1, startRow - 1, tage.Count + 1);
                r.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                r.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                startRow += 2;
            }

            workbook.Save();
        }

        private static void FärbeZelle(IXLCell cell, int wert)
        {
            switch (wert)
            {
                case -3:
                    cell.Style.Border.DiagonalBorder = XLBorderStyleValues.Thick;
                    cell.Style.Border.DiagonalUp = true;
                    cell.Style.Border.DiagonalDown = true;
                    break;

                case -2: cell.Style.Fill.BackgroundColor = XLColor.Red; break;
                case -1: cell.Style.Fill.BackgroundColor = XLColor.LightPink; break;
                case 1: cell.Style.Fill.BackgroundColor = XLColor.LightGreen; break;
                case 2: cell.Style.Fill.BackgroundColor = XLColor.Green; break;
                case 3: cell.Style.Fill.BackgroundColor = XLColor.DarkGreen; break;
            }
        }
    }
}