using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Stundenplan_V2
{
    public static class ExcelLoader
    {
        public static StundenplanInput Lade(string excelPfad)
        {
            var unterrichtListe = new List<UnterrichtsBlock>();
            var zeitRaster = new List<ZeitSlot>();
            var fachgruppenRaeume = new Dictionary<string, int>();
            var extraFreieTage = new Dictionary<string, int>();
            using var workbook = new XLWorkbook(excelPfad);

            // =====================================================
            // TABELLE 1 – UNTERRICHT
            // =====================================================

            var sheet1 = workbook.Worksheet("U-Verteilung");
            var header1 = GetHeaderMap(sheet1);

            System.Diagnostics.Debug.WriteLine("=== HEADER U-Verteilung ===");

            foreach (var h in header1.Keys)
            {
                System.Diagnostics.Debug.WriteLine($"'{h}'");
            }



            var rows1 = sheet1.RangeUsed().RowsUsed().Skip(1).ToList();


            // 🔍 DEBUG HIER
            var firstRow = rows1.FirstOrDefault();

            if (firstRow != null)
            {
                System.Diagnostics.Debug.WriteLine("=== ERSTE DATENZEILE ===");

                foreach (var h in header1)
                {
                    string value = firstRow.Cell(h.Value).GetString();
                    System.Diagnostics.Debug.WriteLine($"{h.Key}: '{value}'");
                }
            }





            var alleTeile = new List<TeilUnterricht>();
            // UNrn die mindestens eine aktive (nicht-i) Zeile haben
            var aktivUNrn = new HashSet<int>();
            // UNrn die nur i-Zeilen haben → komplett ignoriert (für Fix-UNrn-Filter)
            var ignorierteUNrn = new HashSet<int>();

            // Erst-Durchlauf: welche UNrn haben aktive Zeilen?
            foreach (var row in rows1)
            {
                if (!int.TryParse(Cell(row, header1, "U-Nr").GetString(), out int uNr))
                    continue;
                string ignoreWert = GetOptional(row, header1, "Ignore (i)").Trim().ToLower();
                if (ignoreWert != "i")
                    aktivUNrn.Add(uNr);
                else
                    ignorierteUNrn.Add(uNr);
            }
            // Nur UNrn die ausschließlich i-Zeilen haben sind wirklich ignoriert
            ignorierteUNrn.ExceptWith(aktivUNrn);

            foreach (var row in rows1)
            {
                if (!int.TryParse(Cell(row, header1, "U-Nr").GetString(), out int uNr))
                    continue;

                // Ignore-Spalte prüfen: steht "i" drin → nur diese Zeile überspringen
                // (nicht die gesamte UNr – andere Zeilen der UNr können aktiv bleiben)
                string ignoreWert = GetOptional(row, header1, "Ignore (i)").Trim().ToLower();
                if (ignoreWert == "i")
                    continue;

                int wst = Cell(row, header1, "Wst").GetValue<int>();
                string lehrer = Cell(row, header1, "Lehrer").GetString();
                string fach = Cell(row, header1, "Fach").GetString();
                string klassenRaw = Cell(row, header1, "Klasse(n)").GetString();
                string dStdRaw = GetOptional(row, header1, "Dopp.Std.");
                string ltkz = GetOptional(row, header1, "LTKZ");
                string eWert = GetOptional(row, header1, "(E)").Trim().ToLower();

                int minD = 0;
                int maxD = 0;

                if (!string.IsNullOrWhiteSpace(dStdRaw))
                {
                    var teile = dStdRaw.Split('-');

                    if (teile.Length == 2)
                    {
                        int.TryParse(teile[0], out minD);
                        int.TryParse(teile[1], out maxD);
                    }
                }

                var klassenListe = klassenRaw
                    .Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(k => k.Trim())
                    .ToList();

                alleTeile.Add(new TeilUnterricht
                {
                    UNr = uNr,
                    Lehrer = lehrer,
                    Fach = fach,
                    Klassen = klassenListe,
                    MinDoppel = minD,
                    MaxDoppel = maxD,
                    FachGruppe = BestimmeFachgruppe(fach),
                    Ltkz = ltkz,
                    DoppelÜberPauseErlaubt = eWert == "x"
                });
            }

            var gruppen = alleTeile.GroupBy(t => t.UNr);

            foreach (var gruppe in gruppen)
            {
                int uNr = gruppe.Key;

                // Bereits durch Ignore-Check gefiltert – nur zur Sicherheit
                if (ignorierteUNrn.Contains(uNr)) continue;

                // Wst und Zeilentext aus der ersten AKTIVEN Zeile lesen
                var ersteAktiveZeile = rows1.FirstOrDefault(r =>
                    int.TryParse(Cell(r, header1, "U-Nr").GetString(), out int val) &&
                    val == uNr &&
                    GetOptional(r, header1, "Ignore (i)").Trim().ToLower() != "i");

                if (ersteAktiveZeile == null) continue;

                int wst = Cell(ersteAktiveZeile, header1, "Wst").GetValue<int>();
                string zeilentext = GetOptional(ersteAktiveZeile, header1, "ZeilenText");

                unterrichtListe.Add(new UnterrichtsBlock
                {
                    UNr = uNr,
                    Wst = wst,
                    Zeilentext = zeilentext,
                    Teile = gruppe.ToList(),
                    WochenDoppelstunden = 0,
                    TagesDoppelstunden = new Dictionary<string, int>(),
                    DoppelÜberPauseErlaubt = gruppe.Any(t => t.DoppelÜberPauseErlaubt)
                });
            }

            // =====================================================
            // TABELLE 2 – ZEITRASTER
            // =====================================================

            var sheet2 = workbook.Worksheet("Lösungen");
            var rows2 = sheet2.RangeUsed().RowsUsed().Skip(1);

            foreach (var row in rows2)
            {
                string wtag = row.Cell(1).GetString();

                if (!int.TryParse(row.Cell(2).GetString(), out int stunde))
                    continue;

                zeitRaster.Add(new ZeitSlot
                {
                    WTag = wtag,
                    Stunde = stunde
                });
            }

            // schneller Lookup für Slots
            var slotLookup = zeitRaster.ToDictionary(
                z => $"{z.WTag}_{z.Stunde}",
                z => z
            );

            // =====================================================
            // FIXUNR EINLESEN
            // =====================================================

            if (workbook.Worksheets.Any(ws => ws.Name == "Fix UNrn"))
            {
                var sheetFix = workbook.Worksheet("Fix UNrn");

                foreach (var row in sheetFix.RangeUsed().RowsUsed().Skip(1))
                {
                    string wtag = row.Cell(1).GetString().Trim();

                    if (!int.TryParse(row.Cell(2).GetString(), out int stunde))
                        continue;

                    string key = $"{wtag}_{stunde}";

                    if (!slotLookup.TryGetValue(key, out var slot))
                        continue;

                    int lastCol = row.LastCellUsed().Address.ColumnNumber;

                    for (int c = 3; c <= lastCol; c++)
                    {
                        if (int.TryParse(row.Cell(c).GetString(), out int unr))
                        {
                            // Ignorierte UNrn werden auch aus Fix-Slots herausgefiltert
                            if (!ignorierteUNrn.Contains(unr))
                                slot.FixUNrn.Add(unr);
                        }
                    }
                }
            }

            // =====================================================
            // ZEITWÜNSCHE
            // =====================================================

            if (workbook.Worksheets.Any(ws => ws.Name == "ZeitWL"))
                LeseZeitWunschTabelle(workbook.Worksheet("ZeitWL"), zeitRaster, true, extraFreieTage);

            if (workbook.Worksheets.Any(ws => ws.Name == "ZeitWK"))
                LeseZeitWunschTabelle(workbook.Worksheet("ZeitWK"), zeitRaster, false, extraFreieTage);

            // =====================================================
            // FACHGRUPPENRÄUME
            // =====================================================

            if (workbook.Worksheets.Any(ws => ws.Name == "Fachgruppenräume"))
            {
                var sheetFG = workbook.Worksheet("Fachgruppenräume");

                foreach (var row in sheetFG.RangeUsed().RowsUsed().Skip(1))
                {
                    string gruppe = row.Cell(1).GetString().Trim();
                    int anzahl = row.Cell(2).GetValue<int>();

                    if (!string.IsNullOrWhiteSpace(gruppe))
                        fachgruppenRaeume[gruppe] = anzahl;
                }
            }

            // =====================================================
            // DEBUG FIXUNR
            // =====================================================

            foreach (var s in zeitRaster)
            {
                if (s.FixUNrn.Count > 0)
                    System.Diagnostics.Debug.WriteLine(
     $"FIX: {s.WTag} {s.Stunde} -> {string.Join(",", s.FixUNrn)}");
            }

            //VerteileFreieTage(extraFreieTage, zeitRaster);

            // =====================================================
            // PARAMETER-SHEET
            // B1 = ZeitlimitSekunden
            // B3 = AnzahlLösungenOhneTausch
            // B4 = AnzahlLösungenMitTausch
            // =====================================================
            int zeitlimit = 30;
            int anzahlOhne = 2;
            int anzahlMit = 2;
            var nichtFreieTage = new HashSet<string>();
            int gewichtFrüh = 1;
            int gewichtSpät = 5;
            int gewichtPäd = 5;
            int gewichtFrei = 2;
            int strafeHohl = 1;
            int strafeDoppelHohl = 5;
            int strafeDreifachHohl = 5;
            int strafeStdFolge = 5;
            int strafeEinzel = 0;
            int strafeSpäteLk = 0;
            bool verbotSpäteDoppel = false;
            int hauptfachSpätAnteil = 50;
            int strafeHauptfachSpät = 0;
            var grossePausen = new List<(int stundeVor, int stundeNach)>();

            if (workbook.Worksheets.Any(ws => ws.Name == "Parameter"))
            {
                var sheetParam = workbook.Worksheet("Parameter");

                // Parameter per Beschriftung in Spalte A suchen (robuster als feste Zeilennummern)
                foreach (var row in sheetParam.RangeUsed()?.RowsUsed() ?? Enumerable.Empty<IXLRangeRow>())
                {
                    string label = row.Cell(1).GetString().Trim().ToLower();
                    string wert  = row.Cell(2).GetString().Trim();

                    if (label.Contains("zeitlimit"))
                        int.TryParse(wert, out zeitlimit);
                    else if (label.Contains("ohne tausch"))
                        int.TryParse(wert, out anzahlOhne);
                    else if (label.Contains("mit tausch"))
                        int.TryParse(wert, out anzahlMit);
                    else if (label.Contains("nichtfreieta") || label.Contains("freiet"))
                    {
                        if (!string.IsNullOrWhiteSpace(wert))
                            nichtFreieTage = new HashSet<string>(
                                wert.Split(',').Select(t => t.Trim()).Where(t => !string.IsNullOrEmpty(t)),
                                StringComparer.OrdinalIgnoreCase);
                    }
                    else if (label.Contains("frühe"))
                        int.TryParse(wert, out gewichtFrüh);
                    else if (label.Contains("späte dopp") || label.Contains("strafe späte dopp"))
                        int.TryParse(wert, out gewichtSpät);
                    else if (label.Contains("pädagog") || label.Contains("päd"))
                        int.TryParse(wert, out gewichtPäd);
                    else if (label.Contains("belohnung") || label.Contains("freie tage"))
                        int.TryParse(wert, out gewichtFrei);
                    else if (label.Contains("dreifachhohlstunde"))
                        int.TryParse(wert, out strafeDreifachHohl);
                    else if (label.Contains("doppelhohlstunde"))
                        int.TryParse(wert, out strafeDoppelHohl);
                    else if (label.Contains("hohlstunden"))
                        int.TryParse(wert, out strafeHohl);
                    else if (label.Contains("std.folge") || label.Contains("stdfolge"))
                        int.TryParse(wert, out strafeStdFolge);
                    else if (label.Contains("einzelstunde") || label.Contains("einzelstd"))
                        int.TryParse(wert, out strafeEinzel);
                    else if (label.Contains("späte lk") || label.Contains("lk stunden") || label.Contains("zuviele späte"))
                        int.TryParse(wert, out strafeSpäteLk);
                    else if (label.Contains("verbot doppelstunde") || label.Contains("verbot späte dopp"))
                        verbotSpäteDoppel = wert.Trim().ToLower() == "ja";
                    else if (label.Contains("hauptfach anteil") || label.Contains("hauptfach spät anteil"))
                        int.TryParse(wert, out hauptfachSpätAnteil);
                    else if (label.Contains("strafe hauptfach") || label.Contains("hauptfach strafe"))
                        int.TryParse(wert, out strafeHauptfachSpät);
                    else if (label.Contains("große pause") || label.Contains("grosse pause"))
                    {
                        // Format: "2-3" → stundeVor=2, stundeNach=3
                        var pausenTeile = wert.Split('-');
                        if (pausenTeile.Length == 2 &&
                            int.TryParse(pausenTeile[0].Trim(), out int pVor) &&
                            int.TryParse(pausenTeile[1].Trim(), out int pNach))
                            grossePausen.Add((pVor, pNach));
                    }
                }
            }

            // =====================================================
            // STAMMDATEN – HohlStd. soll + Std.Folge
            // =====================================================
            var lehrerStammdaten = new Dictionary<string, LehrerStammdaten>();

            if (workbook.Worksheets.Any(ws => ws.Name == "Stammdaten"))
            {
                var sheetSD = workbook.Worksheet("Stammdaten");
                var headerSD = GetHeaderMap(sheetSD);

                foreach (var row in sheetSD.RangeUsed().RowsUsed().Skip(1))
                {
                    string name = GetOptional(row, headerSD, "Name").Trim();
                    if (string.IsNullOrEmpty(name)) continue;

                    var sd = new LehrerStammdaten { Name = name };

                    // HohlStd. soll: "1-3" → min=1, max=3
                    string hohlRaw = GetOptional(row, headerSD, "HohlStd. soll").Trim();
                    if (!string.IsNullOrEmpty(hohlRaw))
                    {
                        var teile = hohlRaw.Split('-');
                        if (teile.Length == 2 &&
                            int.TryParse(teile[0].Trim(), out int hMin) &&
                            int.TryParse(teile[1].Trim(), out int hMax))
                        {
                            sd.HohlStdMin = hMin;
                            sd.HohlStdMax = hMax;
                        }
                    }

                    // Std.Folge: "6" → max 6 aufeinanderfolgende Stunden
                    string folgeRaw = GetOptional(row, headerSD, "Std.Folge").Trim();
                    if (!string.IsNullOrEmpty(folgeRaw) &&
                        int.TryParse(folgeRaw, out int folge))
                        sd.StdFolge = folge;

                    lehrerStammdaten[name] = sd;
                }
            }

            return new StundenplanInput
            {
                Blocks = unterrichtListe,
                Slots = zeitRaster,
                Fachraeume = fachgruppenRaeume,
                ExtraFreieTage = extraFreieTage,
                ExcelPfad = excelPfad,
                LehrerStammdaten = lehrerStammdaten,
                ZeitlimitSekunden = zeitlimit,
                AnzahlLösungenOhneTausch = anzahlOhne,
                AnzahlLösungenMitTausch = anzahlMit,
                NichtFreieTage = nichtFreieTage,
                GewichtFrüheDoppel = gewichtFrüh,
                GewichtSpäteDoppel = gewichtSpät,
                GewichtSpätePädEinheiten = gewichtPäd,
                GewichtFreieTage = gewichtFrei,
                StrafeHohlstunde = strafeHohl,
                StrafeDoppelHohlstunde = strafeDoppelHohl,
                StrafeDreifachHohlstunde = strafeDreifachHohl,
                StrafeStdFolge = strafeStdFolge,
                StrafeEinzelstunde = strafeEinzel,
                StrafeSpäteLkStunden = strafeSpäteLk,
                VerbotSpäteDoppel = verbotSpäteDoppel,
                HauptfachSpätAnteilProzent = hauptfachSpätAnteil,
                StrafeHauptfachSpät = strafeHauptfachSpät,
                GrossePausen = grossePausen,
            };
        }

        // =====================================================
        // ZEITWUNSCH-TABELLE
        // =====================================================

        private static void LeseZeitWunschTabelle(
            IXLWorksheet sheet,
            List<ZeitSlot> zeitRaster,
            bool istLehrer,
            Dictionary<string, int> extraFreieTage)
        {
            int row = 1;

            while (!sheet.Cell(row, 1).IsEmpty())
            {
                string name = sheet.Cell(row, 1).GetString().Trim();

                int extra = 0;

                var extraCell = sheet.Cell(row, 2);

                if (!extraCell.IsEmpty())
                    int.TryParse(extraCell.GetString(), out extra);

                if (istLehrer && extra > 0)
                {
                    if (!extraFreieTage.ContainsKey(name))
                        extraFreieTage[name] = extra;
                }

                row += 2;

                for (int stunde = 1; stunde <= 11; stunde++)
                {
                    for (int tag = 1; tag <= 5; tag++)
                    {
                        var cell = sheet.Cell(row, tag + 1);

                        if (!cell.IsEmpty())
                        {
                            int wert = cell.GetValue<int>();
                            string wtag = TagNummerZuString(tag);

                            var slot = zeitRaster
                                .FirstOrDefault(z =>
                                    z.WTag == wtag &&
                                    z.Stunde == stunde);

                            if (slot != null)
                            {
                                if (istLehrer)
                                    slot.LehrerWunsch[name] = wert;
                                else
                                    slot.KlassenWunsch[name] = wert;
                            }
                        }
                    }

                    row++;
                }

                row += 2;
            }
        }

        private static string TagNummerZuString(int tag)
        {
            return tag switch
            {
                1 => "Mo",
                2 => "Di",
                3 => "Mi",
                4 => "Do",
                5 => "Fr",
                _ => ""
            };
        }

        private static string BestimmeFachgruppe(string fach)
        {
            if (string.IsNullOrWhiteSpace(fach))
                return "";

            if (fach.StartsWith("BI", StringComparison.OrdinalIgnoreCase))
                return "Bio";
            if (fach.StartsWith("Sp", StringComparison.OrdinalIgnoreCase))
                return "Sport";
            if (fach.StartsWith("Ch", StringComparison.OrdinalIgnoreCase))
                return "Chemie";
            if (fach.StartsWith("Ph", StringComparison.OrdinalIgnoreCase))
                return "Physik";
            if (fach.StartsWith("Mu", StringComparison.OrdinalIgnoreCase))
                return "Musik";
            if (fach.StartsWith("Ku", StringComparison.OrdinalIgnoreCase))
                return "Kunst";
            if (fach.StartsWith("IF", StringComparison.OrdinalIgnoreCase))
                return "Informatik";

            return "Sonstige";
        }
        private static string GetOptional(IXLRangeRow row, Dictionary<string, int> map, string name)
        {
            return map.ContainsKey(name)
                ? row.Cell(map[name]).GetString()
                : "";
        }
        private static Dictionary<string, int> GetHeaderMap(IXLWorksheet sheet)
        {
            var headerRow = sheet.Row(1);

            return headerRow.CellsUsed()
                .ToDictionary(
                    c => c.GetString().Trim(),
                    c => c.Address.ColumnNumber,
                    StringComparer.OrdinalIgnoreCase);
        }
        private static IXLCell Cell(IXLRangeRow row, Dictionary<string, int> map, string name)
        {
            if (!map.ContainsKey(name))
                throw new Exception($"Spalte '{name}' nicht gefunden.");

            return row.Cell(map[name]);
        }
    }
}
