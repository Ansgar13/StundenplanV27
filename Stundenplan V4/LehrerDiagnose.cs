using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Stundenplan_V2
{
    /// <summary>
    /// Diagnose-Ergebnis für einen Lehrer in einer Lösung
    /// </summary>
    public class LehrerDiagnoseErgebnis
    {
        public string Lehrer { get; set; }
        public int HohlstundenGesamt { get; set; }
        public int DoppelHohlstunden { get; set; }   // 2 aufeinanderfolgende Hohlstunden
        public int DreifachHohlstunden { get; set; } // 3+ aufeinanderfolgende Hohlstunden
        public int MaxStdFolge { get; set; }          // größte aufeinanderfolgende Unterrichtsfolge
        public int Einzelstunden { get; set; }        // Tage mit genau 1 Unterrichtsstunde
        public int StrafeGesamt { get; set; }

        // Vorgaben aus Stammdaten
        public int? HohlStdSollMin { get; set; }
        public int? HohlStdSollMax { get; set; }
        public int? StdFolgeMax { get; set; }

        // Auffälligkeiten
        public bool HohlstundenZuViel => HohlStdSollMax.HasValue && HohlstundenGesamt > HohlStdSollMax;
        public bool HohlstundenZuWenig => HohlStdSollMin.HasValue && HohlstundenGesamt < HohlStdSollMin;
        public bool StdFolgeÜberschritten => StdFolgeMax.HasValue && MaxStdFolge > StdFolgeMax;
    }

    public static class LehrerDiagnose
    {
        /// <summary>
        /// Berechnet die Diagnose aller Lehrer für eine Lösung
        /// </summary>
        public static List<LehrerDiagnoseErgebnis> Berechne(
            int[,] belegung,
            List<UnterrichtsBlock> blocks,
            List<ZeitSlot> slots,
            Dictionary<string, LehrerStammdaten> stammdaten,
            int strafeHohl,
            int strafeDoppelHohl,
            int strafeDreifachHohl,
            int strafeStdFolge)
        {
            // Alle Lehrer ermitteln
            var alleLehrern = blocks
                .SelectMany(b => b.Teile.Select(t => t.Lehrer))
                .Distinct().OrderBy(l => l).ToList();

            // Tage und Stunden strukturieren
            var tage = slots.Select(s => s.WTag).Distinct().ToList();

            var result = new List<LehrerDiagnoseErgebnis>();

            foreach (var lehrer in alleLehrern)
            {
                var diag = new LehrerDiagnoseErgebnis { Lehrer = lehrer };

                // Stammdaten zuordnen
                if (stammdaten.TryGetValue(lehrer, out var sd))
                {
                    diag.HohlStdSollMin = sd.HohlStdMin;
                    diag.HohlStdSollMax = sd.HohlStdMax;
                    diag.StdFolgeMax    = sd.StdFolge;
                }

                // Für jeden Tag: Stunden-Sequenz des Lehrers aufbauen
                foreach (var tag in tage)
                {
                    // Slots dieses Tages sortiert nach Stunde
                    var tagesSlots = slots
                        .Select((s, i) => (s, i))
                        .Where(x => x.s.WTag == tag)
                        .OrderBy(x => x.s.Stunde)
                        .ToList();

                    if (tagesSlots.Count == 0) continue;

                    int minStunde = tagesSlots.First().s.Stunde;
                    int maxStunde = tagesSlots.Last().s.Stunde;

                    // Für jede Stunde: hat der Lehrer Unterricht?
                    var stundenMitUnterricht = new HashSet<int>();
                    foreach (var (slot, sIdx) in tagesSlots)
                    {
                        for (int b = 0; b < blocks.Count; b++)
                        {
                            if (belegung[b, sIdx] == 1 &&
                                blocks[b].Teile.Any(t => t.Lehrer == lehrer))
                            {
                                stundenMitUnterricht.Add(slot.Stunde);
                                break;
                            }
                        }
                    }

                    if (stundenMitUnterricht.Count == 0) continue;

                    int ersteStunde = stundenMitUnterricht.Min();
                    int letzteStunde = stundenMitUnterricht.Max();

                    // Hohlstunden: Stunden zwischen erster und letzter Unterrichtsstunde ohne Unterricht
                    for (int std = ersteStunde + 1; std < letzteStunde; std++)
                    {
                        if (!stundenMitUnterricht.Contains(std))
                            diag.HohlstundenGesamt++;
                    }

                    // Doppel- und Dreifach-Hohlstunden: aufeinanderfolgende Hohlstunden
                    int hohlFolge = 0;
                    for (int std = ersteStunde + 1; std <= letzteStunde; std++)
                    {
                        if (!stundenMitUnterricht.Contains(std) && std < letzteStunde)
                            hohlFolge++;
                        else
                        {
                            if (hohlFolge >= 3) diag.DreifachHohlstunden++;
                            else if (hohlFolge == 2) diag.DoppelHohlstunden++;
                            hohlFolge = 0;
                        }
                    }

                    // Einzelstunden: genau 1 Unterrichtsstunde am Tag
                    if (stundenMitUnterricht.Count == 1)
                        diag.Einzelstunden++;

                    // Std.Folge: längste aufeinanderfolgende Unterrichtssequenz
                    int aktFolge = 0;
                    int maxFolge = 0;
                    for (int std = ersteStunde; std <= letzteStunde; std++)
                    {
                        if (stundenMitUnterricht.Contains(std))
                        {
                            aktFolge++;
                            maxFolge = Math.Max(maxFolge, aktFolge);
                        }
                        else
                            aktFolge = 0;
                    }
                    diag.MaxStdFolge = Math.Max(diag.MaxStdFolge, maxFolge);
                }

                // Strafe berechnen
                diag.StrafeGesamt =
                    diag.HohlstundenGesamt  * strafeHohl +
                    diag.DoppelHohlstunden  * strafeDoppelHohl +
                    diag.DreifachHohlstunden * strafeDreifachHohl +
                    (diag.StdFolgeÜberschritten ? strafeStdFolge : 0);

                result.Add(diag);
            }

            return result;
        }

        /// <summary>
        /// Exportiert die Diagnosetabelle für alle Lösungen in ein Excel-Sheet
        /// </summary>
        public static void Exportiere(
            string excelPfad,
            List<(string label, List<LehrerDiagnoseErgebnis> diagnosen)> lösungen)
        {
            using var wb = new XLWorkbook(excelPfad);

            const string sheetName = "Diagnose";
            if (wb.Worksheets.Any(ws => ws.Name == sheetName))
                wb.Worksheet(sheetName).Delete();

            var sheet = wb.Worksheets.Add(sheetName);

            // Header-Zeile 1: Lösungs-Label (gemergte Zellen)
            int startCol = 2;
            int colsProLösung = 8;

            sheet.Cell(1, 1).Value = "Lehrer";
            sheet.Cell(1, 1).Style.Font.Bold = true;

            for (int i = 0; i < lösungen.Count; i++)
            {
                int col = startCol + i * colsProLösung;
                sheet.Cell(1, col).Value = lösungen[i].label;
                sheet.Cell(1, col).Style.Font.Bold = true;
                sheet.Cell(1, col).Style.Fill.BackgroundColor = XLColor.LightBlue;
                sheet.Range(1, col, 1, col + colsProLösung - 1).Merge();
            }

            // Header-Zeile 2: Spaltenbezeichnungen
            sheet.Cell(2, 1).Value = "Lehrer";
            sheet.Cell(2, 1).Style.Font.Bold = true;

            for (int i = 0; i < lösungen.Count; i++)
            {
                int col = startCol + i * colsProLösung;
                sheet.Cell(2, col    ).Value = "Hohlstd.";
                sheet.Cell(2, col + 1).Value = "Soll min";
                sheet.Cell(2, col + 2).Value = "Soll max";
                sheet.Cell(2, col + 3).Value = "DoppelHohl";
                sheet.Cell(2, col + 4).Value = "DreiHohl";
                sheet.Cell(2, col + 5).Value = "Max Folge";
                sheet.Cell(2, col + 6).Value = "Folge max";
                sheet.Cell(2, col + 7).Value = "Einzelstd.";

                for (int c = col; c < col + colsProLösung; c++)
                {
                    sheet.Cell(2, c).Style.Font.Bold = true;
                    sheet.Cell(2, c).Style.Fill.BackgroundColor = XLColor.LightGray;
                }
            }

            // Alle Lehrer aus erster Lösung
            var alleLehrern = lösungen.Count > 0
                ? lösungen[0].diagnosen.Select(d => d.Lehrer).ToList()
                : new List<string>();

            // Daten
            for (int lIdx = 0; lIdx < alleLehrern.Count; lIdx++)
            {
                string lehrer = alleLehrern[lIdx];
                int zeile = lIdx + 3;

                sheet.Cell(zeile, 1).Value = lehrer;

                for (int i = 0; i < lösungen.Count; i++)
                {
                    int col = startCol + i * colsProLösung;
                    var d = lösungen[i].diagnosen.FirstOrDefault(x => x.Lehrer == lehrer);
                    if (d == null) continue;

                    sheet.Cell(zeile, col    ).Value = d.HohlstundenGesamt;
                    sheet.Cell(zeile, col + 1).Value = d.HohlStdSollMin?.ToString() ?? "–";
                    sheet.Cell(zeile, col + 2).Value = d.HohlStdSollMax?.ToString() ?? "–";
                    sheet.Cell(zeile, col + 3).Value = d.DoppelHohlstunden;
                    sheet.Cell(zeile, col + 4).Value = d.DreifachHohlstunden;
                    sheet.Cell(zeile, col + 5).Value = d.MaxStdFolge;
                    sheet.Cell(zeile, col + 6).Value = d.StdFolgeMax?.ToString() ?? "–";
                    sheet.Cell(zeile, col + 7).Value = d.Einzelstunden;

                    // Auffälligkeiten rot markieren
                    if (d.HohlstundenZuViel || d.HohlstundenZuWenig)
                        sheet.Cell(zeile, col).Style.Fill.BackgroundColor = XLColor.LightPink;
                    if (d.DoppelHohlstunden > 0)
                        sheet.Cell(zeile, col + 3).Style.Fill.BackgroundColor = XLColor.LightPink;
                    if (d.DreifachHohlstunden > 0)
                        sheet.Cell(zeile, col + 4).Style.Fill.BackgroundColor = XLColor.LightPink;
                    if (d.StdFolgeÜberschritten)
                        sheet.Cell(zeile, col + 5).Style.Fill.BackgroundColor = XLColor.LightPink;
                    if (d.Einzelstunden > 0)
                        sheet.Cell(zeile, col + 7).Style.Fill.BackgroundColor = XLColor.LightPink;
                }
            }

            // Summenzeile direkt unter den Daten (nach letzter Lehrer-Zeile)
            int letzteDataZeile = alleLehrern.Count + 2;
            int sumZeile = letzteDataZeile + 1;

            sheet.Cell(sumZeile, 1).Value = "Summe";
            sheet.Cell(sumZeile, 1).Style.Font.Bold = true;
            sheet.Cell(sumZeile, 1).Style.Fill.BackgroundColor = XLColor.LightGray;

            for (int i = 0; i < lösungen.Count; i++)
            {
                int col = startCol + i * colsProLösung;
                var diags = lösungen[i].diagnosen;

                // Summe Hohlstunden
                int sumHohl = diags.Sum(d => d.HohlstundenGesamt);
                sheet.Cell(sumZeile, col).Value = sumHohl;
                sheet.Cell(sumZeile, col).Style.Font.Bold = true;
                sheet.Cell(sumZeile, col).Style.Fill.BackgroundColor = XLColor.LightGray;

                // Soll-Spalten leer lassen
                sheet.Cell(sumZeile, col + 1).Value = "";
                sheet.Cell(sumZeile, col + 2).Value = "";

                // Summe DoppelHohlstunden
                int sumDoppel = diags.Sum(d => d.DoppelHohlstunden);
                sheet.Cell(sumZeile, col + 3).Value = sumDoppel;
                sheet.Cell(sumZeile, col + 3).Style.Font.Bold = true;
                sheet.Cell(sumZeile, col + 3).Style.Fill.BackgroundColor =
                    sumDoppel > 0 ? XLColor.LightPink : XLColor.LightGray;

                // Summe DreifachHohlstunden
                int sumDrei = diags.Sum(d => d.DreifachHohlstunden);
                sheet.Cell(sumZeile, col + 4).Value = sumDrei;
                sheet.Cell(sumZeile, col + 4).Style.Font.Bold = true;
                sheet.Cell(sumZeile, col + 4).Style.Fill.BackgroundColor =
                    sumDrei > 0 ? XLColor.LightPink : XLColor.LightGray;

                // Summe Einzelstunden
                int sumEinzel = diags.Sum(d => d.Einzelstunden);
                sheet.Cell(sumZeile, col + 7).Value = sumEinzel;
                sheet.Cell(sumZeile, col + 7).Style.Font.Bold = true;
                sheet.Cell(sumZeile, col + 7).Style.Fill.BackgroundColor =
                    sumEinzel > 0 ? XLColor.LightPink : XLColor.LightGray;
            }

            sheet.Columns().AdjustToContents();
            wb.Save();
        }
    }
}
