using System.Collections.Generic;

namespace Stundenplan_V2
{
    public class UnterrichtsBlock
    {
        public int UNr { get; set; }
        public int Wst { get; set; }
        public string Zeilentext { get; set; } = "";
        public int WochenDoppelstunden { get; set; }
        public bool DoppelÜberPauseErlaubt { get; set; } = false; // (E)-Spalte: x = erlaubt

        public Dictionary<string, int> TagesDoppelstunden { get; set; } = new();
        public List<TeilUnterricht> Teile { get; set; } = new();
    }
    public class TeilUnterricht
    {
        public int UNr { get; set; }
        public string Lehrer { get; set; } = "";
        public string Fach { get; set; } = "";
        public List<string> Klassen { get; set; } = new();
        public int MinDoppel { get; set; }
        public int MaxDoppel { get; set; }
        public string FachGruppe { get; set; }
        public int AktuelleDoppelstunden { get; set; }
        public string Ltkz { get; set; } = "";
        public bool DoppelÜberPauseErlaubt { get; set; } = false; // (E)-Spalte
    }

 
}