Attribute VB_Name = "Unterrichtsoptimierung"
Option Explicit
' ============================================================
' UNTERRICHTSVERTEILUNG v9.22e
' Aenderungen v9.22e: BT-Endlosschleife behoben: leere Domain wird als SKIP markiert, nicht wiederholt versucht
' Aenderungen v9.22d: stKand/stScor von 60 auf 100 erweitert (PufferÃ¼berlauf bei >60 Lehrern behoben)
'                     BT: Fix-Eintraege werden vor SchreibeLoesungen immer in loesung1/2 eingetragen
' Aenderungen v9.22c: InitDomains: Fallback ohne Kapazitaetspruefung wenn Domain leer (BT findet immer Loesung)
'                     BT: UrsprungsLehrer-Reset nach EintraegeEinlesen entfernt (Fix-Spalte wird jetzt beachtet)
' Aenderungen v9.22a: Chr(8594) durch "->" ersetzt
' Aenderungen v9.22: g_stDom auf 2000 erweitert, Nachoptimierung prueft alle Lehrer
' Aenderungen v9.20: MAXDEPTH dynamisch auf nE, wertVal-Reset, ZuweisungZuruecksetzen loescht Wert-Spalte
' Aenderungen v9.19: BT-Fixes: Timer oben in Schleife, UrsprungsLehrer=? vor BT
' Aenderungen v9.18: ZuweisungZuruecksetzen leert auch Fix-Spalte
' Aenderungen v9.17: Brute-Force Lastausgleich in VerbessereSA
' Aenderungen v9.16: VerbessereSA ersetzt durch Greedy-Lastausgleich
' Aenderungen v9.15: temp-Kuehlung bei jedem Schleifendurchlauf (auch bei GoTo WeiterSA)
' Aenderungen v9.14: BerechneScore gegen Ueberlauf abgesichert
' Aenderungen v9.13: Bedarf in Engpassanalyse = nur noch nicht fixierte Werte
' Aenderungen v9.12: Engpassanalyse liest Bedarf+IstWSt aus Klassen Sp.9+Sp.6
' Aenderungen v9.11: Verbesserung L3+L4 laeuft automatisch ohne Unterbrechung
' Aenderungen v9.10: Fallback WertUV=WSt bei Wert=0 entfernt
' Aenderungen v9.9: Import schreibt Wst in Sp.3 (nicht Wert=)
' Aenderungen v9.8: Konsistenz-Nachbearbeitung in VerbessereSA (wie SA_Solve),
'                   Op4 gezielter Lastausgleich, Button-Fixes
' Aenderungen v9.7: Konsistenzpruefung Klasse+Fach in VerbessereSA,
'                   Spaltensummen-Makro fuer Klassen/KlassenUV/Lehrerliste
' Aenderungen v9.6: Header-Erkennung fuer "Wert =" (mit Leerzeichen) ergaenzt
' Aenderungen v9.5: Wert-Spalte (Sp.9) im Import, IstWSt auf WertUV,
'                   curScore alle 500 Iter. neu berechnet (Overflow-Schutz)
' Aenderungen v9.4: Alle Exp()-Ueberlauf-Stellen abgesichert (VerbessereSA + SA_Solve)
' Aenderungen v9.2: Doppelzaehlung IstWSt behoben, Spaltenbezeichnungen
'                   Soll/Ist getauscht, Ueberlast nur tatsaechlich unterrichtete Faecher
' ============================================================

' ============================================================
' UNTERRICHTSVERTEILUNG v8
' Backtracking + Constraint Propagation  ODER  OR-Tools
' Erzeugt 2 Loesungen + Diagnose1/2 + Lehrerbelegung
' ============================================================

Const TOLERANZ        As Double = 2
Const SHEET_DIAG1     As String = "Diagnose1"
Const SHEET_DIAG2     As String = "Diagnose2"
Const SHEET_BELEGUNG    As String = "Lehrerbelegung"
Const SHEET_FACHGRUPPEN As String = "FachgruppenLehrer"
Const SHEET_FGDEF       As String = "Fachgruppen"
Const SHEET_DIAG3       As String = "Diagnose3"
Const SHEET_DIAG4       As String = "Diagnose4"
Const SOLVER_EXE_NAME As String = "stundenplan_solver.exe"

Dim g_lehrer()   As tLehrer
Dim g_nL         As Long
Dim g_wuensche() As tWunsch
Dim g_nW         As Long
Dim g_toleranz   As Double
Dim g_sperrIdx   As Long
Dim g_sperrName  As String
Dim g_btStartTime As Double  ' Startzeit Backtracking fuer Zeitlimit
Dim g_btMaxSek    As Double  ' Max. Sekunden (0 = kein Limit)
Dim g_zeitlimit   As Double  ' Zeitlimit BT-Phase in Sekunden (aus Parameter-Sheet)

' Score-Parameter (werden aus Tabelle "Parameter" gelesen)
Dim g_scoreKL     As Double   ' Klassenleitung
Dim g_scoreW3     As Double   ' Wunsch Prio 3
Dim g_scoreW2     As Double   ' Wunsch Prio 2
Dim g_scoreW1     As Double   ' Wunsch Prio 1
Dim g_scoreA2     As Double   ' Anti-Wunsch Prio 2
Dim g_scoreA1     As Double   ' Anti-Wunsch Prio 1
Dim g_scoreKont   As Double   ' Kontinuität
Dim g_scoreFreiF  As Double   ' Faktor freie Stunden
Dim g_scoreUeberF As Double   ' Faktor Überlastung (wird negativ angewendet)
Dim g_scoreUnterF As Double   ' Faktor Unterbesetzung (Malus)
Dim g_schutzFG1   As String   ' Schutzfachgruppe 1 (leer = inaktiv)
Dim g_schutzFG2   As String   ' Schutzfachgruppe 2 (leer = inaktiv)
Dim g_schutzMalus As Double   ' Malus-Faktor fuer Schutzfachgruppen (Default 20)

' Backtracking-Stack (global um Stack-Overflow zu vermeiden)
Dim g_stDom(1 To 2000, 1 To 2000) As String
Dim g_fgFach(1 To 500)   As String
Dim g_fgGruppe(1 To 500) As String
Dim g_fgAnz As Long

Type tLehrer
    name              As String
    fach(1 To 16)     As String
    OberstufenOK(1 To 16) As Boolean
    KlassenleiterIn   As String
    sollWst           As Double
    istWSt            As Double
    IstSchutzLehrer   As Boolean  ' True wenn Lehrer zu Schutzfachgruppe 1 oder 2 gehoert
End Type

Type tEintrag
    klasse            As String
    fach              As String
    WSt               As Double   ' Wochenstunden
    WertUV            As Double   ' Gewichteter Wert aus KlassenUV
    lehrer            As String
    UrsprungsLehrer   As String
    zeile             As Long
    IstOberstufe      As Boolean
End Type

Type tWunsch
    lehrerName        As String
    WunschKlasse      As String
    WunschPrio        As Long
    AntiKlasse        As String
    AntiPrio          As Long
End Type

' ============================================================
' HAUPTMAKRO
' ============================================================
Sub UnterrichtsverteilungOptimieren()
    Dim methode As Long
    Dim tolVal  As Double

    methode = MethodenwahlDialog()
    If methode = 0 Then Exit Sub

    ' Parameter aus Tabelle lesen
    Call ParameterEinlesen
    tolVal = g_toleranz

    If methode = 2 Then
        Call Optimieren_SA
    Else
        Call Optimieren_Backtracking
    End If
End Sub

Function MethodenwahlDialog() As Long
    Dim msg As String
    Dim antwort As VbMsgBoxResult
    msg = "OPTIMIERUNGSMETHODE WAEHLEN" & vbCrLf & vbCrLf
    msg = msg & "1)  Backtracking + Constraint Propagation" & vbCrLf
    msg = msg & "    - Systematische Suche, sehr genau" & vbCrLf
    msg = msg & "    - Kann bei grossen Daten langsamer sein" & vbCrLf & vbCrLf
    msg = msg & "2)  Simulated Annealing (empfohlen)" & vbCrLf
    msg = msg & "    - Laeuft direkt in Excel, kein Python noetig" & vbCrLf
    msg = msg & "    - Schnell auch bei vielen Eintraegen" & vbCrLf
    msg = msg & "    - Findet immer eine Loesung" & vbCrLf & vbCrLf
    msg = msg & "Ja = Backtracking+CP   Nein = Simulated Annealing   Abbrechen = Abbruch"
    antwort = MsgBox(msg, vbYesNoCancel + vbQuestion, "Optimierungsmethode")
    Select Case antwort
        Case vbYes:    MethodenwahlDialog = 1
        Case vbNo:     MethodenwahlDialog = 2
        Case vbCancel: MethodenwahlDialog = 0
    End Select
End Function

' ============================================================
' METHODE 1: BACKTRACKING + CONSTRAINT PROPAGATION
' ============================================================
Sub Optimieren_Backtracking()
    Dim wsL As Worksheet
    Dim wsK As Worksheet
    Dim wsW As Worksheet
    Dim eintraege() As tEintrag
    Dim nE As Long
    Dim domains() As String
    Dim loesung1() As String
    Dim loesung2() As String
    Dim istWSt1() As Double
    Dim istWSt2() As Double
    Dim ok1 As Boolean
    Dim ok2 As Boolean
    Dim i As Long
    Dim e As Long
    Dim e2 As Long
    Dim tolVal As Double

    On Error GoTo Fehlerbehandlung
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wsL = SheetByName("Lehrerliste")
    Set wsK = SheetByName("Klassen")
    Set wsW = SheetByName("Lehrerw" & Chr(252) & "nsche")
    If wsL Is Nothing Or wsK Is Nothing Or wsW Is Nothing Then
        MsgBox "Benoetigt: Lehrerliste, Klassen, Lehrerw" & Chr(252) & "nsche", vbCritical
        GoTo Cleanup
    End If

    ReDim g_lehrer(1 To 1)
    g_nL = LehrerEinlesen(wsL, g_lehrer)
    Call SchutzFlagsSetzen
    If g_nL = 0 Then MsgBox "Keine Lehrer!", vbCritical: GoTo Cleanup

    ReDim g_wuensche(1 To 1)
    g_nW = WuenscheEinlesen(wsW, g_wuensche)

    nE = EintraegeEinlesen(wsK, eintraege)
    If nE = 0 Then MsgBox "Keine Eintraege!", vbCritical: GoTo Cleanup

    ' Sicherheitspruefung fuer g_stDom Array-Grenze (1 To 2000)
    If nE > 2000 Then
        MsgBox "Zu viele Eintraege (" & nE & "). Maximum: 2000." & vbCrLf & _
               "Bitte Methode 2 (Simulated Annealing) verwenden.", vbCritical
        GoTo Cleanup
    End If




    Call ResetIstWSt(eintraege, nE)

    ReDim domains(1 To nE)
    ReDim loesung1(1 To nE)
    ReDim istWSt1(1 To g_nL)
    For e = 1 To nE: loesung1(e) = "?": Next e

    tolVal = g_toleranz   ' Vom Dialog gesetzten Wert merken
    ' g_toleranz wurde bereits im Hauptdialog gesetzt
    ' Prüfen ob für jeden Eintrag mindestens ein Lehrer qualifiziert ist
    Dim pruefMsg As String
    Dim pruefFehler As Long
    pruefMsg = ""
    pruefFehler = 0
    Dim pruefE As Long
    Dim pruefI As Long
    Dim pruefOK As Boolean
    For pruefE = 1 To nE
        If IstFixiert(eintraege(pruefE).lehrer) Then GoTo NaechsterPruef
        pruefOK = False
        For pruefI = 1 To g_nL
            If KannFach(g_lehrer(pruefI), eintraege(pruefE).fach, eintraege(pruefE).IstOberstufe) Then
                pruefOK = True: Exit For
            End If
        Next pruefI
        If Not pruefOK Then
            pruefFehler = pruefFehler + 1
            If pruefFehler <= 10 Then
                pruefMsg = pruefMsg & "  - " & eintraege(pruefE).fach & " in " & eintraege(pruefE).klasse & vbCrLf
            End If
        End If
NaechsterPruef:
    Next pruefE
    If pruefFehler > 0 Then
        Dim warnMsg As String
        warnMsg = "WARNUNG: F" & Chr(252) & "r " & pruefFehler & " Eintrag(e) gibt es keinen qualifizierten Lehrer:" & vbCrLf & vbCrLf
        warnMsg = warnMsg & pruefMsg
        If pruefFehler > 10 Then warnMsg = warnMsg & "  ... und " & (pruefFehler - 10) & " weitere." & vbCrLf
        warnMsg = warnMsg & vbCrLf & "Diese Eintr" & Chr(228) & "ge bleiben unbesetzt (Lehrer = ?)." & vbCrLf & vbCrLf
        warnMsg = warnMsg & "Trotzdem fortfahren?"
        If MsgBox(warnMsg, vbYesNo + vbExclamation, "Fehlende Lehrer") = vbNo Then
            GoTo Cleanup
        End If
    End If

    g_sperrIdx = 0
    g_sperrName = ""
    g_btMaxSek = g_zeitlimit
    Call InitDomains(eintraege, nE, domains)
    g_btStartTime = Timer  ' NACH InitDomains: Init-Zeit nicht vom Limit abziehen
    ok1 = BacktrackingMitCP(eintraege, nE, domains, loesung1)
    If Not ok1 Then
        g_toleranz = 9999
        g_btStartTime = Timer
        For e = 1 To nE: loesung1(e) = "?": Next e
        Call ResetIstWSt(eintraege, nE)
        Call InitDomains(eintraege, nE, domains)
        ok1 = BacktrackingMitCP(eintraege, nE, domains, loesung1)
    End If
    g_btMaxSek = 0
    g_toleranz = tolVal
    If Not ok1 Then
        MsgBox "Backtracking konnte keine vollst" & Chr(228) & "ndige L" & Chr(246) & "sung finden." & vbCrLf & _
               "Zeitlimit (" & CStr(CLng(g_zeitlimit)) & " Sek.) wurde ggf. erreicht." & vbCrLf & _
               "Einige Eintr" & Chr(228) & "ge bleiben unbesetzt." & vbCrLf & vbCrLf & _
               "Tipp: Zeitlimit im Parameter-Sheet erhoehen.", vbExclamation, "Backtracking"
    End If
    DoEvents
    For i = 1 To g_nL
        istWSt1(i) = g_lehrer(i).istWSt
    Next i

    ' Sperr-Eintrag fuer Loesung2: ersten nicht-fixierten Eintrag MIT zugewiesenem Lehrer
    g_sperrIdx = 0
    g_sperrName = ""
    For e2 = 1 To nE
        If Not IstFixiert(eintraege(e2).UrsprungsLehrer) Then
            If loesung1(e2) <> "" And loesung1(e2) <> "?" Then
                g_sperrIdx = e2
                g_sperrName = loesung1(e2)
                Exit For
            End If
        End If
    Next e2

    ReDim loesung2(1 To nE)
    ReDim istWSt2(1 To g_nL)

    ' Wenn alle Eintraege fixiert oder Lauf1 fehlgeschlagen: Loesung 2 = Loesung 1
    If g_sperrIdx = 0 Or Not ok1 Then
        For e2 = 1 To nE: loesung2(e2) = loesung1(e2): Next e2
        For i = 1 To g_nL: istWSt2(i) = istWSt1(i): Next i
    Else
        Call ResetIstWSt(eintraege, nE)
        g_toleranz = tolVal
        g_btMaxSek = g_zeitlimit
        Call InitDomains(eintraege, nE, domains)
        g_btStartTime = Timer
        ok2 = BacktrackingMitCP(eintraege, nE, domains, loesung2)
        If Not ok2 Then
            g_toleranz = 9999
            g_btStartTime = Timer
            Call ResetIstWSt(eintraege, nE)
            Call InitDomains(eintraege, nE, domains)
            ok2 = BacktrackingMitCP(eintraege, nE, domains, loesung2)
        End If
        g_btMaxSek = 0
        g_toleranz = tolVal
        DoEvents
        For i = 1 To g_nL
            istWSt2(i) = g_lehrer(i).istWSt
        Next i
    End If

    ' Fix-Eintraege immer in loesung1/2 eintragen (unabhaengig vom BT-Ergebnis)
    For e = 1 To nE
        If IstFixiert(eintraege(e).UrsprungsLehrer) Then
            loesung1(e) = eintraege(e).UrsprungsLehrer
            loesung2(e) = eintraege(e).UrsprungsLehrer
        End If
    Next e

    Call SchreibeLoesungen(wsK, eintraege, nE, loesung1, loesung2)
    Call BaueErgebnisSheets(loesung1, loesung2, istWSt1, istWSt2, eintraege, nE, "Backtracking+CP")

Cleanup:
    g_btMaxSek = 0
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

Fehlerbehandlung:
    g_btMaxSek = 0
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Fehler " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Tipp: Zeitlimit im Parameter-Sheet erhoehen oder Methode 2 (SA) verwenden.", _
           vbExclamation, "Optimierung abgebrochen"
End Sub

' ============================================================
' METHODE 2: OR-TOOLS
' ============================================================
Sub Debug_SucheExe()
    Dim pfad As String
    Dim fso As Object
    Dim msg As String

    msg = "ThisWorkbook.Path: [" & ThisWorkbook.Path & "]" & vbCrLf
    msg = msg & "ThisWorkbook.FullName: [" & ThisWorkbook.FullName & "]" & vbCrLf
    msg = msg & "ActiveWorkbook.Path: [" & ActiveWorkbook.Path & "]" & vbCrLf
    msg = msg & "CurDir: [" & CurDir & "]" & vbCrLf & vbCrLf

    pfad = ThisWorkbook.Path & "\" & SOLVER_EXE_NAME
    msg = msg & "Suche EXE unter: [" & pfad & "]" & vbCrLf & vbCrLf

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        msg = msg & "FileExists: " & fso.FileExists(pfad) & vbCrLf
        Set fso = Nothing
    End If
    On Error GoTo 0

    msg = msg & vbCrLf & "SucheExe() gibt zurueck: [" & SucheExe() & "]"
    MsgBox msg, vbInformation, "EXE-Suche Debug"
End Sub

Sub Optimieren_ORTools()
    Dim exePfad As String
    Dim excelPfad As String
    Dim shell As Object
    Dim cmd As String
    Dim exitCode As Long
    Dim setupMsg As String
    Dim wsL As Worksheet
    Dim wsK As Worksheet
    Dim wsW As Worksheet
    Dim eintraege() As tEintrag
    Dim nE As Long
    Dim loesung1() As String
    Dim loesung2() As String
    Dim istWSt1() As Double
    Dim istWSt2() As Double
    Dim i As Long
    Dim e As Long

    exePfad = SucheExe()
    If exePfad = "" Then
        setupMsg = "stundenplan_solver.exe nicht gefunden!" & vbCrLf & vbCrLf
        setupMsg = setupMsg & "So einrichten:" & vbCrLf
        setupMsg = setupMsg & "1) Python installieren (python.org)" & vbCrLf
        setupMsg = setupMsg & "2) build_solver_exe.bat ausfuehren" & vbCrLf
        setupMsg = setupMsg & "3) EXE in selben Ordner wie Excel legen" & vbCrLf & vbCrLf
        setupMsg = setupMsg & "Stattdessen Backtracking+CP verwenden?"
        If MsgBox(setupMsg, vbYesNo + vbQuestion) = vbYes Then
            Call Optimieren_Backtracking
        End If
        Exit Sub
    End If

    ' Excel-Pfad ermitteln - OneDrive-kompatibel (FullName gibt URL zurueck)
    excelPfad = LokalisierePfad(ThisWorkbook.FullName)
    If excelPfad = "" Or left(excelPfad, 4) = "http" Then
        MsgBox "Lokalen Dateipfad konnte nicht ermittelt werden." & vbCrLf & _
               "Bitte Excel-Datei lokal speichern (nicht nur OneDrive).", vbExclamation
        Exit Sub
    End If

    ThisWorkbook.Save
    Set shell = CreateObject("WScript.Shell")
    ' Logdatei im selben Ordner wie EXE
    Dim logPfad As String
    logPfad = left(exePfad, InStrRev(exePfad, "")) & "solver_log.txt"
    ' EXE mit Ausgabe in Logdatei starten
    cmd = "cmd /C """ & exePfad & """ """ & excelPfad & """ > """ & logPfad & """ 2>&1"
    Application.StatusBar = "OR-Tools rechnet..."
    DoEvents
    exitCode = shell.Run(cmd, 0, True)
    Application.StatusBar = False

    If exitCode <> 0 Then
        ' Logdatei lesen und anzeigen
        Dim logMsg As String
        logMsg = ""
        On Error Resume Next
        Dim iFile As Integer: iFile = FreeFile
        Open logPfad For Input As #iFile
        Dim logLine As String
        Dim lineCount As Long: lineCount = 0
        Do While Not EOF(iFile) And lineCount < 20
            Line Input #iFile, logLine
            logMsg = logMsg & logLine & vbCrLf
            lineCount = lineCount + 1
        Loop
        Close #iFile
        On Error GoTo 0
        Dim errMsg As String
        errMsg = "OR-Tools Fehler (Code " & exitCode & ")." & vbCrLf & vbCrLf
        If logMsg <> "" Then errMsg = errMsg & "Fehlermeldung:" & vbCrLf & logMsg & vbCrLf
        errMsg = errMsg & "Backtracking+CP verwenden?"
        If MsgBox(errMsg, vbYesNo + vbExclamation, "OR-Tools Fehler") = vbYes Then
            Call Optimieren_Backtracking
        End If
        Exit Sub
    End If

    MsgBox "OR-Tools abgeschlossen. Datei wird neu geladen.", vbInformation
    ThisWorkbook.Close SaveChanges:=False
    Workbooks.Open excelPfad

    Set wsL = SheetByName("Lehrerliste")
    Set wsK = SheetByName("Klassen")
    Set wsW = SheetByName("Lehrerw" & Chr(252) & "nsche")

    ReDim g_lehrer(1 To 1)
    g_nL = LehrerEinlesen(wsL, g_lehrer)
    Call SchutzFlagsSetzen
    ReDim g_wuensche(1 To 1)
    g_nW = WuenscheEinlesen(wsW, g_wuensche)
    nE = EintraegeEinlesen(wsK, eintraege)

    ReDim loesung1(1 To nE)
    ReDim loesung2(1 To nE)
    ReDim istWSt1(1 To g_nL)
    ReDim istWSt2(1 To g_nL)

    For e = 1 To nE
        loesung1(e) = Trim(CStr(wsK.Cells(eintraege(e).zeile, 4).Value))
        loesung2(e) = Trim(CStr(wsK.Cells(eintraege(e).zeile, 5).Value))
    Next e
    For e = 1 To nE
        For i = 1 To g_nL
            If g_lehrer(i).name = loesung1(e) Then istWSt1(i) = istWSt1(i) + eintraege(e).WertUV
            If g_lehrer(i).name = loesung2(e) Then istWSt2(i) = istWSt2(i) + eintraege(e).WertUV
        Next i
    Next e

    Call BaueErgebnisSheets(loesung1, loesung2, istWSt1, istWSt2, eintraege, nE, "OR-Tools CP-SAT")
End Sub

' ============================================================
' DOMAINS (Constraint Propagation)
' ============================================================
Sub InitDomains(eintraege() As tEintrag, nE As Long, ByRef domains() As String)
    Dim e     As Long
    Dim e2    As Long
    Dim i     As Long
    Dim d     As String
    Dim zwang As String   ' Erzwungener Lehrer: selbe Klasse + selbes Fach bereits zugewiesen

    For e = 1 To nE
        If IstFixiert(eintraege(e).UrsprungsLehrer) Then
            domains(e) = "FIX"
        Else
            ' Pruefen ob ein anderer Eintrag mit gleicher Klasse+Fach schon zugewiesen ist
            ' -> dann denselben Lehrer erzwingen (Konsistenz-Regel)
            zwang = ""
            For e2 = 1 To nE
                If e2 <> e Then
                    If LCase(eintraege(e2).klasse) = LCase(eintraege(e).klasse) Then
                        If LCase(eintraege(e2).fach) = LCase(eintraege(e).fach) Then
                            If IstFixiert(eintraege(e2).lehrer) Then
                                zwang = eintraege(e2).lehrer
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next e2

            d = ""
            For i = 1 To g_nL
                If e = g_sperrIdx And g_lehrer(i).name = g_sperrName Then GoTo NaechsterI
                ' Konsistenz-Regel: wenn anderer Eintrag denselben Slot hat, nur diesen Lehrer zulassen
                If zwang <> "" And g_lehrer(i).name <> zwang Then GoTo NaechsterI
                If Not KannFach(g_lehrer(i), eintraege(e).fach, eintraege(e).IstOberstufe) Then GoTo NaechsterI
                If HatAntiPflicht(g_wuensche, g_nW, g_lehrer(i).name, eintraege(e).klasse) Then GoTo NaechsterI
                If g_lehrer(i).istWSt + eintraege(e).WertUV > g_lehrer(i).sollWst + g_toleranz Then GoTo NaechsterI
                If d <> "" Then d = d & ","
                d = d & i
NaechsterI:
            Next i
            ' Wenn Zwang gesetzt aber Lehrer nicht verfuegbar (z.B. Kapazitaet) -> ohne Zwang nochmal
            If zwang <> "" And d = "" Then
                For i = 1 To g_nL
                    If e = g_sperrIdx And g_lehrer(i).name = g_sperrName Then GoTo NaechsterI2
                    If g_lehrer(i).name <> zwang Then GoTo NaechsterI2
                    If Not KannFach(g_lehrer(i), eintraege(e).fach, eintraege(e).IstOberstufe) Then GoTo NaechsterI2
                    If d <> "" Then d = d & ","
                    d = d & i
NaechsterI2:
                Next i
            End If
            ' Fallback: wenn Domain leer (alle Lehrer ueber Kapazitaet), nochmal ohne Kapazitaetspruefung
            ' -> BT findet immer eine Loesung (mit moeglicher Ueberlast), analog zu SA
            If d = "" Then
                For i = 1 To g_nL
                    If e = g_sperrIdx And g_lehrer(i).name = g_sperrName Then GoTo NaechsterI3
                    If zwang <> "" And g_lehrer(i).name <> zwang Then GoTo NaechsterI3
                    If Not KannFach(g_lehrer(i), eintraege(e).fach, eintraege(e).IstOberstufe) Then GoTo NaechsterI3
                    If HatAntiPflicht(g_wuensche, g_nW, g_lehrer(i).name, eintraege(e).klasse) Then GoTo NaechsterI3
                    If d <> "" Then d = d & ","
                    d = d & i
NaechsterI3:
                Next i
            End If
            domains(e) = d
        End If
    Next e
End Sub

' ============================================================
' BACKTRACKING + ARC CONSISTENCY
' ============================================================
Function BacktrackingMitCP(eintraege() As tEintrag, nE As Long, _
                            domains() As String, _
                            loesung() As String) As Boolean
    ' Iterativer Stack statt Rekursion
    ' MAXDEPTH = nE: BT muss so tief gehen koennen wie es Eintraege gibt
    Dim MAXDEPTH As Long: MAXDEPTH = nE
    If MAXDEPTH > 2000 Then MAXDEPTH = 2000  ' Sicherheitsgrenze (g_stDom ist 1 To 2000)

    Dim stE(1 To 2000)             As Long
    Dim stLIdx(1 To 2000)          As Long
    Dim stKand(1 To 2000, 1 To 100) As Long
    Dim stScor(1 To 2000, 1 To 100) As Double
    Dim stNKand(1 To 2000)         As Long
    Dim stKPos(1 To 2000)          As Long

    Dim depth   As Long: depth = 0
    Dim e       As Long
    Dim bestE   As Long
    Dim bestSize As Long
    Dim dSize   As Long
    Dim lIdx    As Long
    Dim e3      As Long
    Dim cpOK    As Boolean
    Dim goDown  As Boolean  ' True = vorwaerts, False = naechsten Kandidaten probieren

    goDown = True

 Do While True

        If g_btMaxSek > 0 Then
            If Timer - g_btStartTime > g_btMaxSek Then
                BacktrackingMitCP = False: Exit Function
            End If
        End If

        If goDown Then
            ' === Schritt 1: naechsten unbesetzten Eintrag suchen (MRV) ===
            bestE = 0: bestSize = 999999
            For e = 1 To nE
                If Not IstFixiert(eintraege(e).UrsprungsLehrer) Then
                    If eintraege(e).lehrer = "" Or eintraege(e).lehrer = "?" Then
                        If domains(e) = "SKIP" Then GoTo NaechsterMRV
                        dSize = DomainGroesse(domains(e))
                        If dSize < bestSize Then bestSize = dSize: bestE = e
                    End If
                End If
NaechsterMRV:
            Next e

            ' Alle besetzt -> Loesung gefunden
            If bestE = 0 Then
                For e = 1 To nE
                    If IstFixiert(eintraege(e).UrsprungsLehrer) Then
                        loesung(e) = eintraege(e).UrsprungsLehrer
                    End If
                Next e
                BacktrackingMitCP = True
                Exit Function
            End If

            ' Domain leer -> Eintrag dauerhaft als unbesetzt markieren (SKIP), nicht wiederholen
            If bestSize = 0 Then
                eintraege(bestE).lehrer = "?"
                loesung(bestE) = "?"
                domains(bestE) = "SKIP"  ' verhindert erneute Auswahl durch MRV
                ' goDown bleibt True -> naechsten Eintrag suchen
            Else
                ' Stack voll -> Backtrack
                If depth >= MAXDEPTH Then
                    goDown = False
                Else
                    ' Neue Ebene anlegen
                    depth = depth + 1
                    stE(depth) = bestE
                    stLIdx(depth) = 0  ' noch kein Kandidat zugewiesen
                    Call DomainZuKandidaten2D(domains(bestE), eintraege, nE, bestE, depth, _
                         stKand, stScor, stNKand(depth))
                    stKPos(depth) = 1
                    ' Domains dieser Ebene sichern
                    For e3 = 1 To nE: g_stDom(depth, e3) = domains(e3): Next e3
                    goDown = False  ' jetzt Kandidaten probieren
                End If
            End If
        End If

        ' === Schritt 2: naechsten Kandidaten auf aktueller Ebene probieren ===
        If Not goDown Then
            ' Alle Kandidaten erschoepft -> Backtrack
            If depth = 0 Then
                BacktrackingMitCP = False
                Exit Function
            End If

            If stKPos(depth) > stNKand(depth) Then
                ' Backtrack: Ebene verlassen
                bestE = stE(depth)
                lIdx = stLIdx(depth)
                ' Zuweisung rueckgaengig machen falls gesetzt
                If eintraege(bestE).lehrer <> "?" And eintraege(bestE).lehrer <> "" Then
                    g_lehrer(lIdx).istWSt = g_lehrer(lIdx).istWSt - eintraege(bestE).WertUV
                    eintraege(bestE).lehrer = "?"
                    loesung(bestE) = "?"
                End If
                ' Domains dieser Ebene wiederherstellen
                For e3 = 1 To nE: domains(e3) = g_stDom(depth, e3): Next e3
                depth = depth - 1
                ' goDown bleibt False -> naechsten Kandidaten auf vorheriger Ebene
            Else
                ' Kandidaten probieren
                bestE = stE(depth)
                ' Vorherige Zuweisung zuruecknehmen falls vorhanden
                lIdx = stLIdx(depth)
                If lIdx > 0 Then
                    bestE = stE(depth)
                    If eintraege(bestE).lehrer <> "?" And eintraege(bestE).lehrer <> "" Then
                        g_lehrer(lIdx).istWSt = g_lehrer(lIdx).istWSt - eintraege(bestE).WertUV
                        eintraege(bestE).lehrer = "?"
                        loesung(bestE) = "?"
                    End If
                End If
                ' Domains dieser Ebene wiederherstellen
                For e3 = 1 To nE: domains(e3) = g_stDom(depth, e3): Next e3

                lIdx = stKand(depth, stKPos(depth))
                stLIdx(depth) = lIdx
                stKPos(depth) = stKPos(depth) + 1

                eintraege(bestE).lehrer = g_lehrer(lIdx).name
                g_lehrer(lIdx).istWSt = g_lehrer(lIdx).istWSt + eintraege(bestE).WertUV
                loesung(bestE) = g_lehrer(lIdx).name

                cpOK = ArcConsistency(eintraege, nE, domains, bestE, lIdx)

                If cpOK Then
                    goDown = True  ' vorwaerts zur naechsten Ebene
                Else
                    ' CP fehlgeschlagen: naechsten Kandidaten
                    g_lehrer(lIdx).istWSt = g_lehrer(lIdx).istWSt - eintraege(bestE).WertUV
                    eintraege(bestE).lehrer = "?"
                    loesung(bestE) = "?"
                    For e3 = 1 To nE: domains(e3) = g_stDom(depth, e3): Next e3
                    goDown = False
                End If
            End If
        End If

    Loop

    BacktrackingMitCP = False
End Function

Function ArcConsistency(eintraege() As tEintrag, nE As Long, _
                         domains() As String, _
                         zugewiesenerE As Long, zugewiesenerL As Long) As Boolean
    ' Leere Domains werden NICHT als Fehler gewertet:
    ' Eintraege ohne qualifizierten Lehrer (z.B. Meth ERG) werden
    ' vom BT-Hauptloop als "?" markiert und uebersprungen.
    ' ArcConsistency gibt nur dann False zurueck wenn ein Eintrag
    ' mit vorhandenen Kandidaten durch diese Zuweisung unlösbar wird
    ' UND g_toleranz=9999 nicht greift.
    Dim e As Long

    For e = 1 To nE
        If e = zugewiesenerE Then GoTo NaechsterE
        If IstFixiert(eintraege(e).UrsprungsLehrer) Then GoTo NaechsterE
        If eintraege(e).lehrer <> "" And eintraege(e).lehrer <> "?" Then GoTo NaechsterE
        ' Leere oder bereits uebersprungene Domain -> ignorieren
        If domains(e) = "" Or domains(e) = "SKIP" Then GoTo NaechsterE

        ' Konsistenz-Regel: gleiche Klasse + gleiches Fach -> zwingend selber Lehrer
        If LCase(eintraege(e).klasse) = LCase(eintraege(zugewiesenerE).klasse) Then
            If LCase(eintraege(e).fach) = LCase(eintraege(zugewiesenerE).fach) Then
                If g_lehrer(zugewiesenerL).istWSt + eintraege(e).WertUV > _
                   g_lehrer(zugewiesenerL).sollWst + g_toleranz Then
                    ' Lehrer hat keine Kapazitaet -> Domain leeren, aber nicht abbrechen
                    ' (zweiter Lauf mit g_toleranz=9999 loest das)
                    domains(e) = ""
                Else
                    domains(e) = CStr(zugewiesenerL)
                End If
                GoTo NaechsterE
            End If
        End If

        ' Lehrer aus Domain entfernen falls Kapazitaet erschoepft
        If g_lehrer(zugewiesenerL).istWSt + eintraege(e).WertUV > _
           g_lehrer(zugewiesenerL).sollWst + g_toleranz Then
            domains(e) = EntferneLehrerAusDomain(domains(e), zugewiesenerL)
        End If
        ' Domain leer nach Entfernen -> ignorieren, nicht abbrechen
NaechsterE:
    Next e
    ArcConsistency = True
End Function

Function DomainGroesse(domain As String) As Long
    If domain = "" Or domain = "FIX" Or domain = "SKIP" Then DomainGroesse = 0: Exit Function
    DomainGroesse = UBound(Split(domain, ",")) + 1
End Function

Function EntferneLehrerAusDomain(domain As String, lIdx As Long) As String
    Dim parts() As String
    Dim neu As String
    Dim p As Long
    parts = Split(domain, ",")
    neu = ""
    For p = 0 To UBound(parts)
        If CLng(Trim(parts(p))) <> lIdx Then
            If neu <> "" Then neu = neu & ","
            neu = neu & Trim(parts(p))
        End If
    Next p
    EntferneLehrerAusDomain = neu
End Function

Sub DomainZuKandidaten(domain As String, eintraege() As tEintrag, nE As Long, _
                        e As Long, _
                        ByRef kandidaten() As Long, ByRef scores() As Double, _
                        ByRef nKand As Long)
    Dim parts() As String
    Dim p As Long
    Dim lIdx As Long
    Dim si As Long
    Dim sj As Long
    Dim tmpI As Long
    Dim tmpS As Double

    parts = Split(domain, ",")
    nKand = UBound(parts) + 1
    ReDim kandidaten(1 To nKand)
    ReDim scores(1 To nKand)
    For p = 0 To UBound(parts)
        lIdx = CLng(Trim(parts(p)))
        kandidaten(p + 1) = lIdx
        scores(p + 1) = BerechneScore(lIdx, e, eintraege, nE)
    Next p
    For si = 2 To nKand
        tmpI = kandidaten(si): tmpS = scores(si): sj = si - 1
        Do While sj >= 1
            If scores(sj) < tmpS Then
                kandidaten(sj + 1) = kandidaten(sj)
                scores(sj + 1) = scores(sj)
                sj = sj - 1
            Else
                Exit Do
            End If
        Loop
        kandidaten(sj + 1) = tmpI: scores(sj + 1) = tmpS
    Next si
End Sub

Sub DomainZuKandidaten2D(domain As String, eintraege() As tEintrag, nE As Long, _
                          e As Long, depth As Long, _
                          ByRef stKand() As Long, ByRef stScor() As Double, _
                          ByRef nKand As Long)
    ' Schreibt Kandidaten direkt in Stack-Arrays (vermeidet ReDim auf Stack)
    Dim parts() As String
    Dim p As Long
    Dim lIdx As Long
    Dim si As Long
    Dim sj As Long
    Dim tmpI As Long
    Dim tmpS As Double

    parts = Split(domain, ",")
    nKand = UBound(parts) + 1
    If nKand > 60 Then nKand = 60  ' Sicherheitsgrenze (Stack-Array ist 60 breit)
    For p = 0 To nKand - 1
        lIdx = CLng(Trim(parts(p)))
        stKand(depth, p + 1) = lIdx
        stScor(depth, p + 1) = BerechneScore(lIdx, e, eintraege, nE)
    Next p
    For si = 2 To nKand
        tmpI = stKand(depth, si): tmpS = stScor(depth, si): sj = si - 1
        Do While sj >= 1
            If stScor(depth, sj) < tmpS Then
                stKand(depth, sj + 1) = stKand(depth, sj)
                stScor(depth, sj + 1) = stScor(depth, sj)
                sj = sj - 1
            Else
                Exit Do
            End If
        Loop
        stKand(depth, sj + 1) = tmpI: stScor(depth, sj + 1) = tmpS
    Next si
End Sub

Function BerechneScore(lIdx As Long, e As Long, _
                        eintraege() As tEintrag, nE As Long) As Double
    Dim sc As Double
    Dim wp As Long
    Dim ap As Long
    Dim k As Long

    sc = 0
    If g_lehrer(lIdx).KlassenleiterIn = eintraege(e).klasse And _
       g_lehrer(lIdx).KlassenleiterIn <> "" Then sc = sc + g_scoreKL
    wp = WunschPrioFuer(g_wuensche, g_nW, g_lehrer(lIdx).name, eintraege(e).klasse, True)
    If wp = 3 Then sc = sc + g_scoreW3
    If wp = 2 Then sc = sc + g_scoreW2
    If wp = 1 Then sc = sc + g_scoreW1
    ap = WunschPrioFuer(g_wuensche, g_nW, g_lehrer(lIdx).name, eintraege(e).klasse, False)
    If ap = 2 Then sc = sc - g_scoreA2
    If ap = 1 Then sc = sc - g_scoreA1
    For k = 1 To e - 1
        If eintraege(k).klasse = eintraege(e).klasse And _
           eintraege(k).lehrer = g_lehrer(lIdx).name Then
            sc = sc + g_scoreKont: Exit For
        End If
    Next k
    ' Kapazitaets-Score: einheitliche Skalierung
    ' freieStd > 0: Lehrer hat noch Luft  -> Bonus
    ' freieStd < 0: Lehrer ist ueberlastet -> Malus
    ' Beide Faktoren wirken direkt auf freieStd, daher vergleichbare Wertebereiche
    Dim freieStd As Double
    freieStd = g_lehrer(lIdx).sollWst - g_lehrer(lIdx).istWSt
    ' Auf sinnvollen Bereich begrenzen um Ueberlauf zu vermeiden
    If freieStd > 1000 Then freieStd = 1000
    If freieStd < -1000 Then freieStd = -1000
    If freieStd >= 0 Then
        sc = sc + freieStd * g_scoreFreiF
    Else
        sc = sc + freieStd * g_scoreUeberF
    End If
    ' Unterbesetzung: zusaetzlicher Malus wenn Lehrer nach Zuweisung leicht unter Soll faellt
    Dim nachZuweisung As Double
    nachZuweisung = freieStd - eintraege(e).WSt
    If nachZuweisung < 0 And nachZuweisung > -g_toleranz Then
        sc = sc + nachZuweisung * g_scoreUnterF  ' z.B. -0.5 Std * Faktor 5 = -2.5
    End If

    ' Schutzfachgruppen: Malus wenn Schutzlehrer hoch ausgelastet ist
    If g_lehrer(lIdx).IstSchutzLehrer Then
        Dim sfSoll As Double: sfSoll = g_lehrer(lIdx).sollWst
        Dim sfIst  As Double: sfIst = g_lehrer(lIdx).istWSt
        If sfSoll > 0 Then
            Dim sfAusl As Double: sfAusl = sfIst / sfSoll
            If sfAusl >= 0.8 Then
                sc = sc - g_schutzMalus * (sfAusl - 0.8) * 5
            End If
        End If
    End If

    BerechneScore = sc
End Function

Sub SchutzFlagsSetzen()
    ' Setzt IstSchutzLehrer-Flag einmalig nach LehrerEinlesen.
    ' Vermeidet wiederholten FachEngpassKey-Aufruf in BerechneScore.
    Dim i As Long, j As Long
    Dim sfk As String
    For i = 1 To g_nL
        g_lehrer(i).IstSchutzLehrer = False
        If g_schutzFG1 = "" And g_schutzFG2 = "" Then GoTo NaechsterSFL
        For j = 1 To 16
            If g_lehrer(i).fach(j) = "" Then GoTo NaechsterSFJ
            sfk = FachEngpassKey(g_lehrer(i).fach(j))
            If g_schutzFG1 <> "" Then
                If LCase(sfk) = LCase(g_schutzFG1) Then
                    g_lehrer(i).IstSchutzLehrer = True: GoTo NaechsterSFL
                End If
            End If
            If g_schutzFG2 <> "" Then
                If LCase(sfk) = LCase(g_schutzFG2) Then
                    g_lehrer(i).IstSchutzLehrer = True: GoTo NaechsterSFL
                End If
            End If
NaechsterSFJ:
        Next j
NaechsterSFL:
    Next i
End Sub

Sub ResetIstWSt(eintraege() As tEintrag, nE As Long)
    Dim i As Long
    Dim e As Long
    For i = 1 To g_nL: g_lehrer(i).istWSt = 0: Next i
    For e = 1 To nE
        If IstFixiert(eintraege(e).UrsprungsLehrer) Then
            eintraege(e).lehrer = eintraege(e).UrsprungsLehrer
            For i = 1 To g_nL
                If g_lehrer(i).name = eintraege(e).UrsprungsLehrer Then
                    g_lehrer(i).istWSt = g_lehrer(i).istWSt + eintraege(e).WertUV
                    Exit For
                End If
            Next i
        Else
            eintraege(e).lehrer = "?"
        End If
    Next e
End Sub

' ============================================================
' ERGEBNISSE SCHREIBEN
' ============================================================
Sub SchreibeLoesungen(wsK As Worksheet, eintraege() As tEintrag, nE As Long, _
                       loesung1() As String, loesung2() As String)
    Dim e As Long
    wsK.Cells(1, 4).Value = "Lehrer (L1)"
    wsK.Cells(1, 5).Value = "Lehrer (L2)"
    With wsK.Range(wsK.Cells(1, 4), wsK.Cells(1, 5))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    For e = 1 To nE
        wsK.Cells(eintraege(e).zeile, 4).Value = loesung1(e)
        wsK.Cells(eintraege(e).zeile, 5).Value = loesung2(e)
    Next e
End Sub

Sub BaueErgebnisSheets(loesung1() As String, loesung2() As String, _
                        istWSt1() As Double, istWSt2() As Double, _
                        eintraege() As tEintrag, nE As Long, _
                        methode As String)
    Dim lehr1() As tLehrer
    Dim lehr2() As tLehrer
    Dim eintr1() As tEintrag
    Dim eintr2() As tEintrag
    Dim i As Long
    Dim e As Long
    Dim meldung As String

    ReDim lehr1(1 To g_nL)
    ReDim lehr2(1 To g_nL)
    ReDim eintr1(1 To nE)
    ReDim eintr2(1 To nE)

    For i = 1 To g_nL
        lehr1(i) = g_lehrer(i): lehr1(i).istWSt = istWSt1(i)
        lehr2(i) = g_lehrer(i): lehr2(i).istWSt = istWSt2(i)
    Next i
    For e = 1 To nE
        eintr1(e) = eintraege(e): eintr1(e).lehrer = loesung1(e)
        eintr2(e) = eintraege(e): eintr2(e).lehrer = loesung2(e)
    Next e

    Call DiagnoseSheet_Erstellen(lehr1, g_nL, g_wuensche, g_nW, eintr1, nE, _
                                  SHEET_DIAG1, "Loesung 1 [" & methode & "]")
    Call DiagnoseSheet_Erstellen(lehr2, g_nL, g_wuensche, g_nW, eintr2, nE, _
                                  SHEET_DIAG2, "Loesung 2 [" & methode & "]")
    Call LehrerbelegungSheet_Erstellen(lehr1, g_nL, eintr1, nE, lehr2, eintr2, nE)
    Call FachgruppenLehrerSheet_Erstellen(g_lehrer, g_nL)

    SheetByName("Klassen").Activate

    meldung = "Optimierung abgeschlossen! Methode: " & methode & vbCrLf
    meldung = meldung & "Sheets: Klassen (Sp.4+5), " & SHEET_DIAG1 & ", " & SHEET_DIAG2 & ", " & SHEET_BELEGUNG
    MsgBox meldung, vbInformation, "Unterrichtsverteilung"
End Sub

' ============================================================
' DIAGNOSE-SHEET
' ============================================================
Sub DiagnoseSheet_Erstellen(lehrer() As tLehrer, nL As Long, _
                             wuensche() As tWunsch, nW As Long, _
                             eintraege() As tEintrag, nE As Long, _
                             sheetName As String, loesungsTitel As String)
    Dim wsD As Worksheet
    Dim cH  As Long
    Dim cSH As Long
    Dim cOK As Long
    Dim cWa As Long
    Dim cFe As Long
    Dim cNe As Long
    Dim cWe As Long
    Dim zRow As Long
    Dim i As Long
    Dim e As Long
    Dim w As Long
    Dim gSoll As Double
    Dim gIst As Double
    Dim diff As Double
    Dim sTxt As String
    Dim rc As Long
    Dim kl As String
    Dim prevL As String
    Dim altR As Boolean
    Dim hatE As Boolean
    Dim ec As Long
    Dim eigKl As Boolean
    Dim nb As Long
    Dim hatW As Boolean
    Dim wErf As Boolean
    Dim aVerl As Boolean
    Dim klA(200) As String
    Dim nKl As Long
    Dim ki As Long
    Dim fnd As Boolean
    Dim kF As Long
    Dim kW As Long
    Dim kB As Long
    Dim kO As Long

    Set wsD = SheetErzeugenOderLeeren(sheetName)
    cH = RGB(31, 73, 125): cSH = RGB(68, 114, 196)
    cOK = RGB(198, 239, 206): cWa = RGB(255, 235, 156)
    cFe = RGB(220, 80, 80): cNe = RGB(242, 242, 242)
    cWe = RGB(255, 255, 255)
    zRow = 1

    Call SectionHeader(wsD, zRow, loesungsTitel & " - DIAGNOSE", RGB(0, 70, 127), 8)
    zRow = zRow + 2

    Call SectionHeader(wsD, zRow, "1. LEHRERAUSLASTUNG", cH, 8)
    zRow = zRow + 1
    Call TableHeader(wsD, zRow, Array("Lehrer", "Ist (Wert)", "Soll (Wert)", "Ist-Soll", "+/-Tol?", "Klassen", "Status"), cSH)
    zRow = zRow + 1

    gSoll = 0: gIst = 0
    For i = 1 To nL
        diff = lehrer(i).istWSt - lehrer(i).sollWst
        If lehrer(i).istWSt = 0 Then
            sTxt = "Keine Zuweisung": rc = cWa
        ElseIf Abs(diff) <= g_toleranz Then
            sTxt = "OK": rc = cOK
        ElseIf diff > g_toleranz Then
            sTxt = "UEBERLASTET (+" & Format(Round(diff, 2), "0.00") & ")": rc = cFe
        Else
            sTxt = "Unterbesetzt (" & Format(Round(diff, 2), "0.00") & ")": rc = cWa
        End If
        kl = ""
        For e = 1 To nE
            If eintraege(e).lehrer = lehrer(i).name Then
                If InStr(kl, eintraege(e).klasse) = 0 Then
                    If kl <> "" Then kl = kl & ", "
                    kl = kl & eintraege(e).klasse
                End If
            End If
        Next e
        wsD.Cells(zRow, 1).Value = lehrer(i).name
        wsD.Cells(zRow, 2).Value = lehrer(i).istWSt
        wsD.Cells(zRow, 2).NumberFormat = "0.00"
        wsD.Cells(zRow, 3).Value = lehrer(i).sollWst
        wsD.Cells(zRow, 3).NumberFormat = "0.00"
        wsD.Cells(zRow, 4).Value = diff
        wsD.Cells(zRow, 4).NumberFormat = "0.00"
        wsD.Cells(zRow, 5).Value = IIf(Abs(diff) <= g_toleranz, "Ja", "Nein")
        wsD.Cells(zRow, 6).Value = kl
        wsD.Cells(zRow, 7).Value = sTxt
        Call FarbeZeile(wsD, zRow, 1, 7, rc)
        If diff > g_toleranz Then
            wsD.Cells(zRow, 4).Font.Bold = True
            wsD.Cells(zRow, 4).Font.Color = RGB(156, 0, 6)
        End If
        gSoll = gSoll + lehrer(i).sollWst: gIst = gIst + lehrer(i).istWSt
        zRow = zRow + 1
    Next i
    wsD.Cells(zRow, 1).Value = "GESAMT": wsD.Cells(zRow, 1).Font.Bold = True
    wsD.Cells(zRow, 2).Value = gIst:     wsD.Cells(zRow, 2).Font.Bold = True
    wsD.Cells(zRow, 2).NumberFormat = "0.00"
    wsD.Cells(zRow, 3).Value = gSoll:    wsD.Cells(zRow, 3).Font.Bold = True
    wsD.Cells(zRow, 3).NumberFormat = "0.00"
    wsD.Cells(zRow, 4).Value = gIst - gSoll: wsD.Cells(zRow, 4).Font.Bold = True
    wsD.Cells(zRow, 4).NumberFormat = "0.00"
    Call FarbeZeile(wsD, zRow, 1, 7, cNe)
    zRow = zRow + 2

    Call SectionHeader(wsD, zRow, "2. BELEGUNGSPLAN (je Lehrer)", cH, 6)
    zRow = zRow + 1
    Call TableHeader(wsD, zRow, Array("Lehrer", "Klasse", "Fach", "WSt", "Fixiert?", "KL?"), cSH)
    zRow = zRow + 1
    prevL = "": altR = False
    For i = 1 To nL
        hatE = False
        For e = 1 To nE
            If eintraege(e).lehrer = lehrer(i).name Then
                hatE = True
                If lehrer(i).name <> prevL Then altR = Not altR: prevL = lehrer(i).name
                ec = IIf(altR, cWe, cNe)
                eigKl = (lehrer(i).KlassenleiterIn = eintraege(e).klasse And lehrer(i).KlassenleiterIn <> "")
                wsD.Cells(zRow, 1).Value = lehrer(i).name
                wsD.Cells(zRow, 2).Value = eintraege(e).klasse
                wsD.Cells(zRow, 3).Value = eintraege(e).fach
                wsD.Cells(zRow, 4).Value = eintraege(e).WSt
                wsD.Cells(zRow, 5).Value = IIf(IstUrsprungsFixiert(eintraege(e)), "Ja", "-")
                wsD.Cells(zRow, 6).Value = IIf(eigKl, "KL", "")
                If eigKl Then wsD.Cells(zRow, 6).Font.Bold = True
                Call FarbeZeile(wsD, zRow, 1, 6, ec): zRow = zRow + 1
            End If
        Next e
        If Not hatE Then
            wsD.Cells(zRow, 1).Value = lehrer(i).name
            wsD.Cells(zRow, 2).Value = "(keine Zuweisung)": wsD.Cells(zRow, 2).Font.Italic = True
            Call FarbeZeile(wsD, zRow, 1, 6, cWa): zRow = zRow + 1
        End If
    Next i
    zRow = zRow + 1

    nb = 0
    For e = 1 To nE
        If eintraege(e).lehrer = "" Or eintraege(e).lehrer = "?" Then nb = nb + 1
    Next e
    Call SectionHeader(wsD, zRow, "3. NICHT BELEGTE STUNDEN (" & nb & ")", cH, 4)
    zRow = zRow + 1
    If nb = 0 Then
        wsD.Cells(zRow, 1).Value = "Alle Stunden belegt."
        wsD.Cells(zRow, 1).Font.Bold = True: wsD.Cells(zRow, 1).Font.Color = RGB(0, 97, 0)
        Call FarbeZeile(wsD, zRow, 1, 4, cOK): zRow = zRow + 2
    Else
        Call TableHeader(wsD, zRow, Array("Klasse", "Fach", "WSt", "Grund"), cSH): zRow = zRow + 1
        For e = 1 To nE
            If eintraege(e).lehrer = "" Or eintraege(e).lehrer = "?" Then
                wsD.Cells(zRow, 1).Value = eintraege(e).klasse
                wsD.Cells(zRow, 2).Value = eintraege(e).fach
                wsD.Cells(zRow, 3).Value = eintraege(e).WSt
                wsD.Cells(zRow, 4).Value = "Kein Lehrer zugewiesen (kein qualifizierter Lehrer oder Kapazitaet erschoepft)"
                Call FarbeZeile(wsD, zRow, 1, 4, cFe): zRow = zRow + 1
            End If
        Next e
        zRow = zRow + 1
    End If

    Call SectionHeader(wsD, zRow, "4. WUNSCH-ANALYSE", cH, 6): zRow = zRow + 1
    Call TableHeader(wsD, zRow, Array("Lehrer", "Art", "Klasse", "Prio", "Erfuellt?", "Bemerkung"), cSH)
    zRow = zRow + 1
    hatW = False
    For w = 1 To nW
        If wuensche(w).WunschKlasse <> "" And wuensche(w).WunschPrio > 0 Then
            hatW = True: wErf = False
            For e = 1 To nE
                If eintraege(e).klasse = wuensche(w).WunschKlasse And _
                   eintraege(e).lehrer = wuensche(w).lehrerName Then wErf = True: Exit For
            Next e
            wsD.Cells(zRow, 1).Value = wuensche(w).lehrerName
            wsD.Cells(zRow, 2).Value = "Wunsch"
            wsD.Cells(zRow, 3).Value = wuensche(w).WunschKlasse
            wsD.Cells(zRow, 4).Value = wuensche(w).WunschPrio
            wsD.Cells(zRow, 5).Value = IIf(wErf, "Ja", "Nein")
            If wuensche(w).WunschPrio = 3 And Not wErf Then
                wsD.Cells(zRow, 6).Value = "!!! PFLICHT-WUNSCH NICHT ERFUELLT"
                Call FarbeZeile(wsD, zRow, 1, 6, cFe)
            ElseIf Not wErf Then
                wsD.Cells(zRow, 6).Value = "Nicht erfuellt (Prio " & wuensche(w).WunschPrio & ")"
                Call FarbeZeile(wsD, zRow, 1, 6, cWa)
            Else
                wsD.Cells(zRow, 6).Value = "Erfuellt"
                Call FarbeZeile(wsD, zRow, 1, 6, cOK)
            End If
            zRow = zRow + 1
        End If
        If wuensche(w).AntiKlasse <> "" And wuensche(w).AntiPrio > 0 Then
            hatW = True: aVerl = False
            For e = 1 To nE
                If eintraege(e).klasse = wuensche(w).AntiKlasse And _
                   eintraege(e).lehrer = wuensche(w).lehrerName Then aVerl = True: Exit For
            Next e
            wsD.Cells(zRow, 1).Value = wuensche(w).lehrerName
            wsD.Cells(zRow, 2).Value = "Anti-Wunsch"
            wsD.Cells(zRow, 3).Value = wuensche(w).AntiKlasse
            wsD.Cells(zRow, 4).Value = wuensche(w).AntiPrio
            wsD.Cells(zRow, 5).Value = IIf(Not aVerl, "Eingehalten", "VERLETZT")
            If wuensche(w).AntiPrio = 3 And aVerl Then
                wsD.Cells(zRow, 6).Value = "!!! PFLICHT-ANTI-WUNSCH VERLETZT"
                Call FarbeZeile(wsD, zRow, 1, 6, cFe)
            ElseIf aVerl Then
                wsD.Cells(zRow, 6).Value = "Verletzt (Prio " & wuensche(w).AntiPrio & ")"
                Call FarbeZeile(wsD, zRow, 1, 6, cWa)
            Else
                wsD.Cells(zRow, 6).Value = "Eingehalten"
                Call FarbeZeile(wsD, zRow, 1, 6, cOK)
            End If
            zRow = zRow + 1
        End If
    Next w
    If Not hatW Then
        wsD.Cells(zRow, 1).Value = "Keine Wuensche.": wsD.Cells(zRow, 1).Font.Italic = True: zRow = zRow + 1
    End If
    zRow = zRow + 1

    wsD.Columns(1).ColumnWidth = 16: wsD.Columns(2).ColumnWidth = 12
    wsD.Columns(3).ColumnWidth = 16: wsD.Columns(4).ColumnWidth = 10
    wsD.Columns(5).ColumnWidth = 12: wsD.Columns(6).ColumnWidth = 28
    wsD.Columns(7).ColumnWidth = 28


    ' ------ Abschnitt 4: Ueberlast-Fach-Verteilung ------
    zRow = zRow + 2
    Call SectionHeader(wsD, zRow, "4. UEBERLAST-VERTEILUNG NACH FACHGRUPPE", cH, 8): zRow = zRow + 1
    wsD.Cells(zRow, 1).Value = "Je Fachgruppe: absolute Ueberlast-WSt aller betroffenen Lehrer + 50%-Anteil"
    wsD.Cells(zRow, 1).Font.Italic = True: zRow = zRow + 1
    Call TableHeader(wsD, zRow, Array("Fachgruppe", "Ueberlast gesamt (WSt)", "50%-Anteil (WSt)", "Betroffene Lehrer"), cSH)
    zRow = zRow + 1

    ' Alle Fachgruppen und ihre Ueberlast-Anteile sammeln
    Dim olFach(1 To 200)    As String
    Dim olWSt(1 To 200)     As Double  ' 50%-Anteil
    Dim olWStGes(1 To 200)  As Double  ' absolut
    Dim olLehrer(1 To 200)  As Long
    Dim olAnz As Long: olAnz = 0
    Dim oli As Long

    For oli = 1 To nL
        Dim olUeberl As Double
        ' Lehrer ohne Soll-WSt (0) ignorieren
        If lehrer(oli).sollWst <= 0 Then GoTo NaechsterOLLehrer
        olUeberl = lehrer(oli).istWSt - lehrer(oli).sollWst
        If olUeberl <= 0 Then GoTo NaechsterOLLehrer  ' nur ueberlastete Lehrer
        Dim olHalbe As Double: olHalbe = olUeberl / 2

        ' Nur tatsaechlich unterrichtete Fachgruppen beruecksichtigen
        ' (nicht alle Lehrerlisten-Faecher, sondern nur was in eintraege steht)
        Dim olGrpKeys(1 To 50) As String
        Dim olGrpAnz As Long: olGrpAnz = 0
        Dim olE As Long
        For olE = 1 To nE
            If eintraege(olE).lehrer <> lehrer(oli).name Then GoTo NaechsterOLE
            Dim olEKey As String: olEKey = FachEngpassKey(eintraege(olE).fach)
            If olEKey = "" Then GoTo NaechsterOLE
            ' Duplikat-Check
            Dim olEDup As Boolean: olEDup = False
            Dim olEK As Long
            For olEK = 1 To olGrpAnz
                If LCase(olGrpKeys(olEK)) = LCase(olEKey) Then olEDup = True: Exit For
            Next olEK
            If Not olEDup Then
                olGrpAnz = olGrpAnz + 1
                olGrpKeys(olGrpAnz) = olEKey
            End If
NaechsterOLE:
        Next olE
        If olGrpAnz = 0 Then GoTo NaechsterOLLehrer

        ' Ueberlast gleichmaessig auf tatsaechlich unterrichtete Fachgruppen aufteilen
        Dim olAnteil As Double: olAnteil = olUeberl / olGrpAnz
        Dim olHalbeA As Double: olHalbeA = olAnteil / 2
        Dim olGi As Long
        For olGi = 1 To olGrpAnz
            Dim olFI As Long: olFI = 0
            Dim olFK As Long
            For olFK = 1 To olAnz
                If LCase(olFach(olFK)) = LCase(olGrpKeys(olGi)) Then olFI = olFK: Exit For
            Next olFK
            If olFI = 0 Then
                olAnz = olAnz + 1: olFI = olAnz: olFach(olFI) = olGrpKeys(olGi)
            End If
            olWSt(olFI) = olWSt(olFI) + olHalbeA
            olWStGes(olFI) = olWStGes(olFI) + olAnteil
            olLehrer(olFI) = olLehrer(olFI) + 1
        Next olGi
NaechsterOLLehrer:
    Next oli

    ' Sortieren nach Gesamt-Ueberlast absteigend
    Dim olSI As Long, olSJ As Long
    Dim olTF As String, olTW As Double, olTWG As Double, olTL As Long
    For olSI = 1 To olAnz - 1
        For olSJ = olSI + 1 To olAnz
            If olWStGes(olSJ) > olWStGes(olSI) Then
                olTF = olFach(olSI):     olFach(olSI) = olFach(olSJ):       olFach(olSJ) = olTF
                olTW = olWSt(olSI):      olWSt(olSI) = olWSt(olSJ):         olWSt(olSJ) = olTW
                olTWG = olWStGes(olSI):  olWStGes(olSI) = olWStGes(olSJ):   olWStGes(olSJ) = olTWG
                olTL = olLehrer(olSI):   olLehrer(olSI) = olLehrer(olSJ):   olLehrer(olSJ) = olTL
            End If
        Next olSJ
    Next olSI

    ' Ausgabe
    For olSI = 1 To olAnz
        If olWStGes(olSI) <= 0 Then GoTo NaechsteOLZeile
        wsD.Cells(zRow, 1).Value = olFach(olSI)
        wsD.Cells(zRow, 2).Value = olWStGes(olSI)
        wsD.Cells(zRow, 3).Value = olWSt(olSI)
        wsD.Cells(zRow, 4).Value = olLehrer(olSI)
        wsD.Cells(zRow, 2).NumberFormat = "0.0"
        wsD.Cells(zRow, 3).NumberFormat = "0.0"
        Dim olColor As Long
        If olWStGes(olSI) > 20 Then
            olColor = cFe
        ElseIf olWStGes(olSI) > 10 Then
            olColor = RGB(255, 150, 80)
        ElseIf olWStGes(olSI) > 4 Then
            olColor = cWa
        Else
            olColor = cNe
        End If
        Call FarbeZeile(wsD, zRow, 1, 4, olColor)
        zRow = zRow + 1
NaechsteOLZeile:
    Next olSI

    wsD.Activate
    wsD.Cells(1, 1).Select
End Sub

' ============================================================
' LEHRERBELEGUNGS-SHEET
' ============================================================
Sub LehrerbelegungSheet_Erstellen(lehr1() As tLehrer, nL As Long, _
                                   eintr1() As tEintrag, nE1 As Long, _
                                   lehr2() As tLehrer, _
                                   eintr2() As tEintrag, nE2 As Long)
    Dim wsB As Worksheet
    Dim cH   As Long
    Dim cSH1 As Long
    Dim cSH2 As Long
    Dim cOK  As Long
    Dim cWa  As Long
    Dim cNe  As Long
    Dim cWe  As Long
    Dim zRow As Long
    Dim i As Long
    Dim e As Long
    Dim j As Long
    Dim hdrs As Variant
    Dim e1L(200) As Long
    Dim e2L(200) As Long
    Dim n1 As Long
    Dim n2 As Long
    Dim maxN As Long
    Dim rowIdx As Long
    Dim rc As Long
    Dim e1 As Long
    Dim e2 As Long
    Dim kL1 As Boolean
    Dim kL2 As Boolean
    Dim col As Long

    Set wsB = SheetErzeugenOderLeeren(SHEET_BELEGUNG)
    cH = RGB(31, 73, 125): cSH1 = RGB(68, 114, 196): cSH2 = RGB(142, 169, 219)
    cOK = RGB(198, 239, 206): cWa = RGB(255, 235, 156)
    cNe = RGB(242, 242, 242): cWe = RGB(255, 255, 255)
    zRow = 1

    Call SectionHeader(wsB, zRow, "LEHRERBELEGUNG - Loesung 1 vs. Loesung 2", RGB(0, 70, 127), 12)
    zRow = zRow + 2
    hdrs = Array("Klasse", "Fach", "WSt", "Fix?", "KL?", "Ist-WSt")

    For i = 1 To nL
        wsB.Range(wsB.Cells(zRow, 1), wsB.Cells(zRow, 12)).Merge
        wsB.Cells(zRow, 1).Value = lehr1(i).name & _
            "   Soll: " & lehr1(i).sollWst & " WSt" & _
            "   KL: " & IIf(lehr1(i).KlassenleiterIn <> "", lehr1(i).KlassenleiterIn, "-") & _
            "   Faecher: " & LehrerFaecherListe(lehr1(i))
        wsB.Cells(zRow, 1).Font.Bold = True: wsB.Cells(zRow, 1).Font.Size = 11
        wsB.Cells(zRow, 1).Interior.Color = cH: wsB.Cells(zRow, 1).Font.Color = RGB(255, 255, 255)
        wsB.rows(zRow).RowHeight = 18: zRow = zRow + 1

        wsB.Cells(zRow, 1).Value = "LOESUNG 1"
        With wsB.Range(wsB.Cells(zRow, 1), wsB.Cells(zRow, 6))
            .Interior.Color = cSH1: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
        wsB.Cells(zRow, 7).Value = "LOESUNG 2"
        With wsB.Range(wsB.Cells(zRow, 7), wsB.Cells(zRow, 12))
            .Interior.Color = cSH2: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
        zRow = zRow + 1

        For j = 0 To 5
            wsB.Cells(zRow, j + 1).Value = hdrs(j)
            wsB.Cells(zRow, j + 1).Interior.Color = cSH1
            wsB.Cells(zRow, j + 1).Font.Color = RGB(255, 255, 255)
            wsB.Cells(zRow, j + 7).Value = hdrs(j)
            wsB.Cells(zRow, j + 7).Interior.Color = cSH2
            wsB.Cells(zRow, j + 7).Font.Color = RGB(255, 255, 255)
        Next j
        zRow = zRow + 1

        n1 = 0: n2 = 0
        For e = 1 To nE1
            If eintr1(e).lehrer = lehr1(i).name Then n1 = n1 + 1: e1L(n1) = e
        Next e
        For e = 1 To nE2
            If eintr2(e).lehrer = lehr2(i).name Then n2 = n2 + 1: e2L(n2) = e
        Next e

        maxN = IIf(n1 > n2, n1, n2)
        For rowIdx = 1 To maxN
            rc = IIf(rowIdx Mod 2 = 1, cWe, cNe)
            If rowIdx <= n1 Then
                e1 = e1L(rowIdx)
                kL1 = (lehr1(i).KlassenleiterIn = eintr1(e1).klasse And lehr1(i).KlassenleiterIn <> "")
                wsB.Cells(zRow, 1).Value = eintr1(e1).klasse
                wsB.Cells(zRow, 2).Value = eintr1(e1).fach
                wsB.Cells(zRow, 3).Value = eintr1(e1).WSt
                wsB.Cells(zRow, 4).Value = IIf(IstUrsprungsFixiert(eintr1(e1)), "Ja", "-")
                wsB.Cells(zRow, 5).Value = IIf(kL1, "KL", "")
                If rowIdx = 1 Then wsB.Cells(zRow, 6).Value = lehr1(i).istWSt
                Call FarbeZeile(wsB, zRow, 1, 6, rc)
                If kL1 Then wsB.Cells(zRow, 5).Font.Bold = True
            End If
            If rowIdx <= n2 Then
                e2 = e2L(rowIdx)
                kL2 = (lehr2(i).KlassenleiterIn = eintr2(e2).klasse And lehr2(i).KlassenleiterIn <> "")
                wsB.Cells(zRow, 7).Value = eintr2(e2).klasse
                wsB.Cells(zRow, 8).Value = eintr2(e2).fach
                wsB.Cells(zRow, 9).Value = eintr2(e2).WSt
                wsB.Cells(zRow, 10).Value = IIf(IstUrsprungsFixiert(eintr2(e2)), "Ja", "-")
                wsB.Cells(zRow, 11).Value = IIf(kL2, "KL", "")
                If rowIdx = 1 Then wsB.Cells(zRow, 12).Value = lehr2(i).istWSt
                Call FarbeZeile(wsB, zRow, 7, 12, rc)
                If kL2 Then wsB.Cells(zRow, 11).Font.Bold = True
            End If
            zRow = zRow + 1
        Next rowIdx
        If maxN = 0 Then
            wsB.Cells(zRow, 1).Value = "(keine Zuweisung)": wsB.Cells(zRow, 1).Font.Italic = True
            wsB.Cells(zRow, 7).Value = "(keine Zuweisung)": wsB.Cells(zRow, 7).Font.Italic = True
            Call FarbeZeile(wsB, zRow, 1, 12, cWa): zRow = zRow + 1
        End If
        zRow = zRow + 1
    Next i

    For col = 1 To 12
        Select Case col
            Case 1, 7:  wsB.Columns(col).ColumnWidth = 10
            Case 2, 8:  wsB.Columns(col).ColumnWidth = 14
            Case 3, 9:  wsB.Columns(col).ColumnWidth = 6
            Case 4, 10: wsB.Columns(col).ColumnWidth = 7
            Case 5, 11: wsB.Columns(col).ColumnWidth = 6
            Case 6, 12: wsB.Columns(col).ColumnWidth = 8
        End Select
    Next col
    wsB.Activate
    wsB.Cells(1, 1).Select
End Sub

' ============================================================
' HILFSMAKROS
' ============================================================
Sub ZuweisungZuruecksetzen()
    Dim wsK As Worksheet
    Dim r As Long
    Set wsK = ThisWorkbook.Sheets("Klassen")
    Dim msg As String
    msg = "Welche Spalten zuruecksetzen?" & vbCrLf & vbCrLf & _
          "Ja    = L1, L2, L3, L4 und Fix auf '?' setzen" & vbCrLf & _
          "Nein  = Nur L1, L2, L3, L4 (Fix behalten)"
    Dim antwort As VbMsgBoxResult
    antwort = MsgBox(msg, vbYesNoCancel + vbQuestion, "Zuweisung zuruecksetzen")
    If antwort = vbCancel Then Exit Sub
    Dim resetFix As Boolean: resetFix = (antwort = vbYes)
    For r = 2 To wsK.Cells(wsK.rows.Count, 2).End(xlUp).row
        If wsK.Cells(r, 2).Value <> "" Then
            wsK.Cells(r, 4).Value = "?"
            wsK.Cells(r, 5).Value = "?"
            wsK.Cells(r, 7).Value = "?"
            wsK.Cells(r, 8).Value = "?"
            wsK.Cells(r, 9).Value = 0
            If resetFix Then wsK.Cells(r, 6).Value = "?"
        End If
    Next r
    MsgBox "Zurueckgesetzt.", vbInformation
End Sub

Sub SectionHeader(ws As Worksheet, row As Long, titel As String, farbe As Long, breite As Long)
    ws.Range(ws.Cells(row, 1), ws.Cells(row, breite)).Merge
    ws.Cells(row, 1).Value = titel: ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Color = RGB(255, 255, 255): ws.Cells(row, 1).Font.Size = 12
    ws.Cells(row, 1).Interior.Color = farbe: ws.rows(row).RowHeight = 20
End Sub

Sub TableHeader(ws As Worksheet, row As Long, headers As Variant, farbe As Long)
    Dim j As Long
    For j = 0 To UBound(headers)
        ws.Cells(row, j + 1).Value = headers(j): ws.Cells(row, j + 1).Font.Bold = True
        ws.Cells(row, j + 1).Font.Color = RGB(255, 255, 255): ws.Cells(row, j + 1).Interior.Color = farbe
        ws.Cells(row, j + 1).HorizontalAlignment = xlCenter
    Next j
End Sub

Sub FarbeZeile(ws As Worksheet, row As Long, colVon As Long, colBis As Long, farbe As Long)
    ws.Range(ws.Cells(row, colVon), ws.Cells(row, colBis)).Interior.Color = farbe
    ws.Range(ws.Cells(row, colVon), ws.Cells(row, colBis)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range(ws.Cells(row, colVon), ws.Cells(row, colBis)).Borders(xlEdgeBottom).Color = RGB(200, 200, 200)
End Sub

Function SheetByName(sName As String) As Worksheet
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.name = sName Then Set SheetByName = sh: Exit Function
    Next sh
    Set SheetByName = Nothing
End Function

Function SheetErzeugenOderLeeren(sName As String) As Worksheet
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.name = sName Then
            sh.Cells.Clear: sh.Cells.UnMerge
            Set SheetErzeugenOderLeeren = sh: Exit Function
        End If
    Next sh
    Set sh = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    sh.name = sName: Set SheetErzeugenOderLeeren = sh
End Function

Function LokalisierePfad(urlOderPfad As String) As String
    ' Wandelt OneDrive-URL in lokalen Windows-Pfad um.
    ' Gibt den Pfad unveraendert zurueck wenn es bereits ein lokaler Pfad ist.
    Dim fN        As String
    Dim odBase    As String
    Dim urlPfad   As String
    Dim ci        As Long
    Dim slashCount As Long
    Dim lastBS    As Long
    Dim firstBS   As Long

    fN = urlOderPfad
    If left(fN, 4) <> "http" Then
        LokalisierePfad = fN   ' Bereits lokaler Pfad
        Exit Function
    End If

    ' OneDrive-URL aufloesen
    odBase = Environ("OneDrive")
    If odBase = "" Then odBase = Environ("OneDriveConsumer")
    If odBase = "" Then odBase = Environ("OneDriveCommercial")
    If odBase = "" Then LokalisierePfad = "": Exit Function

    ' Pfad nach dem 4. Slash extrahieren
    slashCount = 0
    urlPfad = ""
    For ci = 1 To Len(fN)
        If Mid(fN, ci, 1) = "/" Then slashCount = slashCount + 1
        If slashCount = 4 Then
            urlPfad = Mid(fN, ci + 1)
            Exit For
        End If
    Next ci

    ' URL-Separatoren in Windows-Pfad umwandeln
    urlPfad = Replace(urlPfad, "/", "\")

    ' Ersten Teilpfad (OneDrive-Rootordnername in URL) entfernen
    firstBS = InStr(urlPfad, "\")
    If firstBS > 0 Then urlPfad = Mid(urlPfad, firstBS)

    LokalisierePfad = odBase & urlPfad
End Function

Function SucheExe() As String
    ' Sucht stundenplan_solver.exe - OneDrive-kompatibel
    Dim fso       As Object
    Dim pfad      As String
    Dim fN        As String
    Dim odBase    As String
    Dim urlPfad   As String
    Dim ci        As Long
    Dim slashCount As Long
    Dim lastBS    As Long
    Dim firstBS   As Long
    Dim fe        As Boolean

    SucheExe = ""
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    If fso Is Nothing Then Exit Function

    fN = ThisWorkbook.FullName

    If left(fN, 4) = "http" Then
        ' OneDrive-URL: lokalen Sync-Pfad rekonstruieren
        odBase = Environ("OneDrive")
        If odBase = "" Then odBase = Environ("OneDriveConsumer")
        If odBase = "" Then odBase = Environ("OneDriveCommercial")

        If odBase <> "" Then
            ' Pfad nach dem 4. Slash extrahieren
            slashCount = 0
            urlPfad = ""
            For ci = 1 To Len(fN)
                If Mid(fN, ci, 1) = "/" Then slashCount = slashCount + 1
                If slashCount = 4 Then
                    urlPfad = Mid(fN, ci + 1)
                    Exit For
                End If
            Next ci
            ' Dateinamen entfernen
            urlPfad = Replace(urlPfad, "/", "\")
            lastBS = 0
            For ci = Len(urlPfad) To 1 Step -1
                If Mid(urlPfad, ci, 1) = "\" Then lastBS = ci: Exit For
            Next ci
            If lastBS > 0 Then urlPfad = left(urlPfad, lastBS - 1)
            ' Ersten Teilpfad (OneDrive-Rootname) entfernen
            firstBS = InStr(urlPfad, "\")
            If firstBS > 0 Then urlPfad = Mid(urlPfad, firstBS)
            pfad = odBase & urlPfad & "\" & SOLVER_EXE_NAME
            fe = False
            On Error Resume Next: fe = fso.FileExists(pfad): On Error GoTo 0
            If fe Then SucheExe = pfad: Set fso = Nothing: Exit Function
        End If
    Else
        ' Normaler lokaler Pfad
        pfad = ThisWorkbook.Path & "\" & SOLVER_EXE_NAME
        fe = False
        On Error Resume Next: fe = fso.FileExists(pfad): On Error GoTo 0
        If fe Then SucheExe = pfad: Set fso = Nothing: Exit Function
    End If

    ' Fallback: CurDir
    pfad = CurDir & "\" & SOLVER_EXE_NAME
    fe = False
    On Error Resume Next: fe = fso.FileExists(pfad): On Error GoTo 0
    If fe Then SucheExe = pfad: Set fso = Nothing: Exit Function

    ' Letzter Ausweg: Nutzer manuell nach EXE fragen
    Dim manPfad As String
    manPfad = Application.GetOpenFilename( _
        FileFilter:="stundenplan_solver.exe,stundenplan_solver.exe", _
        Title:="stundenplan_solver.exe auswaehlen")
    If manPfad <> "False" And manPfad <> "" Then
        SucheExe = manPfad
    End If

    Set fso = Nothing
End Function

Function LehrerFaecherListe(lehr As tLehrer) As String
    Dim s As String
    Dim f As Long
    s = ""
    For f = 1 To 16
        If lehr.fach(f) <> "" Then
            If s <> "" Then s = s & ", "
            s = s & lehr.fach(f)
        End If
    Next f
    LehrerFaecherListe = s
End Function

' ============================================================
' FACHGRUPPEN laden (nur lesen, kein Activate/MsgBox)
' ============================================================
Sub FachgruppenLaden()
    Dim ws As Worksheet
    Set ws = SheetByName(SHEET_FGDEF)
    If ws Is Nothing Then
        Call FachgruppenSheet_Erstellen(True)
        Set ws = SheetByName(SHEET_FGDEF)
        If ws Is Nothing Then Exit Sub
    End If
    g_fgAnz = 0
    Dim lastR As Long: lastR = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    Dim r As Long
    For r = 2 To lastR
        Dim fN As String: fN = Trim(CStr(ws.Cells(r, 1).Value))
        Dim gN As String: gN = Trim(CStr(ws.Cells(r, 2).Value))
        If fN = "" Or gN = "" Then GoTo NaechsteFGZ
        If left(fN, 1) = "'" Then GoTo NaechsteFGZ
        g_fgAnz = g_fgAnz + 1
        g_fgFach(g_fgAnz) = fN
        g_fgGruppe(g_fgAnz) = gN
NaechsteFGZ:
    Next r
End Sub

' ============================================================
' FACHGRUPPEN-SHEET erstellen
' ============================================================
Sub FachgruppenSheet_Erstellen(Optional stille As Boolean = False)
    Dim ws As Worksheet
    Set ws = SheetByName(SHEET_FGDEF)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = SHEET_FGDEF
    Else
        ws.Cells.Clear
    End If
    Dim cH As Long: cH = RGB(31, 73, 125)
    Dim cSH As Long: cSH = RGB(68, 114, 196)
    ws.Cells(1, 1).Value = "Fachgruppen-Definition"
    ws.Cells(1, 1).Font.Bold = True: ws.Cells(1, 1).Font.Size = 13
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 3)).Merge
    ws.Cells(1, 1).Interior.Color = cH: ws.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    ws.Cells(2, 1).Value = "Fach (Kuerzel)"
    ws.Cells(2, 2).Value = "Fachgruppe"
    ws.Cells(2, 3).Value = "Hinweis (optional)"
    With ws.Range(ws.Cells(2, 1), ws.Cells(2, 3))
        .Font.Bold = True: .Interior.Color = cSH: .Font.Color = RGB(255, 255, 255)
    End With
    Dim r As Long: r = 3
    Dim fg(1 To 35, 1 To 3) As String
    fg(1, 1) = "D":     fg(1, 2) = "D":     fg(1, 3) = "Deutsch"
    fg(2, 1) = "M":     fg(2, 2) = "M":     fg(2, 3) = "Mathematik"
    fg(3, 1) = "E":     fg(3, 2) = "E":     fg(3, 3) = "Englisch"
    fg(4, 1) = "F":     fg(4, 2) = "F":     fg(4, 3) = "Franzoesisch"
    fg(5, 1) = "L":     fg(5, 2) = "L":     fg(5, 3) = "Latein"
    fg(6, 1) = "L9":    fg(6, 2) = "L9":    fg(6, 3) = "Latein ab Kl. 9"
    fg(7, 1) = "SP":    fg(7, 2) = "SP":    fg(7, 3) = "Sport"
    fg(8, 1) = "BI":    fg(8, 2) = "BI":    fg(8, 3) = "Biologie"
    fg(9, 1) = "CH":    fg(9, 2) = "CH":    fg(9, 3) = "Chemie"
    fg(10, 1) = "PH":   fg(10, 2) = "PH":   fg(10, 3) = "Physik"
    fg(11, 1) = "GE":   fg(11, 2) = "GE":   fg(11, 3) = "Geschichte"
    fg(12, 1) = "EK":   fg(12, 2) = "EK":   fg(12, 3) = "Erdkunde"
    fg(13, 1) = "KU":   fg(13, 2) = "KU":   fg(13, 3) = "Kunst"
    fg(14, 1) = "MU":   fg(14, 2) = "MU":   fg(14, 3) = "Musik"
    fg(15, 1) = "KR":   fg(15, 2) = "KR":   fg(15, 3) = "Kath. Religion"
    fg(16, 1) = "ER":   fg(16, 2) = "ER":   fg(16, 3) = "Evang. Religion"
    fg(17, 1) = "PP":   fg(17, 2) = "PP":   fg(17, 3) = "Prakt. Philosophie"
    fg(18, 1) = "IF":   fg(18, 2) = "IF":   fg(18, 3) = "Informatik"
    fg(19, 1) = "WiPo": fg(19, 2) = "WiPo": fg(19, 3) = "Wirtschaft/Politik"
    fg(20, 1) = "SW":   fg(20, 2) = "WiPo": fg(20, 3) = "Sozialwiss. -> WiPo"
    fg(21, 1) = "GL":   fg(21, 2) = "GL":   fg(21, 3) = "Gesellschaftslehre"
    fg(22, 1) = "TW":   fg(22, 2) = "TW":   fg(22, 3) = "Technik/Wirtschaft"
    fg(23, 1) = "BK":   fg(23, 2) = "BK":   fg(23, 3) = "Bildende Kunst"
    fg(24, 1) = "DS":   fg(24, 2) = "DS":   fg(24, 3) = "Darstellendes Spiel"
    fg(25, 1) = "RB2":  fg(25, 2) = "RB2":  fg(25, 3) = "Rel. Bildung 2"
    fg(26, 1) = "DELF": fg(26, 2) = "DELF": fg(26, 3) = "DELF-Kurs"
    fg(27, 1) = "IKTreff": fg(27, 2) = "IKTreff": fg(27, 3) = "IKT-Treff"
    fg(28, 1) = "KL_Erg": fg(28, 2) = "KL_Erg": fg(28, 3) = "KL-Ergaenzung"
    fg(29, 1) = "DelF": fg(29, 2) = "DelF": fg(29, 3) = "Delegation Franz."
    fg(30, 1) = "TH":   fg(30, 2) = "TH":   fg(30, 3) = "Theater"
    fg(31, 1) = "ETH":  fg(31, 2) = "ETH":  fg(31, 3) = "Ethik"
    fg(32, 1) = "AG":   fg(32, 2) = "AG":   fg(32, 3) = "Arbeitsgem."
    fg(33, 1) = "PO":   fg(33, 2) = "WiPo": fg(33, 3) = "Politik -> WiPo"
    fg(34, 1) = "CC":   fg(34, 2) = "E":    fg(34, 3) = "Cambridge -> E"
    fg(35, 1) = "LI":   fg(35, 2) = "D":    fg(35, 3) = "Lese-Int. -> D"
    Dim d As Long
    For d = 1 To 35
        ws.Cells(r, 1).Value = fg(d, 1): ws.Cells(r, 2).Value = fg(d, 2)
        ws.Cells(r, 3).Value = fg(d, 3)
        ws.Cells(r, 3).Font.Italic = True: ws.Cells(r, 3).Font.Color = RGB(128, 128, 128)
        If r Mod 2 = 0 Then ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Interior.Color = RGB(242, 242, 242)
        r = r + 1
    Next d
    r = r + 1
    ws.Cells(r, 1).Value = "Neue Zeilen ergaenzen. Faecher ohne Eintrag werden ignoriert."
    ws.Cells(r, 1).Font.Italic = True: ws.Cells(r, 1).Font.Color = RGB(128, 128, 128)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Merge
    ws.Columns(1).ColumnWidth = 16: ws.Columns(2).ColumnWidth = 16: ws.Columns(3).ColumnWidth = 36
    g_fgAnz = 0
    If Not stille Then
        ws.Activate
        MsgBox "Sheet '" & SHEET_FGDEF & "' erstellt. Bitte Fachgruppen pruefen.", vbInformation
    End If
End Sub

' ============================================================
' FACHGRUPPEN-LEHRER-SHEET
' ============================================================
Sub FachgruppenLehrerSheet_Erstellen(lehrer() As tLehrer, nL As Long)
    Dim ws As Worksheet
    Dim cH As Long: cH = RGB(31, 73, 125)
    Dim cSH As Long: cSH = RGB(68, 114, 196)
    Dim cNe As Long: cNe = RGB(242, 242, 242)
    Dim cRed As Long: cRed = RGB(255, 199, 206)
    Dim zRow As Long: zRow = 1
    Dim i As Long, fj As Long, fk As Long
    Dim fGrp As String, fN As String
    Dim grpKeys(1 To 50) As String
    Dim nGrp As Long, isDup As Boolean
    g_fgAnz = 0: Call FachgruppenLaden
    Set ws = SheetErzeugenOderLeeren(SHEET_FACHGRUPPEN)
    Call SectionHeader(ws, zRow, "FACHGRUPPEN JE LEHRER", cH, 5): zRow = zRow + 2
    ws.Cells(zRow, 1).Value = "Lehrer": ws.Cells(zRow, 2).Value = "Soll-WSt"
    ws.Cells(zRow, 3).Value = "Fachgruppen (Engpass-relevant)"
    ws.Cells(zRow, 4).Value = "Alle eingetragenen Faecher"
    ws.Cells(zRow, 5).Value = "Anzahl Gruppen"
    With ws.Range(ws.Cells(zRow, 1), ws.Cells(zRow, 5))
        .Interior.Color = cSH: .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
    End With
    zRow = zRow + 1
    For i = 1 To nL
        nGrp = 0
        Dim allF As String: allF = ""
        For fj = 1 To 16
            fN = Trim(lehrer(i).fach(fj))
            If fN = "" Then GoTo NxtFJ
            If allF <> "" Then allF = allF & ", "
            allF = allF & fN
            fGrp = FachEngpassKey(fN)
            If fGrp = "" Then GoTo NxtFJ
            isDup = False
            For fk = 1 To nGrp
                If LCase(grpKeys(fk)) = LCase(fGrp) Then isDup = True: Exit For
            Next fk
            If Not isDup Then nGrp = nGrp + 1: grpKeys(nGrp) = fGrp
NxtFJ:
        Next fj
        Dim grpLst As String: grpLst = ""
        For fk = 1 To nGrp
            If grpLst <> "" Then grpLst = grpLst & ", "
            grpLst = grpLst & grpKeys(fk)
        Next fk
        Dim rc As Long: rc = IIf(i Mod 2 = 0, cNe, RGB(255, 255, 255))
        ws.Cells(zRow, 1).Value = lehrer(i).name: ws.Cells(zRow, 2).Value = lehrer(i).sollWst
        ws.Cells(zRow, 3).Value = grpLst: ws.Cells(zRow, 4).Value = allF
        ws.Cells(zRow, 5).Value = nGrp
        If allF <> "" And nGrp = 0 Then
            Call FarbeZeile(ws, zRow, 1, 5, cRed)
        Else
            Call FarbeZeile(ws, zRow, 1, 5, rc)
        End If
        ws.Cells(zRow, 2).NumberFormat = "0.#": zRow = zRow + 1
    Next i
    ws.Columns(1).ColumnWidth = 20: ws.Columns(2).ColumnWidth = 10
    ws.Columns(3).ColumnWidth = 40: ws.Columns(4).ColumnWidth = 50
    ws.Columns(5).ColumnWidth = 14
    ws.Activate: ws.Cells(1, 1).Select
End Sub

' ============================================================
' NACHTRAEGLICHE VERBESSERUNG (SA) -> L3 / L4
' Verbessert L1->L3 und L2->L4 durch Simulated Annealing
' Ziel: Ueber-/Unterbelastungen ausgleichen
' ============================================================
Sub NachtraeglicheVerbesserung()
    Dim wsL As Worksheet, wsK As Worksheet, wsW As Worksheet
    Dim eintraege() As tEintrag
    Dim nE As Long
    Dim i As Long, e As Long

    On Error GoTo NVFehler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wsL = SheetByName("Lehrerliste")
    Set wsK = SheetByName("Klassen")
    Set wsW = SheetByName("Lehrerw" & Chr(252) & "nsche")
    If wsL Is Nothing Or wsK Is Nothing Then
        MsgBox "Lehrerliste oder Klassen fehlt!", vbCritical: GoTo NVCleanup
    End If

    Call ParameterEinlesen
    ReDim g_lehrer(1 To 1)
    g_nL = LehrerEinlesen(wsL, g_lehrer)
    Call SchutzFlagsSetzen
    If g_nL = 0 Then MsgBox "Keine Lehrer!", vbCritical: GoTo NVCleanup
    ReDim g_wuensche(1 To 1)
    If Not wsW Is Nothing Then g_nW = WuenscheEinlesen(wsW, g_wuensche)

    nE = EintraegeEinlesen(wsK, eintraege)
    If nE = 0 Then MsgBox "Keine Eintraege!", vbCritical: GoTo NVCleanup

    ' L1 aus Spalte 4 lesen
    Dim loes3() As String: ReDim loes3(1 To nE)
    Dim loes4() As String: ReDim loes4(1 To nE)
    For e = 1 To nE
        loes3(e) = Trim(CStr(wsK.Cells(eintraege(e).zeile, 4).Value))
        loes4(e) = Trim(CStr(wsK.Cells(eintraege(e).zeile, 5).Value))
    Next e

    Dim zeitProLoes As Double: zeitProLoes = g_zeitlimit / 2

    ' L3: L1 verbessern
    Application.StatusBar = "Verbessere L1 -> L3 (" & CInt(zeitProLoes) & " Sek.)..."
    Call VerbessereSA(eintraege, nE, loes3, zeitProLoes)

    ' IstWSt fuer L3 berechnen
    Dim istWSt3() As Double: ReDim istWSt3(1 To g_nL)
    Call BerechneIstWSt(eintraege, nE, loes3, istWSt3)

    ' L4: L2 verbessern
    Application.StatusBar = "Verbessere L2 -> L4 (" & CInt(zeitProLoes) & " Sek.)..."
    Call VerbessereSA(eintraege, nE, loes4, zeitProLoes)

    Dim istWSt4() As Double: ReDim istWSt4(1 To g_nL)
    Call BerechneIstWSt(eintraege, nE, loes4, istWSt4)

    ' In Sheet schreiben (Spalten 7+8)
    wsK.Cells(1, 7).Value = "Lehrer (L3)": wsK.Cells(1, 8).Value = "Lehrer (L4)"
    With wsK.Range(wsK.Cells(1, 7), wsK.Cells(1, 8))
        .Font.Bold = True: .Interior.Color = RGB(142, 169, 219): .Font.Color = RGB(255, 255, 255)
    End With
    For e = 1 To nE
        wsK.Cells(eintraege(e).zeile, 7).Value = loes3(e)
        wsK.Cells(eintraege(e).zeile, 8).Value = loes4(e)
    Next e

    ' Diagnose3 + Diagnose4 erstellen
    Dim lehr3() As tLehrer: ReDim lehr3(1 To g_nL)
    Dim lehr4() As tLehrer: ReDim lehr4(1 To g_nL)
    Dim eintr3() As tEintrag: ReDim eintr3(1 To nE)
    Dim eintr4() As tEintrag: ReDim eintr4(1 To nE)
    For i = 1 To g_nL
        lehr3(i) = g_lehrer(i): lehr3(i).istWSt = istWSt3(i)
        lehr4(i) = g_lehrer(i): lehr4(i).istWSt = istWSt4(i)
    Next i
    For e = 1 To nE
        eintr3(e) = eintraege(e): eintr3(e).lehrer = loes3(e)
        eintr4(e) = eintraege(e): eintr4(e).lehrer = loes4(e)
    Next e
    Call DiagnoseSheet_Erstellen(lehr3, g_nL, g_wuensche, g_nW, eintr3, nE, SHEET_DIAG3, "Loesung 3 [SA-Verbesserung]")
    Call DiagnoseSheet_Erstellen(lehr4, g_nL, g_wuensche, g_nW, eintr4, nE, SHEET_DIAG4, "Loesung 4 [SA-Verbesserung]")

    SheetByName("Klassen").Activate
    Application.StatusBar = False
    MsgBox "Verbesserung abgeschlossen! L3/L4 in Spalten 7/8, Diagnose3/4 erstellt.", vbInformation

NVCleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
NVFehler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Fehler " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

Sub BerechneIstWSt(eintraege() As tEintrag, nE As Long, loesung() As String, istWSt() As Double)
    Dim e As Long, i As Long
    For i = 1 To g_nL: istWSt(i) = 0: Next i
    For e = 1 To nE
        For i = 1 To g_nL
            If g_lehrer(i).name = loesung(e) Then
                istWSt(i) = istWSt(i) + eintraege(e).WertUV: Exit For
            End If
        Next i
    Next e
End Sub

' Prueft ob newLIdx fuer ALLE Eintraege mit gleicher Klasse+Fach qualifiziert ist
' und ob der Tausch konsistent durchgefuehrt werden kann
Function KonsistenzOK(eintraege() As tEintrag, nE As Long, _
                       eRef As Long, newLIdx As Long, _
                       loesung() As String) As Boolean
    Dim e As Long
    For e = 1 To nE
        If e = eRef Then GoTo NaechstKons
        If LCase(eintraege(e).klasse) <> LCase(eintraege(eRef).klasse) Then GoTo NaechstKons
        If LCase(eintraege(e).fach) <> LCase(eintraege(eRef).fach) Then GoTo NaechstKons
        ' Gleiche Klasse+Fach gefunden: neuer Lehrer muss qualifiziert sein
        If Not KannFach(g_lehrer(newLIdx), eintraege(e).fach, eintraege(e).IstOberstufe) Then
            KonsistenzOK = False: Exit Function
        End If
        ' Und der Eintrag darf nicht fixiert sein (sonst kann er nicht geaendert werden)
        If IstFixiert(eintraege(e).UrsprungsLehrer) Then
            ' Fixierter Eintrag hat anderen Lehrer -> Konsistenz nicht moeglich
            If loesung(e) <> g_lehrer(newLIdx).name Then
                KonsistenzOK = False: Exit Function
            End If
        End If
NaechstKons:
    Next e
    KonsistenzOK = True
End Function

' Weist newLIdx allen Eintraegen mit gleicher Klasse+Fach wie eRef zu
' Gibt Liste der geaenderten Eintraege zurueck (fuer Rueckgaengig-Funktion)
Sub KonsistenzZuweisen(eintraege() As tEintrag, nE As Long, _
                        eRef As Long, newLIdx As Long, oldLIdx As Long, _
                        loesung() As String)
    Dim e As Long
    For e = 1 To nE
        If e = eRef Then GoTo NaechstKZ
        If LCase(eintraege(e).klasse) <> LCase(eintraege(eRef).klasse) Then GoTo NaechstKZ
        If LCase(eintraege(e).fach) <> LCase(eintraege(eRef).fach) Then GoTo NaechstKZ
        If IstFixiert(eintraege(e).UrsprungsLehrer) Then GoTo NaechstKZ
        ' Alten Lehrer IstWSt reduzieren
        Dim oldIdx As Long: oldIdx = LehrerIndex(g_lehrer, g_nL, loesung(e))
        If oldIdx > 0 And oldIdx <> newLIdx Then
            g_lehrer(oldIdx).istWSt = g_lehrer(oldIdx).istWSt - eintraege(e).WertUV
        End If
        ' Neuen Lehrer zuweisen
        g_lehrer(newLIdx).istWSt = g_lehrer(newLIdx).istWSt + eintraege(e).WertUV
        loesung(e) = g_lehrer(newLIdx).name
NaechstKZ:
    Next e
End Sub

Sub KonsistenzRueckgaengig(eintraege() As tEintrag, nE As Long, _
                             eRef As Long, newLIdx As Long, _
                             altLoesung() As String, loesung() As String)
    Dim e As Long
    For e = 1 To nE
        If e = eRef Then GoTo NaechstKR
        If LCase(eintraege(e).klasse) <> LCase(eintraege(eRef).klasse) Then GoTo NaechstKR
        If LCase(eintraege(e).fach) <> LCase(eintraege(eRef).fach) Then GoTo NaechstKR
        If IstFixiert(eintraege(e).UrsprungsLehrer) Then GoTo NaechstKR
        If loesung(e) = g_lehrer(newLIdx).name Then
            g_lehrer(newLIdx).istWSt = g_lehrer(newLIdx).istWSt - eintraege(e).WertUV
            Dim rIdx As Long: rIdx = LehrerIndex(g_lehrer, g_nL, altLoesung(e))
            If rIdx > 0 Then g_lehrer(rIdx).istWSt = g_lehrer(rIdx).istWSt + eintraege(e).WertUV
            loesung(e) = altLoesung(e)
        End If
NaechstKR:
    Next e
End Sub

Sub VerbessereSA(eintraege() As tEintrag, nE As Long, _
                  loesung() As String, zeitLimit As Double)
    ' Brute-Force Lastausgleich:
    ' Fuer jeden ueberlasteten Lehrer: suche systematisch alle unterlasteten
    ' qualifizierten Ersatzlehrer. Uebertrage wenn Lastbalance sich verbessert
    ' und Konsistenz (gleiche Klasse+Fach = gleicher Lehrer) gewahrt bleibt.
    ' Wiederhole bis keine Verbesserung mehr moeglich oder Zeitlimit.

    Dim startTime As Double: startTime = Timer
    Dim i As Long, e As Long, eK As Long
    Dim li As Long, lj As Long
    Dim improved As Boolean
    Dim pass As Long

    ' IstWSt aus aktueller Loesung aufbauen
    For i = 1 To g_nL: g_lehrer(i).istWSt = 0: Next i
    For e = 1 To nE
        li = LehrerIndex(g_lehrer, g_nL, loesung(e))
        If li > 0 Then g_lehrer(li).istWSt = g_lehrer(li).istWSt + eintraege(e).WertUV
    Next e

    Do While Timer - startTime < zeitLimit
        pass = pass + 1
        improved = False

        ' Jeden ueberlasteten Lehrer durchgehen
        For li = 1 To g_nL
            If g_lehrer(li).sollWst <= 0 Then GoTo NxtLi
            If g_lehrer(li).istWSt <= g_lehrer(li).sollWst Then GoTo NxtLi

            ' Jeden Eintrag dieses Lehrers durchgehen
            For e = 1 To nE
                If loesung(e) <> g_lehrer(li).name Then GoTo NxtE
                If IstFixiert(eintraege(e).UrsprungsLehrer) Then GoTo NxtE

                ' Alle Eintraege mit gleicher Klasse+Fach sammeln
                ' (muessen alle zusammen uebertragen werden)
                Dim grpAnz As Long: grpAnz = 0
                Dim grpIdx(1 To 100) As Long
                Dim grpWert As Double: grpWert = 0
                For eK = 1 To nE
                    If LCase(eintraege(eK).klasse) <> LCase(eintraege(e).klasse) Then GoTo NxtEK
                    If LCase(eintraege(eK).fach) <> LCase(eintraege(e).fach) Then GoTo NxtEK
                    If IstFixiert(eintraege(eK).UrsprungsLehrer) Then GoTo NxtEK
                    grpAnz = grpAnz + 1
                    grpIdx(grpAnz) = eK
                    grpWert = grpWert + eintraege(eK).WertUV
NxtEK:
                Next eK

                ' Jeden unterlasteten qualifizierten Ersatz suchen
                For lj = 1 To g_nL
                    If lj = li Then GoTo NxtLj
                    If g_lehrer(lj).sollWst <= 0 Then GoTo NxtLj
                    ' Qualifikation pruefen
                    If Not KannFach(g_lehrer(lj), eintraege(e).fach, eintraege(e).IstOberstufe) Then GoTo NxtLj

                    ' Pruefen ob Uebertragung die Lastbalance verbessert:
                    ' Ueberlast_alt = IstWSt_li - SollWSt_li
                    ' Unterauslastung_alt = SollWSt_lj - IstWSt_lj
                    ' Nach Uebertragung:
                    ' IstWSt_li_neu = IstWSt_li - grpWert
                    ' IstWSt_lj_neu = IstWSt_lj + grpWert
                    ' Verbesserung wenn: |Abweichung_neu| < |Abweichung_alt|
                    Dim ueAlt As Double: ueAlt = g_lehrer(li).istWSt - g_lehrer(li).sollWst
                    Dim unAlt As Double: unAlt = g_lehrer(lj).sollWst - g_lehrer(lj).istWSt
                    Dim liNeu As Double: liNeu = g_lehrer(li).istWSt - grpWert
                    Dim ljNeu As Double: ljNeu = g_lehrer(lj).istWSt + grpWert

                    ' Verbesserung: Gesamtabweichung muss sinken
                    Dim abwAlt As Double: abwAlt = Abs(g_lehrer(li).istWSt - g_lehrer(li).sollWst) + Abs(g_lehrer(lj).istWSt - g_lehrer(lj).sollWst)
                    Dim abwNeu As Double: abwNeu = Abs(liNeu - g_lehrer(li).sollWst) + Abs(ljNeu - g_lehrer(lj).sollWst)
                    If abwNeu >= abwAlt Then GoTo NxtLj

                    ' Konsistenz pruefen: alle Gruppeneintraege muessen uebertragbar sein
                    Dim konsOK As Boolean: konsOK = True
                    Dim gk As Long
                    For gk = 1 To grpAnz
                        If Not KannFach(g_lehrer(lj), eintraege(grpIdx(gk)).fach, eintraege(grpIdx(gk)).IstOberstufe) Then
                            konsOK = False: Exit For
                        End If
                        ' Pruefen ob andere fixierte Eintraege mit gleicher Klasse+Fach den Lehrer vorgeben
                        Dim eF As Long
                        For eF = 1 To nE
                            If LCase(eintraege(eF).klasse) <> LCase(eintraege(grpIdx(gk)).klasse) Then GoTo NxtEF
                            If LCase(eintraege(eF).fach) <> LCase(eintraege(grpIdx(gk)).fach) Then GoTo NxtEF
                            If IstFixiert(eintraege(eF).UrsprungsLehrer) Then
                                If LCase(eintraege(eF).UrsprungsLehrer) <> LCase(g_lehrer(lj).name) Then
                                    konsOK = False: Exit For
                                End If
                            End If
NxtEF:
                        Next eF
                        If Not konsOK Then Exit For
                    Next gk
                    If Not konsOK Then GoTo NxtLj

                    ' Uebertragung durchfuehren
                    g_lehrer(li).istWSt = g_lehrer(li).istWSt - grpWert
                    g_lehrer(lj).istWSt = g_lehrer(lj).istWSt + grpWert
                    For gk = 1 To grpAnz
                        loesung(grpIdx(gk)) = g_lehrer(lj).name
                    Next gk
                    improved = True
                    GoTo NxtE  ' naechsten Eintrag von li suchen

NxtLj:
                Next lj
NxtE:
            Next e
NxtLi:
        Next li

        ' Wenn keine Verbesserung: fertig
        If Not improved Then Exit Do

    Loop

    ' Konsistenz-Nachbearbeitung (Sicherheit)
    Dim ePost As Long, ePost2 As Long
    For ePost = 1 To nE
        If IstFixiert(eintraege(ePost).UrsprungsLehrer) Then GoTo PostWeiterV
        If loesung(ePost) = "" Or loesung(ePost) = "?" Then GoTo PostWeiterV
        For ePost2 = ePost + 1 To nE
            If IstFixiert(eintraege(ePost2).UrsprungsLehrer) Then GoTo PostWeiter2V
            If LCase(eintraege(ePost2).klasse) <> LCase(eintraege(ePost).klasse) Then GoTo PostWeiter2V
            If LCase(eintraege(ePost2).fach) <> LCase(eintraege(ePost).fach) Then GoTo PostWeiter2V
            loesung(ePost2) = loesung(ePost)
PostWeiter2V:
        Next ePost2
PostWeiterV:
    Next ePost

    ' IstWSt neu aufbauen
    For i = 1 To g_nL: g_lehrer(i).istWSt = 0: Next i
    For i = 1 To nE
        If loesung(i) <> "?" And loesung(i) <> "" Then
            Dim lIdxPost As Long: lIdxPost = LehrerIndex(g_lehrer, g_nL, loesung(i))
            If lIdxPost > 0 Then g_lehrer(lIdxPost).istWSt = g_lehrer(lIdxPost).istWSt + eintraege(i).WertUV
        End If
    Next i
End Sub

Function EinzelScore(e As Long, lehrerName As String, eintraege() As tEintrag, nE As Long) As Double
    Dim lIdx As Long: lIdx = LehrerIndex(g_lehrer, g_nL, lehrerName)
    If lIdx = 0 Then EinzelScore = 0: Exit Function
    EinzelScore = BerechneScore(lIdx, e, eintraege, nE)
End Function

Function BerechneGesamtScore(eintraege() As tEintrag, nE As Long, loesung() As String) As Double
    Dim sc As Double: sc = 0
    Dim e As Long
    For e = 1 To nE
        Dim lIdx As Long: lIdx = LehrerIndex(g_lehrer, g_nL, loesung(e))
        If lIdx > 0 Then sc = sc + BerechneScore(lIdx, e, eintraege, nE)
    Next e
    BerechneGesamtScore = sc
End Function


Function IstFixiert(v As String) As Boolean
    IstFixiert = (v <> "" And v <> "?")
End Function

Function IstUrsprungsFixiert(e As tEintrag) As Boolean
    IstUrsprungsFixiert = (e.UrsprungsLehrer <> "" And e.UrsprungsLehrer <> "?")
End Function

' ============================================================
' DATEN EINLESEN
' ============================================================
Function LehrerEinlesen(ws As Worksheet, ByRef lehrer() As tLehrer) As Long
    Dim lastRow As Long
    Dim n As Long
    Dim i As Long
    Dim fi As Long
    Dim r As Long
    Dim fN As String

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    n = lastRow - 1
    If n <= 0 Then LehrerEinlesen = 0: Exit Function
    ReDim lehrer(1 To n)
    For i = 1 To n
        r = i + 1
        lehrer(i).name = Trim(CStr(ws.Cells(r, 1).Value))
        For fi = 1 To 16
            fN = Trim(CStr(ws.Cells(r, fi + 1).Value))
            ' Mehrfache Leerzeichen normalisieren (Untis exportiert manchmal "D  G1")
            Do While InStr(fN, "  ") > 0: fN = Replace(fN, "  ", " "): Loop
            lehrer(i).fach(fi) = fN
            lehrer(i).OberstufenOK(fi) = IstOberstufenFach(fN)
        Next fi
        lehrer(i).KlassenleiterIn = Trim(CStr(ws.Cells(r, 18).Value))
        If IsNumeric(ws.Cells(r, 19).Value) Then
            lehrer(i).sollWst = CDbl(ws.Cells(r, 19).Value)
        End If
        lehrer(i).istWSt = 0
    Next i
    LehrerEinlesen = n
End Function

Function IstOberstufenFach(fach As String) As Boolean
    Dim fu As String
    If fach = "" Then IstOberstufenFach = False: Exit Function
    fu = UCase(fach)
    If InStr(fu, "EF") > 0 Or InStr(fu, "Q1") > 0 Or InStr(fu, "Q2") > 0 Then
        IstOberstufenFach = True: Exit Function
    End If
    If FachGruppenPraefix(fach) <> "" Then IstOberstufenFach = True: Exit Function
    If FachLevelPraefix(fach) <> "" Then IstOberstufenFach = True: Exit Function
    If FachSuffixPraefix(fach) <> "" Then IstOberstufenFach = True: Exit Function
    IstOberstufenFach = False
End Function

Function WuenscheEinlesen(ws As Worksheet, ByRef wuensche() As tWunsch) As Long
    Dim lastRow As Long
    Dim n As Long
    Dim i As Long
    Dim r As Long

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    n = lastRow - 1
    If n <= 0 Then WuenscheEinlesen = 0: Exit Function
    ReDim wuensche(1 To n)
    For i = 1 To n
        r = i + 1
        wuensche(i).lehrerName = Trim(CStr(ws.Cells(r, 1).Value))
        wuensche(i).WunschKlasse = Trim(CStr(ws.Cells(r, 3).Value))
        If IsNumeric(ws.Cells(r, 4).Value) Then wuensche(i).WunschPrio = CLng(ws.Cells(r, 4).Value)
        wuensche(i).AntiKlasse = Trim(CStr(ws.Cells(r, 5).Value))
        If IsNumeric(ws.Cells(r, 6).Value) Then wuensche(i).AntiPrio = CLng(ws.Cells(r, 6).Value)
    Next i
    WuenscheEinlesen = n
End Function

Function EintraegeEinlesen(ws As Worksheet, ByRef eintraege() As tEintrag) As Long
    Dim cap As Long
    Dim n As Long
    Dim maxR As Long
    Dim aktK As String
    Dim r As Long
    Dim zF As String
    Dim zK As String
    Dim zL As String
    Dim kl As String

    cap = 300
    ReDim eintraege(1 To cap)
    n = 0
    maxR = ws.Cells(ws.rows.Count, 2).End(xlUp).row
    aktK = ""
    For r = 2 To maxR
        zF = Trim(CStr(ws.Cells(r, 2).Value))
        If zF = "" Then GoTo WeiterE
        zK = Trim(CStr(ws.Cells(r, 1).Value))
        If zK <> "" Then aktK = zK
        n = n + 1
        If n > cap Then cap = cap + 100: ReDim Preserve eintraege(1 To cap)
        eintraege(n).klasse = aktK
        ' Mehrfache Leerzeichen im Fachnamen normalisieren
        Do While InStr(zF, "  ") > 0: zF = Replace(zF, "  ", " "): Loop
        eintraege(n).fach = zF
        If IsNumeric(ws.Cells(r, 3).Value) Then
            eintraege(n).WSt = CDbl(ws.Cells(r, 3).Value)
        End If
        If IsNumeric(ws.Cells(r, 9).Value) Then
            eintraege(n).WertUV = CDbl(ws.Cells(r, 9).Value)
        Else
            eintraege(n).WertUV = 0
        End If
        ' Wert=0 bleibt 0 - kein Fallback auf WSt
        zL = Trim(CStr(ws.Cells(r, 4).Value))
        eintraege(n).lehrer = zL
        ' Spalte 6 = Fix-Lehrer (bewusst fixiert); Spalte 4 = Ergebnis (nicht fixiert)
        Dim zFix As String: zFix = Trim(CStr(ws.Cells(r, 6).Value))
        If zFix <> "" And zFix <> "?" Then
            eintraege(n).UrsprungsLehrer = zFix
            eintraege(n).lehrer = zFix  ' Fix-Lehrer auch als aktuellen Lehrer setzen
        Else
            eintraege(n).UrsprungsLehrer = "?"
        End If
        eintraege(n).zeile = r
        kl = LCase(aktK)
        eintraege(n).IstOberstufe = (InStr(kl, "ef") > 0 Or InStr(kl, "q1") > 0 Or InStr(kl, "q2") > 0)
WeiterE:
    Next r
    EintraegeEinlesen = n
End Function

Function KannFach(lehr As tLehrer, fach As String, oberstufe As Boolean) As Boolean
    ' Prueft ob Lehrer das Fach unterrichten kann.
    ' Regeln:
    '   1) Exakter Treffer: "D" = "D"
    '   2) G-Gruppe (Oberstufe): "Ge G1" passt zu "Ge G2", "Ge G3" etc.
    '   3) L-Gruppe (Leistungskurs-Level): "D L1" passt zu "D L2", "D L3" etc.
    Dim f As Long
    Dim p1 As String
    Dim p2 As String
    For f = 1 To 16
        If lehr.fach(f) = "" Then GoTo nF
        ' Exakter Treffer: Lehrer hat genau dieses Fach -> immer erlaubt
        If LCase(lehr.fach(f)) = LCase(fach) Then
            KannFach = True: Exit Function
        End If
        ' G-Gruppe: "Ge G1" <-> "Ge G2"
        p1 = FachGruppenPraefix(lehr.fach(f))
        p2 = FachGruppenPraefix(fach)
        If p1 <> "" And p2 <> "" And LCase(p1) = LCase(p2) Then
            KannFach = IIf(oberstufe, lehr.OberstufenOK(f), True): Exit Function
        End If
        ' L-Gruppe: "D L1" <-> "D L2" <-> "D L3"
        p1 = FachLevelPraefix(lehr.fach(f))
        p2 = FachLevelPraefix(fach)
        If p1 <> "" And p2 <> "" And LCase(p1) = LCase(p2) Then
            KannFach = IIf(oberstufe, lehr.OberstufenOK(f), True): Exit Function
        End If
        ' Allgemeines Suffix-Muster: "E Ver1" <-> "E Ver2", "E Zus1" <-> "E Zus2" etc.
        p1 = FachSuffixPraefix(lehr.fach(f))
        p2 = FachSuffixPraefix(fach)
        If p1 <> "" And p2 <> "" And LCase(p1) = LCase(p2) Then
            KannFach = IIf(oberstufe, lehr.OberstufenOK(f), True): Exit Function
        End If
nF:
    Next f
    KannFach = False
End Function

' Gibt den Praefix fuer L-Level-Faecher zurueck.
' "D L1"  -> "D"
' "EK L2" -> "EK"
' "Bio"   -> ""  (kein L-Muster)
Function FachLevelPraefix(fach As String) As String
    Dim pos As Long
    Dim suffix As String
    Dim k As Long
    FachLevelPraefix = ""
    Dim fachN As String: fachN = Trim(fach)
    Do While InStr(fachN, "  ") > 0: fachN = Join(Split(fachN, "  "), " "): Loop
    pos = InStr(fachN, " ")
    If pos < 2 Then Exit Function
    suffix = Trim(Mid(fachN, pos + 1))
    ' Muss "L" gefolgt von einer oder mehreren Ziffern sein
    If Len(suffix) < 2 Then Exit Function
    If UCase(left(suffix, 1)) <> "L" Then Exit Function
    For k = 2 To Len(suffix)
        If Mid(suffix, k, 1) < "0" Or Mid(suffix, k, 1) > "9" Then Exit Function
    Next k
    FachLevelPraefix = Trim(left(fachN, pos - 1))
End Function

Function FachSuffixPraefix(fach As String) As String
    Dim pos     As Long
    Dim suffix  As String
    Dim k       As Long
    Dim hasLetter As Boolean
    Dim hasDigit  As Boolean
    FachSuffixPraefix = ""
    Dim fachN As String: fachN = Trim(fach)
    Do While InStr(fachN, "  ") > 0: fachN = Join(Split(fachN, "  "), " "): Loop
    pos = InStr(fachN, " ")
    If pos < 2 Then Exit Function
    suffix = Trim(Mid(fachN, pos + 1))
    If Len(suffix) < 2 Then Exit Function
    If Len(suffix) = 1 Then Exit Function
    If Not (suffix Like "[A-Za-z]*") Then Exit Function
    If Not (suffix Like "*[0-9]*") Then Exit Function
    If UCase(left(suffix, 1)) = "G" And suffix Like "G[0-9]*" Then Exit Function
    If UCase(left(suffix, 1)) = "L" And suffix Like "L[0-9]*" Then Exit Function
    FachSuffixPraefix = Trim(left(fachN, pos - 1))
End Function

Function FachGruppenPraefix(fach As String) As String
    Dim pos As Long
    Dim suffix As String
    Dim k As Long
    FachGruppenPraefix = ""
    ' Normalisiere mehrfache Leerzeichen (z.B. "M  G1" -> "M G1")
    Dim fachN As String: fachN = Trim(fach)
    Do While InStr(fachN, "  ") > 0
        fachN = Join(Split(fachN, "  "), " ")
    Loop
    pos = InStr(fachN, " ")
    If pos < 2 Then Exit Function
    suffix = Mid(fachN, pos + 1)
    If Len(suffix) < 2 Then Exit Function
    If UCase(left(suffix, 1)) <> "G" Then Exit Function
    For k = 2 To Len(suffix)
        If Mid(suffix, k, 1) < "0" Or Mid(suffix, k, 1) > "9" Then Exit Function
    Next k
    FachGruppenPraefix = Trim(left(fachN, pos - 1))
End Function

Function LehrerIndex(lehrer() As tLehrer, nL As Long, name As String) As Long
    Dim i As Long
    For i = 1 To nL
        If lehrer(i).name = name Then LehrerIndex = i: Exit Function
    Next i
    LehrerIndex = 0
End Function

Function HatAntiPflicht(wuensche() As tWunsch, nW As Long, lName As String, klasse As String) As Boolean
    Dim w As Long
    For w = 1 To nW
        If wuensche(w).lehrerName = lName Then
            If wuensche(w).AntiKlasse = klasse And wuensche(w).AntiPrio = 3 Then
                HatAntiPflicht = True: Exit Function
            End If
        End If
    Next w
    HatAntiPflicht = False
End Function

Function WunschPrioFuer(wuensche() As tWunsch, nW As Long, lName As String, klasse As String, positiv As Boolean) As Long
    Dim w As Long
    Dim mx As Long
    mx = 0
    For w = 1 To nW
        If wuensche(w).lehrerName = lName Then
            If positiv Then
                If wuensche(w).WunschKlasse = klasse And wuensche(w).WunschPrio > mx Then mx = wuensche(w).WunschPrio
            Else
                If wuensche(w).AntiKlasse = klasse And wuensche(w).AntiPrio > mx Then mx = wuensche(w).AntiPrio
            End If
        End If
    Next w
    WunschPrioFuer = mx
End Function

' ============================================================
' UNTIS-IMPORTS
' ============================================================
Sub UntisImport_KlassenUV_nach_Klassen()
    Const SRC  As String = "KlassenUV"
    Const DEST As String = "Klassen"
    Dim wsSrc   As Worksheet
    Dim wsDest  As Worksheet
    Dim hRow    As Long
    Dim maxR    As Long
    Dim lHC     As Long
    Dim lDR     As Long
    Dim dRow    As Long
    Dim impCnt     As Long
    Dim skp     As Long
    Dim r       As Long
    Dim c       As Long
    Dim ki      As Long
    Dim cW      As Long
    Dim cWert   As Long
    Dim cL      As Long
    Dim cF      As Long
    Dim cK      As Long
    Dim ci      As Long
    Dim cvTmp   As String
    Dim kr      As String
    Dim fV      As String
    Dim lV      As String
    Dim kN      As String
    Dim oneK    As String
    Dim wV      As Double
    Dim wVje    As Double
    Dim wertVal As Double
    Dim klA()   As String
    Dim nKlassen As Long

    Set wsSrc = SheetByName(SRC)
    Set wsDest = SheetByName(DEST)
    If wsSrc Is Nothing Then MsgBox "Tabelle " & SRC & " fehlt!", vbCritical: Exit Sub
    If wsDest Is Nothing Then MsgBox "Tabelle " & DEST & " fehlt!", vbCritical: Exit Sub
    If MsgBox("Daten aus " & SRC & " in " & DEST & " uebertragen?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    hRow = 0
    maxR = wsSrc.Cells(wsSrc.rows.Count, 1).End(xlUp).row
    For r = 1 To WorksheetFunction.Min(10, maxR)
        For c = 1 To 30
            cvTmp = Trim(CStr(wsSrc.Cells(r, c).Value))
            If cvTmp = "U-Nr" Or cvTmp = "Wst" Or cvTmp = "Lehrer" Then hRow = r: Exit For
        Next c
        If hRow > 0 Then Exit For
    Next r
    If hRow = 0 Then MsgBox "Keine Header-Zeile!", vbCritical: Exit Sub

    lHC = wsSrc.Cells(hRow, wsSrc.Columns.Count).End(xlToLeft).Column
    cW = 0: cWert = 0: cL = 0: cF = 0: cK = 0: ci = 0
    For c = 1 To lHC
        Select Case Trim(CStr(wsSrc.Cells(hRow, c).Value))
            Case "Wst":        cW = c
            Case "Wert=", "Wert =", "Wert": cWert = c
            Case "Lehrer":     cL = c
            Case "Fach":       cF = c
            Case "Klasse(n)":  cK = c
            Case "Ignore (i)": ci = c
        End Select
    Next c
    If cW = 0 Or cF = 0 Or cK = 0 Then MsgBox "Pflicht-Spalten fehlen!", vbCritical: Exit Sub

    lDR = wsDest.Cells(wsDest.rows.Count, 1).End(xlUp).row
    If lDR >= 2 Then wsDest.rows("2:" & lDR).Delete Shift:=xlUp
    ' Header immer aktualisieren (neue Spalten Fix/L3/L4)
    wsDest.Cells(1, 1).Value = "Klasse":      wsDest.Cells(1, 2).Value = "Fach"
    wsDest.Cells(1, 3).Value = "WSt":         wsDest.Cells(1, 4).Value = "Lehrer (L1)"
    wsDest.Cells(1, 5).Value = "Lehrer (L2)": wsDest.Cells(1, 6).Value = "Fix"
    wsDest.Cells(1, 7).Value = "Lehrer (L3)": wsDest.Cells(1, 8).Value = "Lehrer (L4)"
    wsDest.Cells(1, 9).Value = "Wert"
    With wsDest.Cells(1, 9)
        .Font.Bold = True: .Interior.Color = RGB(68, 114, 196): .Font.Color = RGB(255, 255, 255)
    End With
    With wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(1, 5))
        .Font.Bold = True: .Interior.Color = RGB(68, 114, 196): .Font.Color = RGB(255, 255, 255)
    End With
    ' Fix-Spalte gruen markieren
    With wsDest.Range(wsDest.Cells(1, 6), wsDest.Cells(1, 6))
        .Font.Bold = True: .Interior.Color = RGB(0, 128, 0): .Font.Color = RGB(255, 255, 255)
    End With
    ' L3/L4 heller blau
    With wsDest.Range(wsDest.Cells(1, 7), wsDest.Cells(1, 8))
        .Font.Bold = True: .Interior.Color = RGB(142, 169, 219): .Font.Color = RGB(255, 255, 255)
    End With

    dRow = 2: impCnt = 0: skp = 0
    For r = hRow + 1 To maxR
        If ci > 0 Then
            If Trim(CStr(wsSrc.Cells(r, ci).Value)) = "i" Then skp = skp + 1: GoTo NZ
        End If
        kr = Trim(CStr(wsSrc.Cells(r, cK).Value))
        If kr = "" Then skp = skp + 1: GoTo NZ
        fV = Trim(CStr(wsSrc.Cells(r, cF).Value))
        wV = 0
        ' Wst immer aus Wst-Spalte (nicht Wert=)
        If IsNumeric(wsSrc.Cells(r, cW).Value) Then
            wV = CDbl(wsSrc.Cells(r, cW).Value)
        End If
        lV = ""
        If cL > 0 Then lV = Trim(CStr(wsSrc.Cells(r, cL).Value))
        If lV = "" Then lV = "?"
        wertVal = 0  ' Reset bei jedem Durchlauf (Dim ist ausserhalb der Schleife)
        If cWert > 0 Then
            If IsNumeric(wsSrc.Cells(r, cWert).Value) Then
                wertVal = CDbl(wsSrc.Cells(r, cWert).Value)
            End If
        End If
        ' Klassen trennen: NUR Leerzeichen und Semikolon als Trennzeichen.
        ' Komma-getrennte Ausdruecke wie "05A,05B,AG7" werden als EINE Klasse behandelt.
        kN = Replace(kr, ";", " ")
        Do While InStr(kN, "  ") > 0: kN = Replace(kN, "  ", " "): Loop
        klA = Split(Trim(kN), " ")
        nKlassen = UBound(klA) + 1
        ' WSt anteilig aufteilen nur bei Leerzeichen-getrennten Klassen
        If nKlassen > 1 And wV > 0 Then
            wVje = wV / nKlassen
        Else
            wVje = wV
        End If
        For ki = 0 To UBound(klA)
            oneK = Trim(klA(ki))
            If oneK = "" Then GoTo NK
            wsDest.Cells(dRow, 1).Value = oneK
            wsDest.Cells(dRow, 2).Value = fV
            wsDest.Cells(dRow, 3).Value = wVje
            wsDest.Cells(dRow, 4).Value = "?"   ' L1 leer (wird durch Optimierung befuellt)
            wsDest.Cells(dRow, 5).Value = "?"   ' L2 leer
            wsDest.Cells(dRow, 6).Value = lV    ' Fix-Lehrer aus KlassenUV
            wsDest.Cells(dRow, 7).Value = "?"   ' L3 leer
            wsDest.Cells(dRow, 8).Value = "?"   ' L4 leer
            wsDest.Cells(dRow, 9).Value = IIf(nKlassen > 1 And wertVal > 0, wertVal / nKlassen, wertVal)
            If dRow Mod 2 = 0 Then
                wsDest.Range(wsDest.Cells(dRow, 1), wsDest.Cells(dRow, 5)).Interior.Color = RGB(242, 242, 242)
            End If
            dRow = dRow + 1: impCnt = impCnt + 1
NK:
        Next ki
NZ:
    Next r
    If impCnt > 1 Then
        wsDest.Range(wsDest.Cells(2, 1), wsDest.Cells(dRow - 1, 9)).Sort _
            Key1:=wsDest.Cells(2, 1), Order1:=xlAscending, _
            Key2:=wsDest.Cells(2, 2), Order2:=xlAscending, Header:=xlNo
    End If
    wsDest.Columns(1).ColumnWidth = 12: wsDest.Columns(2).ColumnWidth = 14
    wsDest.Columns(3).ColumnWidth = 8:  wsDest.Columns(4).ColumnWidth = 14
    wsDest.Columns(5).ColumnWidth = 14: wsDest.Columns(6).ColumnWidth = 14
    wsDest.Columns(7).ColumnWidth = 14: wsDest.Columns(8).ColumnWidth = 14
    wsDest.Columns(9).ColumnWidth = 8
    MsgBox "Import: " & impCnt & " Zeilen." & IIf(skp > 0, " (" & skp & " uebersprungen)", ""), vbInformation
    wsDest.Activate
End Sub

Sub UntisImport_KlassenUV_nach_Lehrerliste()
    Const SRC  As String = "KlassenUV"
    Const DEST As String = "Lehrerliste"
    Const MF   As Long = 16
    Dim wsSrc   As Worksheet
    Dim wsDest  As Worksheet
    Dim hRow    As Long
    Dim maxR    As Long
    Dim lHC     As Long
    Dim lDR     As Long
    Dim dRow    As Long
    Dim r       As Long
    Dim c       As Long
    Dim fi      As Long
    Dim nL2     As Long
    Dim li      As Long
    Dim cL      As Long
    Dim cF      As Long
    Dim cK      As Long
    Dim ci      As Long
    Dim impCnt     As Long
    Dim neu     As Long
    Dim skp     As Long
    Dim zR      As Long
    Dim fS      As Long
    Dim cvTmp2  As String
    Dim lr      As String
    Dim fR      As String
    Dim lnm     As String
    Dim ow      As String
    Dim msg     As String
    Dim fD      As Boolean
    Dim lN(500)          As String
    Dim lF(500, 1 To 16) As String
    Dim lZ(500)          As Long

    Set wsSrc = SheetByName(SRC)
    Set wsDest = SheetByName(DEST)
    If wsSrc Is Nothing Then MsgBox "Tabelle " & SRC & " fehlt!", vbCritical: Exit Sub
    If wsDest Is Nothing Then MsgBox "Tabelle " & DEST & " fehlt!", vbCritical: Exit Sub
    If MsgBox("Faecher in Lehrerliste eintragen?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    hRow = 0
    maxR = wsSrc.Cells(wsSrc.rows.Count, 1).End(xlUp).row
    For r = 1 To WorksheetFunction.Min(10, maxR)
        For c = 1 To 30
            cvTmp2 = Trim(CStr(wsSrc.Cells(r, c).Value))
            If cvTmp2 = "U-Nr" Or cvTmp2 = "Wst" Or cvTmp2 = "Lehrer" Then hRow = r: Exit For
        Next c
        If hRow > 0 Then Exit For
    Next r
    If hRow = 0 Then MsgBox "Keine Header-Zeile!", vbCritical: Exit Sub

    lHC = wsSrc.Cells(hRow, wsSrc.Columns.Count).End(xlToLeft).Column
    cL = 0: cF = 0: cK = 0: ci = 0
    For c = 1 To lHC
        Select Case Trim(CStr(wsSrc.Cells(hRow, c).Value))
            Case "Lehrer":     cL = c
            Case "Fach":       cF = c
            Case "Klasse(n)":  cK = c
            Case "Ignore (i)": ci = c
        End Select
    Next c
    If cL = 0 Or cF = 0 Then MsgBox "Spalten Lehrer/Fach fehlen!", vbCritical: Exit Sub

    Call LehrlisteHeader_Sicherstellen(wsDest)

    lDR = wsDest.Cells(wsDest.rows.Count, 1).End(xlUp).row
    nL2 = 0
    For dRow = 2 To lDR
        lnm = Trim(CStr(wsDest.Cells(dRow, 1).Value))
        If lnm = "" Then GoTo NLeh
        nL2 = nL2 + 1: lN(nL2) = lnm: lZ(nL2) = dRow
        For fi = 1 To MF
            lF(nL2, fi) = Trim(CStr(wsDest.Cells(dRow, fi + 1).Value))
        Next fi
NLeh:
    Next dRow

    impCnt = 0: neu = 0: skp = 0: ow = ""
    For r = hRow + 1 To maxR
        If ci > 0 Then
            If Trim(CStr(wsSrc.Cells(r, ci).Value)) = "i" Then skp = skp + 1: GoTo NUV
        End If
        lr = Trim(CStr(wsSrc.Cells(r, cL).Value))
        If lr = "" Then GoTo NUV
        fR = Trim(CStr(wsSrc.Cells(r, cF).Value))
        If fR = "" Then GoTo NUV

        li = 0
        For li = 1 To nL2
            If lN(li) = lr Then Exit For
        Next li
        If li = 0 Then
            nL2 = nL2 + 1: li = nL2: lN(li) = lr: lZ(li) = 0: neu = neu + 1
        End If

        fD = False: fS = 0
        For fi = 1 To MF
            If lF(li, fi) = fR Then fD = True: Exit For
            If lF(li, fi) = "" And fS = 0 Then fS = fi
        Next fi

        If Not fD Then
            If fS > 0 Then
                lF(li, fS) = fR: impCnt = impCnt + 1
            Else
                If InStr(ow, lr) = 0 Then
                    If ow <> "" Then ow = ow & ", "
                    ow = ow & lr & "(+" & fR & ")"
                End If
            End If
        End If
NUV:
    Next r

    Application.ScreenUpdating = False
    For li = 1 To nL2
        If lZ(li) = 0 Then
            zR = wsDest.Cells(wsDest.rows.Count, 1).End(xlUp).row + 1
            lZ(li) = zR: wsDest.Cells(zR, 1).Value = lN(li)
            If zR Mod 2 = 0 Then
                wsDest.Range(wsDest.Cells(zR, 1), wsDest.Cells(zR, 19)).Interior.Color = RGB(242, 242, 242)
            End If
        Else
            zR = lZ(li)
        End If
        For fi = 1 To MF
            wsDest.Cells(zR, fi + 1).Value = lF(li, fi)
        Next fi
    Next li
    Application.ScreenUpdating = True

    msg = "Import: " & impCnt & " Faecher, " & neu & " neue Lehrer."
    If skp > 0 Then msg = msg & vbCrLf & skp & " uebersprungen."
    If ow <> "" Then msg = msg & vbCrLf & "WARNUNG >12 Faecher: " & ow
    msg = msg & vbCrLf & "Klassenleitung + Soll-WSt bitte ergaenzen."
    MsgBox msg, IIf(ow <> "", vbExclamation, vbInformation), "Import Lehrerliste"
    wsDest.Activate
End Sub

Sub LehrlisteHeader_Sicherstellen(ws As Worksheet)
    Dim cH As Long
    Dim cW As Long
    Dim h  As Variant
    Dim c  As Long
    Dim ci As Long

    If Trim(CStr(ws.Cells(1, 1).Value)) <> "" Then Exit Sub
    cH = RGB(31, 73, 125): cW = RGB(255, 255, 255)
    h = Array("Name", "Fach 1", "Fach 2", "Fach 3", "Fach 4", "Fach 5", "Fach 6", _
              "Fach 7", "Fach 8", "Fach 9", "Fach 10", "Fach 11", "Fach 12", _
              "Fach 13", "Fach 14", "Fach 15", "Fach 16", _
              "Klassenleitung", "Soll-Anrechnung")
    For c = 0 To UBound(h)
        ws.Cells(1, c + 1).Value = h(c)
    Next c
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, 19))
        .Font.Bold = True: .Font.Color = cW
        .Interior.Color = cH: .HorizontalAlignment = xlCenter
    End With
    ws.rows(1).RowHeight = 20: ws.Columns(1).ColumnWidth = 12
    For ci = 2 To 17
        ws.Columns(ci).ColumnWidth = 10
    Next ci
    ws.Columns(18).ColumnWidth = 15: ws.Columns(19).ColumnWidth = 14
End Sub

' ============================================================
' GPU006-IMPORT: Fachkuerzel + Fachfaktor -> Tabelle "Liste"
' ============================================================
Sub GPU006_Import_Fachliste()
    Dim pfad        As String
    Dim iFile       As Integer
    Dim zeile       As String
    Dim teile()     As String
    Dim kuerzel     As String
    Dim faktorStr   As String
    Dim faktorVal   As Double
    Dim wsDest      As Worksheet
    Dim destName    As String
    Dim dRow        As Long
    Dim impCnt2     As Long
    Dim skp2        As Long
    Dim t           As Long

    destName = "Liste"

    ' Datei auswaehlen
    pfad = Application.GetOpenFilename( _
        FileFilter:="Untis GPU006 (*.TXT;*.txt;*.csv),*.TXT;*.txt;*.csv", _
        Title:="GPU006.TXT auswaehlen")
    If pfad = "False" Or pfad = "" Then Exit Sub

    ' Ziel-Sheet vorbereiten
    Set wsDest = SheetErzeugenOderLeeren(destName)

    ' Header
    wsDest.Cells(1, 1).Value = "Fach"
    wsDest.Cells(1, 2).Value = "Fachfaktor"
    With wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(1, 2))
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    wsDest.Columns(1).ColumnWidth = 18
    wsDest.Columns(2).ColumnWidth = 14

    dRow = 2: impCnt2 = 0: skp2 = 0

    iFile = FreeFile
    Open pfad For Input As #iFile
    Do While Not EOF(iFile)
        Line Input #iFile, zeile
        zeile = Trim(zeile)
        If zeile = "" Then GoTo NaechsteZeile

        ' Semikolon-getrennte Felder parsen
        ' Felder koennen in Anf" & Chr(252) & "hrungszeichen eingeschlossen sein
        teile = SplitCSVZeile(zeile, ";")

        If UBound(teile) < 14 Then skp2 = skp2 + 1: GoTo NaechsteZeile

        ' Spalte 0: Fachkuerzel (ohne Anf" & Chr(252) & "hrungszeichen)
        kuerzel = StripQuotes(teile(0))
        If kuerzel = "" Then skp2 = skp2 + 1: GoTo NaechsteZeile

        ' Spalte 14: Fachfaktor
        faktorStr = Trim(teile(14))
        If faktorStr = "" Then skp2 = skp2 + 1: GoTo NaechsteZeile

        ' Dezimaltrenner: Punkt -> Komma fuer deutsche Locale
        faktorStr = Replace(faktorStr, ".", Application.International(xlDecimalSeparator))
        If Not IsNumeric(faktorStr) Then skp2 = skp2 + 1: GoTo NaechsteZeile

        faktorVal = CDbl(faktorStr)

        ' Doppelte Eintraege ueberspringen
        Dim bereitsVorhanden As Boolean
        bereitsVorhanden = False
        For t = 2 To dRow - 1
            If LCase(Trim(CStr(wsDest.Cells(t, 1).Value))) = LCase(kuerzel) Then
                bereitsVorhanden = True: Exit For
            End If
        Next t
        If bereitsVorhanden Then skp2 = skp2 + 1: GoTo NaechsteZeile

        ' Schreiben
        wsDest.Cells(dRow, 1).Value = kuerzel
        wsDest.Cells(dRow, 2).Value = faktorVal
        wsDest.Cells(dRow, 2).NumberFormat = "0.00000"
        If dRow Mod 2 = 0 Then
            wsDest.Range(wsDest.Cells(dRow, 1), wsDest.Cells(dRow, 2)).Interior.Color = RGB(242, 242, 242)
        End If
        dRow = dRow + 1: impCnt2 = impCnt2 + 1
NaechsteZeile:
    Loop
    Close #iFile

    ' Nach Fachkuerzel sortieren
    If impCnt2 > 1 Then
        wsDest.Range(wsDest.Cells(2, 1), wsDest.Cells(dRow - 1, 2)).Sort _
            Key1:=wsDest.Cells(2, 1), Order1:=xlAscending, Header:=xlNo
    End If

    wsDest.Activate
    MsgBox "GPU006-Import abgeschlossen." & vbCrLf & _
           impCnt2 & " Faecher importiert." & _
           IIf(skp2 > 0, vbCrLf & skp2 & " Zeilen uebersprungen.", ""), _
           vbInformation, "GPU006 Import"
End Sub

' Hilfsfunktion: CSV-Zeile semikolongetrennt parsen
' Beruecksichtigt Felder in Anfuehrungszeichen
Function SplitCSVZeile(zeile As String, trenn As String) As String()
    Dim result() As String
    Dim n As Long
    Dim i As Long
    Dim inQuote As Boolean
    Dim aktuell As String
    Dim cH As String

    ReDim result(0 To 50)
    n = 0: inQuote = False: aktuell = ""

    For i = 1 To Len(zeile)
        cH = Mid(zeile, i, 1)
        If cH = """" Then
            inQuote = Not inQuote
        ElseIf cH = trenn And Not inQuote Then
            result(n) = aktuell: n = n + 1: aktuell = ""
            If n > UBound(result) Then ReDim Preserve result(0 To UBound(result) + 20)
        Else
            aktuell = aktuell & cH
        End If
    Next i
    result(n) = aktuell

    ReDim Preserve result(0 To n)
    SplitCSVZeile = result
End Function

' Entfernt fuehrende/abschliessende Anfuehrungszeichen
Function StripQuotes(s As String) As String
    Dim v As String
    v = Trim(s)
    If Len(v) >= 2 Then
        If left(v, 1) = """" And Right(v, 1) = """" Then
            v = Mid(v, 2, Len(v) - 2)
        End If
    End If
    StripQuotes = Trim(v)
End Function

' ============================================================
' FACH-LEHRER-LISTE: Aus Lehrerliste alle F" & Chr(228) & "cher extrahieren
' Erstellt Sheet "FachLehrer" mit je einer Zeile pro Fach,
' und den Lehrern, die dieses Fach unterrichten k" & Chr(246) & "nnen.
' ============================================================
Sub FachLehrerListe_Erstellen()
    Const DEST As String = "FachLehrer"
    Dim wsL     As Worksheet
    Dim wsDest  As Worksheet
    Dim lastRow As Long
    Dim r       As Long
    Dim fi      As Long
    Dim lName   As String
    Dim fach    As String
    Dim dRow    As Long
    Dim nF      As Long
    Dim cH      As Long
    Dim cSH     As Long

    ' Fach -> Liste der Lehrernamen (Dictionary-Simulation mit Arrays)
    Dim fachArr(500)    As String   ' alle gesammelten eindeutigen Faecher
    Dim lehrerArr(500)  As String   ' kommaseparierte Lehrerliste je Fach
    nF = 0

    Set wsL = SheetByName("Lehrerliste")
    If wsL Is Nothing Then MsgBox "Tabelle 'Lehrerliste' fehlt!", vbCritical: Exit Sub

    lastRow = wsL.Cells(wsL.rows.Count, 1).End(xlUp).row

    ' Alle Lehrer + F" & Chr(228) & "cher einlesen
    For r = 2 To lastRow
        lName = Trim(CStr(wsL.Cells(r, 1).Value))
        If lName = "" Then GoTo NaechsterLehrer
        For fi = 1 To 16
            fach = Trim(CStr(wsL.Cells(r, fi + 1).Value))
            If fach = "" Then GoTo NaechstesFach
            ' Fach in Liste suchen
            Dim fIdx As Long
            fIdx = -1
            Dim fk As Long
            For fk = 1 To nF
                If LCase(fachArr(fk)) = LCase(fach) Then fIdx = fk: Exit For
            Next fk
            If fIdx = -1 Then
                ' Neues Fach
                nF = nF + 1
                fIdx = nF
                fachArr(fIdx) = fach
                lehrerArr(fIdx) = lName
            Else
                ' Lehrer hinzuf" & Chr(252) & "gen falls noch nicht vorhanden
                If InStr("," & lehrerArr(fIdx) & ",", "," & lName & ",") = 0 Then
                    lehrerArr(fIdx) = lehrerArr(fIdx) & ", " & lName
                End If
            End If
NaechstesFach:
        Next fi
NaechsterLehrer:
    Next r

    If nF = 0 Then MsgBox "Keine F" & Chr(228) & "cher gefunden!", vbExclamation: Exit Sub

    ' F" & Chr(228) & "cher alphabetisch sortieren (Insertion Sort)
    Dim si As Long, sj As Long
    Dim tmpF As String, tmpL As String
    For si = 2 To nF
        tmpF = fachArr(si): tmpL = lehrerArr(si): sj = si - 1
        Do While sj >= 1 And LCase(fachArr(sj)) > LCase(tmpF)
            fachArr(sj + 1) = fachArr(sj): lehrerArr(sj + 1) = lehrerArr(sj): sj = sj - 1
        Loop
        fachArr(sj + 1) = tmpF: lehrerArr(sj + 1) = tmpL
    Next si

    ' Sheet erstellen
    Set wsDest = SheetErzeugenOderLeeren(DEST)
    cH = RGB(31, 73, 125): cSH = RGB(68, 114, 196)

    ' Header
    wsDest.Cells(1, 1).Value = "Fach"
    wsDest.Cells(1, 2).Value = "Lehrer (kommasepariert)"
    wsDest.Cells(1, 3).Value = "Anzahl"
    With wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(1, 3))
        .Font.Bold = True
        .Interior.Color = cH
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    wsDest.rows(1).RowHeight = 18

    ' Daten schreiben
    For fk = 1 To nF
        dRow = fk + 1
        wsDest.Cells(dRow, 1).Value = fachArr(fk)
        wsDest.Cells(dRow, 2).Value = lehrerArr(fk)
        ' Anzahl Lehrer zaehlen
        Dim anzahl As Long
        If lehrerArr(fk) = "" Then
            anzahl = 0
        Else
            anzahl = UBound(Split(lehrerArr(fk), ",")) + 1
        End If
        wsDest.Cells(dRow, 3).Value = anzahl
        ' Farbe
        If anzahl = 0 Then
            wsDest.Range(wsDest.Cells(dRow, 1), wsDest.Cells(dRow, 3)).Interior.Color = RGB(255, 199, 206)
        ElseIf anzahl = 1 Then
            wsDest.Range(wsDest.Cells(dRow, 1), wsDest.Cells(dRow, 3)).Interior.Color = RGB(255, 235, 156)
        ElseIf dRow Mod 2 = 0 Then
            wsDest.Range(wsDest.Cells(dRow, 1), wsDest.Cells(dRow, 3)).Interior.Color = RGB(242, 242, 242)
        End If
    Next fk

    ' Spaltenbreiten
    wsDest.Columns(1).ColumnWidth = 16
    wsDest.Columns(2).ColumnWidth = 60
    wsDest.Columns(3).ColumnWidth = 10

    ' Hinweiszeile
    Dim hinweisRow As Long
    hinweisRow = nF + 3
    wsDest.Cells(hinweisRow, 1).Value = "Hinweis: Lehrer in Spalte B erg" & Chr(228) & "nzen (kommasepariert), dann Makro 'FachLehrer_ZurueckInLehrerliste' ausf" & Chr(252) & "hren."
    wsDest.Cells(hinweisRow, 1).Font.Italic = True
    wsDest.Cells(hinweisRow, 1).Font.Color = RGB(128, 128, 128)
    wsDest.Range(wsDest.Cells(hinweisRow, 1), wsDest.Cells(hinweisRow, 3)).Merge

    wsDest.Activate
    wsDest.Cells(2, 1).Select

    MsgBox "FachLehrer-Liste erstellt: " & nF & " F" & Chr(228) & "cher." & vbCrLf & _
           "Rot = kein Lehrer, Gelb = nur 1 Lehrer." & vbCrLf & vbCrLf & _
           "Erg" & Chr(228) & "nzen Sie Lehrer in Spalte B und f" & Chr(252) & "hren Sie dann" & vbCrLf & _
           "'FachLehrer_ZurueckInLehrerliste' aus.", vbInformation, "FachLehrer-Liste"
End Sub

' ============================================================
' R" & Chr(220) & "CKSCHREIBEN: Manuelle Erg" & Chr(228) & "nzungen aus FachLehrer-Sheet
' zur" & Chr(252) & "ck in die Lehrerliste " & Chr(252) & "bertragen.
' Neue Lehrer werden als neue Zeilen angelegt.
' ============================================================
Sub FachLehrer_ZurueckInLehrerliste()
    Const SRC  As String = "FachLehrer"
    Const DEST As String = "Lehrerliste"
    Const MF   As Long = 16
    Dim wsSrc           As Worksheet
    Dim wsL             As Worksheet
    Dim lastSrc         As Long
    Dim lastL           As Long
    Dim r               As Long
    Dim fach            As String
    Dim lehrerStr       As String
    Dim lehrerArr()     As String
    Dim lName           As String
    Dim la              As Long
    Dim li              As Long
    Dim fi              As Long
    Dim fk              As Long
    Dim added           As Long
    Dim newLehrer       As Long
    Dim fachFnd         As Boolean
    Dim slotFrei        As Long
    Dim lRow            As Long
    Dim vorhandenesF    As String
    Dim anderesFach     As String
    Dim andGrp          As String
    Dim gPrx            As String
    Dim gruppeAbgedeckt As Boolean
    Dim msg2            As String

    Set wsSrc = SheetByName(SRC)
    Set wsL = SheetByName(DEST)
    If wsSrc Is Nothing Then MsgBox "Tabelle '" & SRC & "' fehlt! Zuerst FachLehrerListe_Erstellen ausf" & Chr(252) & "hren.", vbCritical: Exit Sub
    If wsL Is Nothing Then MsgBox "Tabelle 'Lehrerliste' fehlt!", vbCritical: Exit Sub

    If MsgBox("Lehrer aus '" & SRC & "' in Lehrerliste " & Chr(252) & "bertragen?" & vbCrLf & _
              "Neue Lehrer werden als neue Zeilen angelegt.", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    lastSrc = wsSrc.Cells(wsSrc.rows.Count, 1).End(xlUp).row
    lastL = wsL.Cells(wsL.rows.Count, 1).End(xlUp).row
    added = 0: newLehrer = 0

    Application.ScreenUpdating = False

    For r = 2 To lastSrc
        fach = Trim(CStr(wsSrc.Cells(r, 1).Value))
        ' Leerzeichen normalisieren (Untis-Export hat manchmal doppelte Leerzeichen z.B. "M  Ver1")
        Do While InStr(fach, "  ") > 0: fach = Replace(fach, "  ", " "): Loop
        If fach = "" Then GoTo NaechsteZeile2
        If left(fach, 7) = "Hinweis" Then GoTo NaechsteZeile2

        lehrerStr = Trim(CStr(wsSrc.Cells(r, 2).Value))
        If lehrerStr = "" Then GoTo NaechsteZeile2

        lehrerArr = Split(lehrerStr, ",")

        For la = 0 To UBound(lehrerArr)
            lName = Trim(lehrerArr(la))
            If lName = "" Then GoTo NaechsterLehrer2

            ' Schritt 1: Lehrer in Lehrerliste suchen
            lRow = 0
            For li = 2 To lastL
                If LCase(Trim(CStr(wsL.Cells(li, 1).Value))) = LCase(lName) Then
                    lRow = li: Exit For
                End If
            Next li

            ' Lehrer nicht gefunden -> neue Zeile anlegen
            If lRow = 0 Then
                lastL = lastL + 1
                lRow = lastL
                wsL.Cells(lRow, 1).Value = lName
                If lRow Mod 2 = 0 Then
                    wsL.Range(wsL.Cells(lRow, 1), wsL.Cells(lRow, 19)).Interior.Color = RGB(242, 242, 242)
                End If
                newLehrer = newLehrer + 1
            End If

            ' Schritt 2: Pruefen ob Fach bereits eingetragen
            fachFnd = False
            slotFrei = 0
            For fi = 1 To MF
                vorhandenesF = Trim(CStr(wsL.Cells(lRow, fi + 1).Value))
                If LCase(vorhandenesF) = LCase(fach) Then
                    fachFnd = True
                    Exit For
                End If
                If vorhandenesF = "" And slotFrei = 0 Then slotFrei = fi
            Next fi

            ' Fach bereits vorhanden: nichts tun
            If fachFnd Then GoTo NaechsterLehrer2

            ' Schritt 3: Pruefen ob Gruppenvertreter bereits eingetragen
            ' (z.B. "D G2" schon drin -> "D G1" nicht mehr noetig)
            gPrx = FachGruppenPraefix(fach)
            If gPrx = "" Then gPrx = FachLevelPraefix(fach)
            If gPrx = "" Then gPrx = FachSuffixPraefix(fach)

            gruppeAbgedeckt = False
            If gPrx <> "" Then
                For fi = 1 To MF
                    anderesFach = Trim(CStr(wsL.Cells(lRow, fi + 1).Value))
                    If anderesFach <> "" And LCase(anderesFach) <> LCase(fach) Then
                        andGrp = FachGruppenPraefix(anderesFach)
                        If andGrp = "" Then andGrp = FachLevelPraefix(anderesFach)
                        If andGrp = "" Then andGrp = FachSuffixPraefix(anderesFach)
                        If LCase(andGrp) = LCase(gPrx) Then
                            gruppeAbgedeckt = True: Exit For
                        End If
                    End If
                Next fi
            End If

            If gruppeAbgedeckt Then GoTo NaechsterLehrer2

            ' Schritt 4: Fach eintragen
            If slotFrei > 0 Then
                wsL.Cells(lRow, slotFrei + 1).Value = fach
                added = added + 1
            Else
                ' Kein Platz -> Warnung als Kommentar
                On Error Resume Next
                If wsL.Cells(lRow, MF + 1).Comment Is Nothing Then
                    wsL.Cells(lRow, MF + 1).AddComment "Kein Platz f" & Chr(252) & "r: " & fach
                Else
                    wsL.Cells(lRow, MF + 1).Comment.Text _
                        wsL.Cells(lRow, MF + 1).Comment.Text & ", " & fach
                End If
                On Error GoTo 0
            End If

NaechsterLehrer2:
        Next la
NaechsteZeile2:
    Next r

    Application.ScreenUpdating = True
    wsL.Activate

    msg2 = "R" & Chr(252) & "ckschreiben abgeschlossen." & vbCrLf & _
           added & " Fach-Eintr" & Chr(228) & "ge erg" & Chr(228) & "nzt." & vbCrLf & _
           newLehrer & " neue Lehrer angelegt."
    MsgBox msg2, vbInformation, "FachLehrer -> Lehrerliste"
End Sub

' ============================================================
' BUTTONS: Optimierungs-Startknopf in alle Tabellen legen
' ============================================================

' ============================================================
' SPALTENSUMMEN in Klassen, KlassenUV und Lehrerliste
' ============================================================
Sub SpaltensummenEintragen()
    Dim ws As Worksheet
    Dim lastR As Long
    Dim cSumme As Long

    ' --- Tabelle Klassen: WSt (Sp.3) und Wert (Sp.9) ---
    Set ws = SheetByName("Klassen")
    If Not ws Is Nothing Then
        lastR = ws.Cells(ws.rows.Count, 2).End(xlUp).row
        If lastR >= 2 Then
            ' Summenzeile loeschen falls vorhanden
            If Trim(CStr(ws.Cells(lastR + 1, 1).Value)) = "SUMME" Then ws.rows(lastR + 1).Delete
            Dim sumRow As Long: sumRow = lastR + 1
            ws.Cells(sumRow, 1).Value = "SUMME"
            ws.Cells(sumRow, 1).Font.Bold = True
            ws.Cells(sumRow, 3).Formula = "=SUM(C2:C" & lastR & ")"
            ws.Cells(sumRow, 3).Font.Bold = True
            ws.Cells(sumRow, 3).NumberFormat = "0.00"
            ws.Cells(sumRow, 9).Formula = "=SUM(I2:I" & lastR & ")"
            ws.Cells(sumRow, 9).Font.Bold = True
            ws.Cells(sumRow, 9).NumberFormat = "0.00"
            ws.Range(ws.Cells(sumRow, 1), ws.Cells(sumRow, 9)).Interior.Color = RGB(68, 114, 196)
            ws.Range(ws.Cells(sumRow, 1), ws.Cells(sumRow, 9)).Font.Color = RGB(255, 255, 255)
        End If
    End If

    ' --- Tabelle KlassenUV: Wst und Wert= ---
    Set ws = SheetByName("KlassenUV")
    If Not ws Is Nothing Then
        ' Header-Zeile finden
        Dim hRowK As Long: hRowK = 0
        Dim r As Long, c As Long
        Dim cWstK As Long: cWstK = 0
        Dim cWertK As Long: cWertK = 0
        For r = 1 To 10
            For c = 1 To 30
                Dim hv As String: hv = Trim(CStr(ws.Cells(r, c).Value))
                If hv = "Wst" Then cWstK = c: hRowK = r
                If hv = "Wert=" Or hv = "Wert =" Or hv = "Wert" Then cWertK = c: hRowK = r
            Next c
            If hRowK > 0 And cWstK > 0 Then Exit For
        Next r
        If hRowK > 0 Then
            lastR = ws.Cells(ws.rows.Count, 1).End(xlUp).row
            Dim sumRowK As Long: sumRowK = lastR + 1
            ws.Cells(sumRowK, 1).Value = "SUMME"
            ws.Cells(sumRowK, 1).Font.Bold = True
            If cWstK > 0 Then
                ws.Cells(sumRowK, cWstK).Formula = "=SUM(" & ws.Cells(hRowK + 1, cWstK).Address(False, False) & ":" & ws.Cells(lastR, cWstK).Address(False, False) & ")"
                ws.Cells(sumRowK, cWstK).Font.Bold = True: ws.Cells(sumRowK, cWstK).NumberFormat = "0.00"
            End If
            If cWertK > 0 Then
                ws.Cells(sumRowK, cWertK).Formula = "=SUM(" & ws.Cells(hRowK + 1, cWertK).Address(False, False) & ":" & ws.Cells(lastR, cWertK).Address(False, False) & ")"
                ws.Cells(sumRowK, cWertK).Font.Bold = True: ws.Cells(sumRowK, cWertK).NumberFormat = "0.00"
            End If
        End If
    End If

    ' --- Tabelle Lehrerliste: Soll-Anrechnungen (Sp.19) ---
    Set ws = SheetByName("Lehrerliste")
    If Not ws Is Nothing Then
        lastR = ws.Cells(ws.rows.Count, 1).End(xlUp).row
        If lastR >= 2 Then
            Dim sumRowL As Long: sumRowL = lastR + 1
            ws.Cells(sumRowL, 1).Value = "SUMME"
            ws.Cells(sumRowL, 1).Font.Bold = True
            ws.Cells(sumRowL, 19).Formula = "=SUM(S2:S" & lastR & ")"
            ws.Cells(sumRowL, 19).Font.Bold = True
            ws.Cells(sumRowL, 19).NumberFormat = "0.00"
            ws.Range(ws.Cells(sumRowL, 1), ws.Cells(sumRowL, 19)).Interior.Color = RGB(242, 242, 242)
        End If
    End If

    MsgBox "Spaltensummen eingetragen.", vbInformation
End Sub

Sub OptimierungsKnoepfe_Erstellen()
    ' Legt in jede relevante Tabelle einen Startknopf oben rechts.
    ' Bestehende Knoepfe werden zuerst entfernt.
    Dim sheetNamen(6) As String
    Dim i             As Long
    Dim ws            As Worksheet
    Dim btn           As Object
    Dim btnName       As String
    Dim lastCol       As Long
    Dim platzCol      As Long

    sheetNamen(0) = "Lehrerliste"
    sheetNamen(1) = "Klassen"
    sheetNamen(2) = "Lehrerw" & Chr(252) & "nsche"
    sheetNamen(3) = "KlassenUV"
    sheetNamen(4) = "FachLehrer"
    sheetNamen(5) = "Liste"
    sheetNamen(6) = "Lehrerbelegung"

    For i = 0 To 6
        Set ws = SheetByName(sheetNamen(i))
        If ws Is Nothing Then GoTo NaechstesSheet

        ws.Activate

        ' Bestehende Optimierungs-Knoepfe entfernen
        Dim shp As Object
        For Each shp In ws.Shapes
            If shp.name Like "BtnOptimierung*" Then shp.Delete
        Next shp

        ' Knopf-Position: Zeile 1, rechts vom letzten benutzten Bereich
        lastCol = ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1
        platzCol = lastCol + 2
        Dim left As Double
        Dim top  As Double
        left = ws.Cells(1, platzCol).left
        top = ws.Cells(1, platzCol).top

        ' Button erstellen (FormControl)
        Set btn = ws.Buttons.Add(left, top, 160, 24)
        btn.name = "BtnOptimierung_" & ws.name
        btn.Caption = "â–¶  Optimierung starten"
        btn.OnAction = "UnterrichtsverteilungOptimieren"
        With btn.Font
            .name = "Calibri"
            .Size = 10
            .Bold = True
        End With

NaechstesSheet:
    Next i

    ' Verbesserungs-Knopf in Klassen
    Dim wsKl As Worksheet
    Set wsKl = SheetByName("Klassen")
    If Not wsKl Is Nothing Then
        wsKl.Activate
        For Each shp In wsKl.Shapes
            If shp.name Like "BtnVerbesserung*" Then shp.Delete
        Next shp
        Dim lastColV As Long: lastColV = wsKl.UsedRange.Columns.Count + wsKl.UsedRange.Column - 1
        Dim btnV As Object
        Set btnV = wsKl.Buttons.Add(wsKl.Cells(1, lastColV + 2).left, wsKl.Cells(1, 1).top, 180, 24)
        btnV.name = "BtnVerbesserung_Klassen"
        btnV.Caption = "Loesungen verbessern (L3/L4)"
        btnV.OnAction = "NachtraeglicheVerbesserung"
        With btnV.Font
            .name = "Calibri": .Size = 10: .Bold = True
        End With
    End If

    ' Spaltensummen-Knopf in Klassen
    Dim wsSum As Worksheet
    Set wsSum = SheetByName("Klassen")
    If Not wsSum Is Nothing Then
        wsSum.Activate
        Dim shpSm As Object
        For Each shpSm In wsSum.Shapes
            If shpSm.name Like "BtnSummen*" Then shpSm.Delete
        Next shpSm
        Dim lastColS As Long: lastColS = wsSum.UsedRange.Columns.Count + wsSum.UsedRange.Column - 1
        Dim btnS As Object
        Set btnS = wsSum.Buttons.Add(wsSum.Cells(1, lastColS + 2).left, wsSum.Cells(1, 1).top + 28, 160, 24)
        btnS.name = "BtnSummen_Klassen"
        btnS.Caption = "Spaltensummen eintragen"
        btnS.OnAction = "SpaltensummenEintragen"
        With btnS.Font
            .name = "Calibri": .Size = 10: .Bold = True
        End With
    End If

    ' Import-Knopf in KlassenUV und Klassen
    Dim importSheets(1) As String
    importSheets(0) = "KlassenUV"
    importSheets(1) = "Klassen"
    For i = 0 To 1
        Set ws = SheetByName(importSheets(i))
        If ws Is Nothing Then GoTo NaechstesImport
        ws.Activate
        ' Alte Import-Knoepfe entfernen
        For Each shp In ws.Shapes
            If shp.name Like "BtnImport*" Then shp.Delete
        Next shp
        ' Position: nach Optimierungs-Knopf
        Dim leftI As Double
        Dim topI  As Double
        leftI = ws.Cells(1, 1).left
        topI = ws.Cells(1, 1).top
        Dim btnExist As Object
        For Each btnExist In ws.Buttons
            If btnExist.name Like "BtnOptimierung*" Or btnExist.name Like "BtnReset*" Then
                If btnExist.left + btnExist.Width > leftI Then
                    leftI = btnExist.left + btnExist.Width + 8
                    topI = btnExist.top
                End If
            End If
        Next btnExist
        Dim btnI As Object
        Set btnI = ws.Buttons.Add(leftI, topI, 170, 24)
        btnI.name = "BtnImport_" & ws.name
        btnI.Caption = "Untis-Import " & ">" & " Klassen"
        btnI.OnAction = "UntisImport_KlassenUV_nach_Klassen"
        With btnI.Font
            .name = "Calibri"
            .Size = 10
            .Bold = False
        End With
NaechstesImport:
    Next i

    ' FachLehrer-Knoepfe in Lehrerliste und FachLehrer-Sheet
    Dim flSheets(1) As String
    flSheets(0) = "Lehrerliste"
    flSheets(1) = "FachLehrer"
    For i = 0 To 1
        Set ws = SheetByName(flSheets(i))
        If ws Is Nothing Then GoTo NaechstesFach
        ws.Activate
        ' Alte FachLehrer-Knoepfe entfernen
        For Each shp In ws.Shapes
            If shp.name Like "BtnFach*" Then shp.Delete
        Next shp
        ' Position: nach letztem bestehenden Knopf
        Dim leftF As Double
        Dim topF  As Double
        leftF = ws.Cells(1, 1).left
        topF = ws.Cells(1, 1).top
        Dim btnEx2 As Object
        For Each btnEx2 In ws.Buttons
            If btnEx2.top = topF Or topF = ws.Cells(1, 1).top Then
                If btnEx2.left + btnEx2.Width + 8 > leftF Then
                    leftF = btnEx2.left + btnEx2.Width + 8
                    topF = btnEx2.top
                End If
            End If
        Next btnEx2
        ' Knopf 1: FachLehrerListe_Erstellen
        Dim btnF1 As Object
        Set btnF1 = ws.Buttons.Add(leftF, topF, 160, 24)
        btnF1.name = "BtnFach1_" & ws.name
        btnF1.Caption = "FachLehrer-Liste erstellen"
        btnF1.OnAction = "FachLehrerListe_Erstellen"
        With btnF1.Font
            .name = "Calibri"
            .Size = 10
            .Bold = False
        End With
        ' Knopf 2: FachLehrer_ZurueckInLehrerliste
        Dim btnF2 As Object
        Set btnF2 = ws.Buttons.Add(leftF + 168, topF, 180, 24)
        btnF2.name = "BtnFach2_" & ws.name
        btnF2.Caption = "FachLehrer > Lehrerliste"
        btnF2.OnAction = "FachLehrer_ZurueckInLehrerliste"
        With btnF2.Font
            .name = "Calibri"
            .Size = 10
            .Bold = False
        End With
NaechstesFach:
    Next i

    ' Auch Diagnose-Sheets falls vorhanden
    Dim diagSheets(1) As String
    diagSheets(0) = "Diagnose1"
    diagSheets(1) = "Diagnose2"
    For i = 0 To 1
        Set ws = SheetByName(diagSheets(i))
        If ws Is Nothing Then GoTo NaechstesDiag
        ws.Activate
        For Each shp In ws.Shapes
            If shp.name Like "BtnOptimierung*" Then shp.Delete
        Next shp
        lastCol = ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1
        platzCol = lastCol + 2
        left = ws.Cells(1, platzCol).left
        top = ws.Cells(1, platzCol).top
        Set btn = ws.Buttons.Add(left, top, 160, 24)
        btn.name = "BtnOptimierung_" & ws.name
        btn.Caption = "â–¶  Optimierung starten"
        btn.OnAction = "UnterrichtsverteilungOptimieren"
        With btn.Font
            .name = "Calibri"
            .Size = 10
            .Bold = True
        End With
NaechstesDiag:
    Next i

        ' Reset-Knopf in Klassen-Sheet
    Dim wsKlassen As Worksheet
    Set wsKlassen = SheetByName("Klassen")
    If Not wsKlassen Is Nothing Then
        wsKlassen.Activate
        Dim shpR As Object
        For Each shpR In wsKlassen.Shapes
            If shpR.name Like "BtnReset*" Then shpR.Delete
        Next shpR
        Dim lastColR As Long
        Dim btnR As Object
        lastColR = wsKlassen.UsedRange.Columns.Count + wsKlassen.UsedRange.Column - 1
        Dim leftR As Double
        Dim topR  As Double
        leftR = wsKlassen.Cells(1, lastColR + 2).left
        topR = wsKlassen.Cells(1, lastColR + 2).top
        ' Reset-Knopf rechts neben Optimierungs-Knopf
        Dim btnO As Object
        For Each btnO In wsKlassen.Buttons
            If btnO.name Like "BtnOptimierung*" Then
                leftR = btnO.left + btnO.Width + 8
                topR = btnO.top
                Exit For
            End If
        Next btnO
        Set btnR = wsKlassen.Buttons.Add(leftR, topR, 140, 24)
        btnR.name = "BtnReset_Klassen"
        btnR.Caption = "Zuweisung zur" & Chr(252) & "cksetzen"
        btnR.OnAction = "ZuweisungZuruecksetzen"
        With btnR.Font
            .name = "Calibri"
            .Size = 10
            .Bold = False
        End With
    End If

    MsgBox "Startkn" & Chr(246) & "pfe erstellt." & vbCrLf & _
           "Klick auf 'â–¶  Optimierung starten' startet das Makro.", _
           vbInformation, "Kn" & Chr(246) & "pfe erstellt"

    ' Engpass-Analyse Button in Sheet KlassenUV
    Dim wsUVBtn As Worksheet: Set wsUVBtn = SheetByName("KlassenUV")
    If Not wsUVBtn Is Nothing Then
        Dim shpUV As Object
        For Each shpUV In wsUVBtn.Shapes
            If shpUV.name = "BtnEngpass" Then shpUV.Delete
        Next shpUV
        Dim btnEng As Object
        Set btnEng = wsUVBtn.Buttons.Add(10, 5, 180, 24)
        btnEng.name = "BtnEngpass"
        btnEng.Caption = "> Engpass-Analyse"
        btnEng.OnAction = "FachEngpassAnalyse"
        With btnEng.Font: .name = "Calibri": .Size = 10: .Bold = True: End With
    End If
End Sub

' Entfernt alle Optimierungs-Knoepfe wieder
Sub OptimierungsKnoepfe_Entfernen()
    Dim ws  As Object
    Dim shp As Object
    Dim cnt As Long
    cnt = 0
    For Each ws In ThisWorkbook.Sheets
        For Each shp In ws.Shapes
            If shp.name Like "BtnOptimierung*" Then
                shp.Delete
                cnt = cnt + 1
            End If
        Next shp
    Next ws
    MsgBox cnt & " Kn" & Chr(246) & "pfe entfernt.", vbInformation
End Sub

' ============================================================
' METHODE 2: SIMULATED ANNEALING
' Startet mit zufaelliger Zuweisung, verbessert durch Tausche.
' Immer terminierend, kein Python noetig.
' ============================================================
Sub Optimieren_SA()
    Dim wsL         As Worksheet
    Dim wsK         As Worksheet
    Dim wsW         As Worksheet
    Dim eintraege() As tEintrag
    Dim nE          As Long
    Dim loesung1()  As String
    Dim loesung2()  As String
    Dim istWSt1()   As Double
    Dim istWSt2()   As Double
    Dim i           As Long
    Dim e           As Long
    Dim tolVal      As Double

    On Error GoTo SAFehler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wsL = SheetByName("Lehrerliste")
    Set wsK = SheetByName("Klassen")
    Set wsW = SheetByName("Lehrerw" & Chr(252) & "nsche")
    If wsL Is Nothing Or wsK Is Nothing Or wsW Is Nothing Then
        MsgBox "Benoetigt: Lehrerliste, Klassen, Lehrerw" & Chr(252) & "nsche", vbCritical
        GoTo SACleanup
    End If

    ReDim g_lehrer(1 To 1)
    g_nL = LehrerEinlesen(wsL, g_lehrer)
    Call SchutzFlagsSetzen
    If g_nL = 0 Then MsgBox "Keine Lehrer!", vbCritical: GoTo SACleanup

    ' Globale Sperre zuruecksetzen (kein Carry-over von vorherigem Lauf)
    g_sperrIdx = 0
    g_sperrName = ""

    ReDim g_wuensche(1 To 1)
    g_nW = WuenscheEinlesen(wsW, g_wuensche)

    nE = EintraegeEinlesen(wsK, eintraege)
    If nE = 0 Then MsgBox "Keine Eintraege!", vbCritical: GoTo SACleanup

    tolVal = g_toleranz

    ' Fach-Pruefung (gleiche Logik wie Backtracking)
    Dim pruefMsg As String
    Dim pruefFehler As Long
    pruefMsg = ""
    pruefFehler = 0
    Dim pruefE As Long
    Dim pruefI As Long
    Dim pruefOK As Boolean
    For pruefE = 1 To nE
        If IstFixiert(eintraege(pruefE).UrsprungsLehrer) Then GoTo SANaechsterPruef
        pruefOK = False
        For pruefI = 1 To g_nL
            If KannFach(g_lehrer(pruefI), eintraege(pruefE).fach, eintraege(pruefE).IstOberstufe) Then
                pruefOK = True: Exit For
            End If
        Next pruefI
        If Not pruefOK Then
            pruefFehler = pruefFehler + 1
            If pruefFehler <= 10 Then pruefMsg = pruefMsg & "  - " & eintraege(pruefE).fach & " in " & eintraege(pruefE).klasse & vbCrLf
        End If
SANaechsterPruef:
    Next pruefE
    If pruefFehler > 0 Then
        Dim saWarn As String
        saWarn = "WARNUNG: " & pruefFehler & " Eintrag(e) ohne qualifizierten Lehrer:" & vbCrLf & pruefMsg
        If MsgBox(saWarn & vbCrLf & "Trotzdem fortfahren?", vbYesNo + vbExclamation) = vbNo Then GoTo SACleanup
    End If

    ' ---- LOESUNG 1: Backtracking als Startloesung fuer SA ----
    ReDim loesung1(1 To nE)
    ReDim istWSt1(1 To g_nL)
    ' Backtracking-Startloesung mit Zeitlimit (max. 30 Sek.)
    Dim saStartLoes() As String
    ReDim saStartLoes(1 To nE)
    Dim saStartDom() As String
    ReDim saStartDom(1 To nE)
    g_sperrIdx = 0: g_sperrName = ""
    g_btMaxSek = g_zeitlimit: g_btStartTime = Timer
    Call ResetIstWSt(eintraege, nE)
    Call InitDomains(eintraege, nE, saStartDom)
    Dim btOK As Boolean
    btOK = BacktrackingMitCP(eintraege, nE, saStartDom, saStartLoes)
    If Not btOK Then
        g_toleranz = 9999
        g_btStartTime = Timer
        Call ResetIstWSt(eintraege, nE)
        Call InitDomains(eintraege, nE, saStartDom)
        btOK = BacktrackingMitCP(eintraege, nE, saStartDom, saStartLoes)
        g_toleranz = tolVal
    End If
    g_btMaxSek = 0  ' Zeitlimit wieder deaktivieren
    ' saStartLoes in loesung1 kopieren, dann SA mit hatStartLoes=True aufrufen
    ' Falls BT kein Ergebnis: SA startet ohne Vorgabe (eigene Greedy-Phase)
    Dim hatBTLoes1 As Boolean: hatBTLoes1 = False
    For e = 1 To nE
        loesung1(e) = saStartLoes(e)
        If saStartLoes(e) <> "" And saStartLoes(e) <> "?" Then hatBTLoes1 = True
    Next e
    Call SA_Solve(eintraege, nE, loesung1, tolVal, False, 0, "", hatBTLoes1)
    For i = 1 To g_nL: istWSt1(i) = g_lehrer(i).istWSt: Next i

    ' ---- LOESUNG 2: ersten Eintrag anders besetzen ----
    ReDim loesung2(1 To nE)
    ReDim istWSt2(1 To g_nL)
    ' Sperr-Eintrag: ersten nicht-fixierten Eintrag suchen
    Dim saSpIdx As Long: saSpIdx = 0
    Dim saSpNm  As String: saSpNm = ""
    For e = 1 To nE
        If Not IstFixiert(eintraege(e).UrsprungsLehrer) Then
            If loesung1(e) <> "" And loesung1(e) <> "?" Then
                saSpIdx = e: saSpNm = loesung1(e): Exit For
            End If
        End If
    Next e

    ' Wenn alle Eintraege fixiert: Loesung 2 = Loesung 1 (kein zweiter Lauf)
    If saSpIdx = 0 Then
        For e = 1 To nE: loesung2(e) = loesung1(e): Next e
        For i = 1 To g_nL: istWSt2(i) = istWSt1(i): Next i
    Else
        ' Backtracking-Startloesung fuer Loesung 2 (mit Sperre, Zeitlimit 30 Sek.)
        Dim saStartLoes2() As String
        ReDim saStartLoes2(1 To nE)
        Dim saStartDom2() As String
        ReDim saStartDom2(1 To nE)
        g_sperrIdx = saSpIdx: g_sperrName = saSpNm
        g_btMaxSek = g_zeitlimit: g_btStartTime = Timer
        Call ResetIstWSt(eintraege, nE)
        Call InitDomains(eintraege, nE, saStartDom2)
        Dim btOK2 As Boolean
        btOK2 = BacktrackingMitCP(eintraege, nE, saStartDom2, saStartLoes2)
        If Not btOK2 Then
            g_toleranz = 9999
            g_btStartTime = Timer
            Call ResetIstWSt(eintraege, nE)
            Call InitDomains(eintraege, nE, saStartDom2)
            btOK2 = BacktrackingMitCP(eintraege, nE, saStartDom2, saStartLoes2)
            g_toleranz = tolVal
        End If
        g_btMaxSek = 0
        g_sperrIdx = 0: g_sperrName = ""
        Dim hatBTLoes2 As Boolean: hatBTLoes2 = False
        For e = 1 To nE
            loesung2(e) = saStartLoes2(e)
            If saStartLoes2(e) <> "" And saStartLoes2(e) <> "?" Then hatBTLoes2 = True
        Next e
        Call SA_Solve(eintraege, nE, loesung2, tolVal, True, saSpIdx, saSpNm, hatBTLoes2)
        For i = 1 To g_nL: istWSt2(i) = g_lehrer(i).istWSt: Next i
    End If

    Call SchreibeLoesungen(wsK, eintraege, nE, loesung1, loesung2)
    Call BaueErgebnisSheets(loesung1, loesung2, istWSt1, istWSt2, eintraege, nE, "Simulated Annealing")

SACleanup:
    g_btMaxSek = 0
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Exit Sub

SAFehler:
    g_btMaxSek = 0
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    MsgBox "Fehler " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Tipp: Zeitlimit im Parameter-Sheet erhoehen.", _
           vbExclamation, "Optimierung abgebrochen"
End Sub

Sub SA_Solve(eintraege() As tEintrag, nE As Long, _
             ByRef loesung() As String, _
             tolVal As Double, _
             hatSperre As Boolean, _
             Optional sperrIdx As Long = 0, _
             Optional sperrName As String = "", _
             Optional hatStartLoes As Boolean = False)
    ' Simulated Annealing:
    ' Temperatur startet hoch (akzeptiert auch schlechte Zuege),
    ' kuehlt langsam ab (nur noch Verbesserungen).
    Dim e       As Long
    Dim i       As Long
    Dim kandidaten() As Long
    Dim nKand   As Long
    Dim zug     As Long

    ' ------ Startloesung ------
    ' ResetIstWSt setzt IstWSt=0 und zaehlt fixierte Eintraege bereits korrekt.
    ' Keine weitere Schleife noetig - das waere Doppelzaehlung.
    Call ResetIstWSt(eintraege, nE)

    ' Wenn Startloesung uebergeben (hatStartLoes=True): loesung() ist bereits gefuellt
    ' IstWSt aus uebergebenem loesung() aufbauen
    If hatStartLoes Then
        For e = 1 To nE
            If Not IstFixiert(eintraege(e).UrsprungsLehrer) Then
                If loesung(e) <> "" And loesung(e) <> "?" Then
                    For i = 1 To g_nL
                        If g_lehrer(i).name = loesung(e) Then
                            g_lehrer(i).istWSt = g_lehrer(i).istWSt + eintraege(e).WertUV: Exit For
                        End If
                    Next i
                End If
            End If
        Next e
        GoTo SAVerbesserung
    End If

    ' Greedy-Startloesung in MRV-Reihenfolge (Eintraege mit wenigsten Kandidaten zuerst)
    ' Verhindert dass spaetere Eintraege keine Kandidaten mehr haben
    Dim greedyOrd()  As Long
    Dim greedyKand() As Long
    Dim greedyN      As Long
    Dim greedyI      As Long
    Dim greedyJ      As Long
    Dim greedyTmp    As Long
    Dim greedyKandI  As Long
    ReDim greedyOrd(1 To nE)
    ReDim greedyKand(1 To nE)
    greedyN = 0
    For e = 1 To nE
        If Not IstFixiert(eintraege(e).UrsprungsLehrer) Then
            greedyN = greedyN + 1
            greedyOrd(greedyN) = e
            ' Kandidatenanzahl zaehlen
            greedyKandI = 0
            For i = 1 To g_nL
                If KannFach(g_lehrer(i), eintraege(e).fach, eintraege(e).IstOberstufe) Then
                    If Not HatAntiPflicht(g_wuensche, g_nW, g_lehrer(i).name, eintraege(e).klasse) Then
                        greedyKandI = greedyKandI + 1
                    End If
                End If
            Next i
            greedyKand(greedyN) = greedyKandI
        Else
            loesung(e) = eintraege(e).UrsprungsLehrer
        End If
    Next e
    ' Insertion-Sort nach Kandidatenanzahl aufsteigend (MRV)
    Dim greedyTmpK As Long
    For greedyI = 2 To greedyN
        greedyTmp = greedyOrd(greedyI)
        greedyTmpK = greedyKand(greedyI)
        greedyJ = greedyI - 1
        Do While greedyJ >= 1
            If greedyKand(greedyJ) > greedyTmpK Then
                greedyOrd(greedyJ + 1) = greedyOrd(greedyJ)
                greedyKand(greedyJ + 1) = greedyKand(greedyJ)
                greedyJ = greedyJ - 1
            Else
                Exit Do
            End If
        Loop
        greedyOrd(greedyJ + 1) = greedyTmp
        greedyKand(greedyJ + 1) = greedyTmpK
    Next greedyI
    ' Greedy in MRV-Reihenfolge belegen
    For greedyI = 1 To greedyN
        e = greedyOrd(greedyI)
        loesung(e) = SA_BesterKandidat(e, eintraege, nE, tolVal, hatSperre, sperrIdx, sperrName, loesung)
        If loesung(e) <> "?" Then
            For i = 1 To g_nL
                If g_lehrer(i).name = loesung(e) Then
                    g_lehrer(i).istWSt = g_lehrer(i).istWSt + eintraege(e).WertUV: Exit For
                End If
            Next i
        End If
    Next greedyI
    ' Zweiter Durchlauf: unbesetzte Eintraege ohne Kapazitaetspruefung belegen
    For greedyI = 1 To greedyN
        e = greedyOrd(greedyI)
        If loesung(e) = "?" Or loesung(e) = "" Then

            ' Besten Kandidaten ohne Kapazitaets-Einschraenkung suchen
            ' Konsistenz: Geschwister-Eintrag pruefen
            Dim zwangFB As String: zwangFB = ""
            Dim eFB As Long
            For eFB = 1 To nE
                If eFB <> e Then
                    If LCase(eintraege(eFB).klasse) = LCase(eintraege(e).klasse) Then
                        If LCase(eintraege(eFB).fach) = LCase(eintraege(e).fach) Then
                            If IstFixiert(eintraege(eFB).lehrer) Then
                                zwangFB = eintraege(eFB).lehrer: Exit For
                            End If
                            If loesung(eFB) <> "" And loesung(eFB) <> "?" Then
                                zwangFB = loesung(eFB): Exit For
                            End If
                        End If
                    End If
                End If
            Next eFB
            Dim bestFallback As Double: bestFallback = -1E+30
            Dim bestFallbackNm As String: bestFallbackNm = "?"
            ' Wenn Zwang gesetzt: nur diesen Lehrer nehmen (Kapazitaet egal)
            If zwangFB <> "" Then
                bestFallbackNm = zwangFB
            Else
                ' Erst MIT Kapazitaetspruefung suchen
                For i = 1 To g_nL
                    If hatSperre And e = sperrIdx And g_lehrer(i).name = sperrName Then GoTo FallbackNext
                    If Not KannFach(g_lehrer(i), eintraege(e).fach, eintraege(e).IstOberstufe) Then GoTo FallbackNext
                    If HatAntiPflicht(g_wuensche, g_nW, g_lehrer(i).name, eintraege(e).klasse) Then GoTo FallbackNext
                    If g_lehrer(i).istWSt + eintraege(e).WertUV > g_lehrer(i).sollWst + g_toleranz Then GoTo FallbackNext
                    Dim fbSc As Double: fbSc = BerechneScore(i, e, eintraege, nE)
                    If fbSc > bestFallback Then bestFallback = fbSc: bestFallbackNm = g_lehrer(i).name
FallbackNext:
                Next i
                ' Wenn immer noch kein Kandidat: OHNE Kapazitaetspruefung (Ueberlastung in Kauf nehmen)
                If bestFallbackNm = "?" Then
                    For i = 1 To g_nL
                        If hatSperre And e = sperrIdx And g_lehrer(i).name = sperrName Then GoTo FallbackNext2
                        If Not KannFach(g_lehrer(i), eintraege(e).fach, eintraege(e).IstOberstufe) Then GoTo FallbackNext2
                        If HatAntiPflicht(g_wuensche, g_nW, g_lehrer(i).name, eintraege(e).klasse) Then GoTo FallbackNext2
                        Dim fbSc2 As Double: fbSc2 = BerechneScore(i, e, eintraege, nE)
                        If fbSc2 > bestFallback Then bestFallback = fbSc2: bestFallbackNm = g_lehrer(i).name
FallbackNext2:
                    Next i
                End If
            End If
            If bestFallbackNm <> "?" Then
                loesung(e) = bestFallbackNm
                For i = 1 To g_nL
                    If g_lehrer(i).name = bestFallbackNm Then
                        g_lehrer(i).istWSt = g_lehrer(i).istWSt + eintraege(e).WertUV: Exit For
                    End If
                Next i
            End If
        End If
    Next greedyI

    ' Konsistenz wird nach der Verbesserungsphase als Nachbearbeitungsschritt behandelt

SAVerbesserung:
    ' ------ SA-Verbesserungsphase ------
    Dim temp        As Double
    Dim tempMin     As Double
    Dim kuehlung    As Double
    Dim iteration   As Long
    Dim maxIter     As Long
    Dim scoreCurr   As Double
    Dim scoreNew    As Double
    Dim delta       As Double
    Dim accept      As Double
    Dim e1          As Long
    Dim e2          As Long
    Dim altL1       As String
    Dim altL2       As String
    Dim neuL1       As String
    Dim neuL2       As String

    ' Parameter skalieren nach Problemgroesse
    Dim nOffen As Long: nOffen = 0
    For e = 1 To nE
        If Not IstFixiert(eintraege(e).UrsprungsLehrer) Then nOffen = nOffen + 1
    Next e

    ' Wenn alle fixiert: keine Verbesserungsphase noetig
    If nOffen = 0 Then GoTo SAFertig

    temp = 50
    tempMin = 0.1
    If nOffen <= 20 Then
        maxIter = 2000
        kuehlung = 0.98
    ElseIf nOffen <= 100 Then
        maxIter = 10000
        kuehlung = 0.995
    Else
        maxIter = 30000
        kuehlung = 0.998
    End If

    scoreCurr = SA_GesamtScore(loesung, eintraege, nE, tolVal)

    Randomize
    For iteration = 1 To maxIter
        ' Zufaelligen nicht-fixierten Eintrag waehlen
        Do
            e1 = Int(Rnd * nE) + 1
        Loop While IstFixiert(eintraege(e1).UrsprungsLehrer)

        If hatSperre And e1 = sperrIdx Then GoTo SAWeiter

        ' Zufaellig: entweder neuen Lehrer zuweisen (70%) oder zwei Eintraege tauschen (30%)
        If Rnd < 0.5 Then
            ' Neuen Lehrer fuer e1 waehlen
            altL1 = loesung(e1)
            neuL1 = SA_ZufallsKandidat(e1, eintraege, nE, tolVal, hatSperre, sperrIdx, sperrName, loesung)
            If neuL1 = altL1 Then GoTo SAWeiter
            If neuL1 = "?" Or neuL1 = "" Then GoTo SAWeiter  ' Kein Kandidat -> ueberspringen

            ' Delta-Score: nur e1 hat sich geaendert
            Dim scoreAlt1 As Double: scoreAlt1 = 0
            Dim scoreNeu1 As Double: scoreNeu1 = 0
            If altL1 <> "?" And altL1 <> "" Then
                Dim aIdx1 As Long: aIdx1 = SA_LehrerIdx(altL1)
                If aIdx1 > 0 Then scoreAlt1 = BerechneScore(aIdx1, e1, eintraege, nE)
                g_lehrer(aIdx1).istWSt = g_lehrer(aIdx1).istWSt - eintraege(e1).WertUV
            End If
            If neuL1 <> "?" And neuL1 <> "" Then
                Dim nIdx1 As Long: nIdx1 = SA_LehrerIdx(neuL1)
                If nIdx1 > 0 Then
                    g_lehrer(nIdx1).istWSt = g_lehrer(nIdx1).istWSt + eintraege(e1).WertUV
                    scoreNeu1 = BerechneScore(nIdx1, e1, eintraege, nE)
                End If
            End If
            loesung(e1) = neuL1
            delta = (scoreNeu1 - scoreAlt1)
            scoreNew = scoreCurr + delta

            If delta >= 0 Then
                ' Verbesserung: immer akzeptieren
                scoreCurr = scoreNew
                ' Konsistenz wird am Ende erzwungen
            ElseIf temp > 0.001 Then
                ' Verschlechterung: mit Wahrscheinlichkeit e^(delta/temp) akzeptieren
                accept = IIf(delta / temp < -700, 0, Exp(delta / temp))
                If Rnd < accept Then
                    scoreCurr = scoreNew
                    ' Konsistenz wird am Ende erzwungen
                Else
                    ' Rueckgaengig machen
                    loesung(e1) = altL1
                    If neuL1 <> "?" Then
                        For i = 1 To g_nL
                            If g_lehrer(i).name = neuL1 Then
                                g_lehrer(i).istWSt = g_lehrer(i).istWSt - eintraege(e1).WertUV
                                Exit For
                            End If
                        Next i
                    End If
                    If altL1 <> "?" Then
                        For i = 1 To g_nL
                            If g_lehrer(i).name = altL1 Then
                                g_lehrer(i).istWSt = g_lehrer(i).istWSt + eintraege(e1).WertUV
                                Exit For
                            End If
                        Next i
                    End If
                End If
            Else
                ' Kalt: nur Verbesserungen
                loesung(e1) = altL1
                If neuL1 <> "?" Then
                    For i = 1 To g_nL
                        If g_lehrer(i).name = neuL1 Then
                            g_lehrer(i).istWSt = g_lehrer(i).istWSt - eintraege(e1).WertUV
                            Exit For
                        End If
                    Next i
                End If
                If altL1 <> "?" Then
                    For i = 1 To g_nL
                        If g_lehrer(i).name = altL1 Then
                            g_lehrer(i).istWSt = g_lehrer(i).istWSt + eintraege(e1).WertUV
                            Exit For
                        End If
                    Next i
                End If
            End If
        Else
            ' Zwei Eintraege tauschen
            Dim e2try As Long: e2try = 0
            Dim tryCount As Long
            For tryCount = 1 To 10
                e2 = Int(Rnd * nE) + 1
                If e2 <> e1 And Not IstFixiert(eintraege(e2).lehrer) Then
                    If Not (hatSperre And e2 = sperrIdx) Then
                        ' Faecher muessen zueinander kompatibel sein
                        Dim lIdx1 As Long: lIdx1 = SA_LehrerIdx(loesung(e1))
                        Dim lIdx2 As Long: lIdx2 = SA_LehrerIdx(loesung(e2))
                        If lIdx1 > 0 And lIdx2 > 0 Then
                            If KannFach(g_lehrer(lIdx1), eintraege(e2).fach, eintraege(e2).IstOberstufe) Then
                                If KannFach(g_lehrer(lIdx2), eintraege(e1).fach, eintraege(e1).IstOberstufe) Then
                                    e2try = e2: Exit For
                                End If
                            End If
                        End If
                    End If
                End If
            Next tryCount
            If e2try = 0 Then GoTo SAWeiter

            altL1 = loesung(e1): altL2 = loesung(e2try)
            ' Delta-Score Tausch: e1 bekommt altL2, e2try bekommt altL1
            Dim sA1 As Double: sA1 = 0: Dim sA2 As Double: sA2 = 0
            Dim sN1 As Double: sN1 = 0: Dim sN2 As Double: sN2 = 0
            Dim iA1 As Long: iA1 = SA_LehrerIdx(altL1)
            Dim iA2 As Long: iA2 = SA_LehrerIdx(altL2)
            If iA1 > 0 Then sA1 = BerechneScore(iA1, e1, eintraege, nE)
            If iA2 > 0 Then sA2 = BerechneScore(iA2, e2try, eintraege, nE)
            If iA1 > 0 Then g_lehrer(iA1).istWSt = g_lehrer(iA1).istWSt - eintraege(e1).WertUV + eintraege(e2try).WertUV
            If iA2 > 0 Then g_lehrer(iA2).istWSt = g_lehrer(iA2).istWSt - eintraege(e2try).WertUV + eintraege(e1).WertUV
            loesung(e1) = altL2: loesung(e2try) = altL1
            If iA2 > 0 Then sN1 = BerechneScore(iA2, e1, eintraege, nE)
            If iA1 > 0 Then sN2 = BerechneScore(iA1, e2try, eintraege, nE)
            delta = (sN1 + sN2) - (sA1 + sA2)
            scoreNew = scoreCurr + delta

            If delta >= 0 Or (temp > 0.001 And Rnd < IIf(delta / temp < -700, 0, Exp(delta / temp))) Then
                scoreCurr = scoreNew
                ' Konsistenz wird am Ende erzwungen
            Else
                ' Tausch rueckgaengig
                loesung(e1) = altL1: loesung(e2try) = altL2
                If iA1 > 0 Then g_lehrer(iA1).istWSt = g_lehrer(iA1).istWSt + eintraege(e1).WertUV - eintraege(e2try).WertUV
                If iA2 > 0 Then g_lehrer(iA2).istWSt = g_lehrer(iA2).istWSt + eintraege(e2try).WertUV - eintraege(e1).WertUV
            End If
        End If

SAWeiter:
        temp = temp * kuehlung
        If temp < tempMin Then Exit For
        If iteration Mod 1000 = 0 Then
            Application.StatusBar = "SA Iteration " & iteration & "/" & maxIter & "  Score=" & Format(scoreCurr, "0")
            DoEvents
        End If
    Next iteration

SAFertig:
    ' ------ Konsistenz-Nachbearbeitung ------
    ' Gleiche Klasse+Fach muessen denselben Lehrer haben.
    ' Wir nehmen fuer jede Gruppe den Lehrer der am haeufigsten vorkommt
    ' (oder den ersten zugewiesenen). Kapazitaet ist egal.
    Dim ePost As Long
    Dim ePost2 As Long
    For ePost = 1 To nE
        If IstFixiert(eintraege(ePost).UrsprungsLehrer) Then GoTo PostWeiter
        If loesung(ePost) = "" Or loesung(ePost) = "?" Then GoTo PostWeiter
        For ePost2 = ePost + 1 To nE
            If IstFixiert(eintraege(ePost2).UrsprungsLehrer) Then GoTo PostWeiter2
            If LCase(eintraege(ePost2).klasse) <> LCase(eintraege(ePost).klasse) Then GoTo PostWeiter2
            If LCase(eintraege(ePost2).fach) <> LCase(eintraege(ePost).fach) Then GoTo PostWeiter2
            loesung(ePost2) = loesung(ePost)  ' Gleicher Lehrer, Kapazitaet egal
PostWeiter2:
        Next ePost2
PostWeiter:
    Next ePost

    ' IstWSt aus finaler Loesung neu aufbauen (nach Konsistenz-Korrektur)
    Call ResetIstWSt(eintraege, nE)
    For e = 1 To nE
        If Not IstFixiert(eintraege(e).UrsprungsLehrer) Then
            If loesung(e) <> "?" And loesung(e) <> "" Then
                For i = 1 To g_nL
                    If g_lehrer(i).name = loesung(e) Then
                        g_lehrer(i).istWSt = g_lehrer(i).istWSt + eintraege(e).WertUV: Exit For
                    End If
                Next i
            End If
        End If
    Next e
End Sub


Sub SA_KonsistenzDurchsetzen(e As Long, lName As String, _
                              eintraege() As tEintrag, nE As Long, _
                              loesung() As String)
    ' Setzt alle Geschwister-Eintraege (gleiche Klasse+Fach) auf denselben Lehrer.
    ' IstWSt wird dabei angepasst.
    Dim e2 As Long
    Dim i  As Long
    For e2 = 1 To nE
        If e2 <> e Then
            If LCase(eintraege(e2).klasse) = LCase(eintraege(e).klasse) Then
                If LCase(eintraege(e2).fach) = LCase(eintraege(e).fach) Then
                    If Not IstFixiert(eintraege(e2).lehrer) Then
                        ' IstWSt des alten Lehrers korrigieren
                        If loesung(e2) <> "" Then
                        If loesung(e2) <> "?" Then
                        If loesung(e2) <> lName Then
                            For i = 1 To g_nL
                                If g_lehrer(i).name = loesung(e2) Then
                                    g_lehrer(i).istWSt = g_lehrer(i).istWSt - eintraege(e2).WertUV
                                    Exit For
                                End If
                            Next i
                            ' IstWSt des neuen Lehrers erhoehen
                            For i = 1 To g_nL
                                If g_lehrer(i).name = lName Then
                                    g_lehrer(i).istWSt = g_lehrer(i).istWSt + eintraege(e2).WertUV
                                    Exit For
                                End If
                            Next i
                        End If: End If: End If
                        loesung(e2) = lName
                    End If
                End If
            End If
        End If
    Next e2
End Sub
Function SA_LehrerIdx(lName As String) As Long
    Dim i As Long
    If lName = "" Or lName = "?" Then SA_LehrerIdx = 0: Exit Function
    For i = 1 To g_nL
        If g_lehrer(i).name = lName Then SA_LehrerIdx = i: Exit Function
    Next i
    SA_LehrerIdx = 0
End Function

Function SA_BesterKandidat(e As Long, eintraege() As tEintrag, nE As Long, _
                            tolVal As Double, hatSperre As Boolean, _
                            sperrIdx As Long, sperrName As String, _
                            loesung() As String) As String
    ' Gibt den Lehrer mit dem hoechsten Score fuer Eintrag e zurueck.
    ' Konsistenz: wenn gleiche Klasse+Fach schon zugewiesen, denselben Lehrer erzwingen.
    Dim i       As Long
    Dim best    As Double
    Dim bestNm  As String
    Dim sc      As Double
    Dim e2      As Long
    Dim zwang   As String
    best = -1E+30
    bestNm = "?"

    ' Konsistenz-Check: gleiche Klasse+Fach bereits in loesung() belegt?
    zwang = ""
    For e2 = 1 To nE
        If e2 <> e Then
            If LCase(eintraege(e2).klasse) = LCase(eintraege(e).klasse) Then
                If LCase(eintraege(e2).fach) = LCase(eintraege(e).fach) Then
                    If IstFixiert(eintraege(e2).lehrer) Then
                        zwang = eintraege(e2).lehrer: Exit For
                    End If
                    If loesung(e2) <> "" And loesung(e2) <> "?" Then
                        zwang = loesung(e2): Exit For
                    End If
                End If
            End If
        End If
    Next e2
    ' Konsistenz wird am Ende in SA_Solve als Nachbearbeitungsschritt behandelt

    For i = 1 To g_nL
        If hatSperre And e = sperrIdx And g_lehrer(i).name = sperrName Then GoTo SABKNext
        If Not KannFach(g_lehrer(i), eintraege(e).fach, eintraege(e).IstOberstufe) Then GoTo SABKNext
        If HatAntiPflicht(g_wuensche, g_nW, g_lehrer(i).name, eintraege(e).klasse) Then GoTo SABKNext
        If tolVal < 9999 Then
            If g_lehrer(i).istWSt + eintraege(e).WertUV > g_lehrer(i).sollWst + tolVal Then GoTo SABKNext
        End If
        sc = BerechneScore(i, e, eintraege, nE)
        If sc > best Then best = sc: bestNm = g_lehrer(i).name
SABKNext:
    Next i
    ' Fallback: wenn kein Kandidat in Toleranz, nochmal ohne Kapazitaetsbeschraenkung
    If bestNm = "?" Then
        For i = 1 To g_nL
            If hatSperre And e = sperrIdx And g_lehrer(i).name = sperrName Then GoTo SABKNext2
            If Not KannFach(g_lehrer(i), eintraege(e).fach, eintraege(e).IstOberstufe) Then GoTo SABKNext2
            If HatAntiPflicht(g_wuensche, g_nW, g_lehrer(i).name, eintraege(e).klasse) Then GoTo SABKNext2
            sc = BerechneScore(i, e, eintraege, nE)
            If sc > best Then best = sc: bestNm = g_lehrer(i).name
SABKNext2:
        Next i
    End If
    ' Letzter Ausweg: Anti-Pflicht ignorieren, nur KannFach pruefen
    If bestNm = "?" Then
        For i = 1 To g_nL
            If hatSperre And e = sperrIdx And g_lehrer(i).name = sperrName Then GoTo SABKNext3
            If Not KannFach(g_lehrer(i), eintraege(e).fach, eintraege(e).IstOberstufe) Then GoTo SABKNext3
            sc = BerechneScore(i, e, eintraege, nE)
            If sc > best Then best = sc: bestNm = g_lehrer(i).name
SABKNext3:
        Next i
    End If
    SA_BesterKandidat = bestNm
End Function

Function SA_ZufallsKandidat(e As Long, eintraege() As tEintrag, nE As Long, _
                             tolVal As Double, hatSperre As Boolean, _
                             sperrIdx As Long, sperrName As String, _
                             loesung() As String) As String
    ' Gibt einen zufaelligen qualifizierten Lehrer zurueck.
    ' Konsistenz: gleiche Klasse+Fach -> selber Lehrer erzwingen.
    Dim i       As Long
    Dim geeignet() As Long
    Dim nG      As Long
    Dim e2      As Long
    Dim zwangSA As String

    ' Konsistenz: Geschwister-Eintrag in loesung() bereits belegt?
    zwangSA = ""
    For e2 = 1 To nE
        If e2 <> e Then
            If LCase(eintraege(e2).klasse) = LCase(eintraege(e).klasse) Then
                If LCase(eintraege(e2).fach) = LCase(eintraege(e).fach) Then
                    If IstFixiert(eintraege(e2).lehrer) Then
                        zwangSA = eintraege(e2).lehrer: Exit For
                    End If
                    If loesung(e2) <> "" And loesung(e2) <> "?" Then
                        zwangSA = loesung(e2): Exit For
                    End If
                End If
            End If
        End If
    Next e2
    ' Konsistenz wird am Ende in SA_Solve als Nachbearbeitungsschritt behandelt

    ReDim geeignet(1 To g_nL)
    nG = 0
    For i = 1 To g_nL
        If hatSperre And e = sperrIdx And g_lehrer(i).name = sperrName Then GoTo SAZKNext
        If Not KannFach(g_lehrer(i), eintraege(e).fach, eintraege(e).IstOberstufe) Then GoTo SAZKNext
        If HatAntiPflicht(g_wuensche, g_nW, g_lehrer(i).name, eintraege(e).klasse) Then GoTo SAZKNext
        If tolVal < 9999 Then
            If g_lehrer(i).istWSt + eintraege(e).WertUV > g_lehrer(i).sollWst + tolVal Then GoTo SAZKNext
        End If
        nG = nG + 1: geeignet(nG) = i
SAZKNext:
    Next i
    If nG = 0 Then SA_ZufallsKandidat = "?": Exit Function
    SA_ZufallsKandidat = g_lehrer(geeignet(Int(Rnd * nG) + 1)).name
End Function

Function SA_GesamtScore(loesung() As String, eintraege() As tEintrag, _
                         nE As Long, tolVal As Double) As Double
    ' Berechnet Gesamtscore der aktuellen Loesung.
    ' Strafpunkte fuer unbesetzte Eintraege und Kapazitaetsueberschreitung.
    Dim e   As Long
    Dim i   As Long
    Dim sc  As Double
    sc = 0
    For e = 1 To nE
        If loesung(e) = "?" Or loesung(e) = "" Then
            sc = sc - 500   ' Strafe fuer unbesetzte Stelle
        Else
            Dim lIdx As Long: lIdx = SA_LehrerIdx(loesung(e))
            If lIdx > 0 Then sc = sc + BerechneScore(lIdx, e, eintraege, nE)
        End If
    Next e
    ' Strafe fuer Kapazitaetsueberschreitung
    If tolVal < 9999 Then
        For i = 1 To g_nL
            If g_lehrer(i).istWSt > g_lehrer(i).sollWst + tolVal Then
                sc = sc - (g_lehrer(i).istWSt - g_lehrer(i).sollWst - tolVal) * 50
            End If
        Next i
    End If
    SA_GesamtScore = sc
End Function

' ============================================================
' PARAMETER-SHEET
' ============================================================
Sub ParameterEinlesen()
    g_fgAnz = 0
    ' Liest Score-Parameter und Toleranz aus Tabelle "Parameter".
    ' Wenn Tabelle fehlt oder Wert leer: Standardwert verwenden.
    Dim ws  As Worksheet
    Dim val As Double

    ' Standardwerte
    g_scoreKL = 100
    g_scoreW3 = 80
    g_scoreW2 = 40
    g_scoreW1 = 15
    g_scoreA2 = 40
    g_scoreA1 = 15
    g_scoreKont = 8
    g_scoreFreiF = 3
    g_scoreUeberF = 10
    g_scoreUnterF = 5
    g_toleranz = 2

    Set ws = SheetByName("Parameter")
    If ws Is Nothing Then
        ' Tabelle fehlt -> erstellen und Standardwerte eintragen
        Call ParameterSheet_Erstellen
        Exit Sub
    End If

    ' Werte aus Spalte B lesen (Zeile 3 bis 12)
    If IsNumeric(ws.Cells(3, 2).Value) Then g_toleranz = CDbl(ws.Cells(3, 2).Value)
    If IsNumeric(ws.Cells(5, 2).Value) Then g_scoreKL = CDbl(ws.Cells(5, 2).Value)
    If IsNumeric(ws.Cells(6, 2).Value) Then g_scoreW3 = CDbl(ws.Cells(6, 2).Value)
    If IsNumeric(ws.Cells(7, 2).Value) Then g_scoreW2 = CDbl(ws.Cells(7, 2).Value)
    If IsNumeric(ws.Cells(8, 2).Value) Then g_scoreW1 = CDbl(ws.Cells(8, 2).Value)
    If IsNumeric(ws.Cells(9, 2).Value) Then g_scoreA2 = CDbl(ws.Cells(9, 2).Value)
    If IsNumeric(ws.Cells(10, 2).Value) Then g_scoreA1 = CDbl(ws.Cells(10, 2).Value)
    If IsNumeric(ws.Cells(11, 2).Value) Then g_scoreKont = CDbl(ws.Cells(11, 2).Value)
    If IsNumeric(ws.Cells(12, 2).Value) Then g_scoreFreiF = CDbl(ws.Cells(12, 2).Value)
    If IsNumeric(ws.Cells(13, 2).Value) Then g_scoreUeberF = CDbl(ws.Cells(13, 2).Value)
    If IsNumeric(ws.Cells(14, 2).Value) Then g_scoreUnterF = CDbl(ws.Cells(14, 2).Value)
    ' Pruefe ob Sheet aktuelles Format hat (Zeile 15 = Zeitlimit)
    ' Altes Format: Zeile 15 = Schutzfachgruppen-Trenner -> Sheet neu erstellen
    Dim zelle15 As String: zelle15 = Trim(CStr(ws.Cells(15, 1).Value))
    If InStr(LCase(zelle15), "schutz") > 0 Or InStr(LCase(zelle15), "---") > 0 Then
        ' Altes Format erkannt -> Sheet aktualisieren und neu einlesen
        Call ParameterSheet_Erstellen
        Exit Sub
    End If
    ' Zeitlimit (Zeile 15), Schutzfachgruppen (Zeilen 17-19)
    ' Immer zuruecksetzen - nie alte Werte aus vorherigem Lauf behalten
    g_zeitlimit = 60
    g_schutzFG1 = ""
    g_schutzFG2 = ""
    g_schutzMalus = 20
    If IsNumeric(ws.Cells(15, 2).Value) Then g_zeitlimit = CDbl(ws.Cells(15, 2).Value)
    Dim sfg1 As String: sfg1 = Trim(CStr(ws.Cells(17, 2).Value))
    Dim sfg2 As String: sfg2 = Trim(CStr(ws.Cells(18, 2).Value))
    ' Nur setzen wenn nicht leer und nicht "0" (leere Zelle liefert manchmal "0")
    If sfg1 <> "" And sfg1 <> "0" And sfg1 <> "False" Then g_schutzFG1 = sfg1
    If sfg2 <> "" And sfg2 <> "0" And sfg2 <> "False" Then g_schutzFG2 = sfg2
    If IsNumeric(ws.Cells(19, 2).Value) Then g_schutzMalus = CDbl(ws.Cells(19, 2).Value)
    ' Debug-Ausgabe in Statusleiste
    If g_schutzFG1 <> "" Or g_schutzFG2 <> "" Then
        Application.StatusBar = "Schutzfachgruppen aktiv: " & g_schutzFG1 & IIf(g_schutzFG2 <> "", " / " & g_schutzFG2, "")
    Else
        Application.StatusBar = "Schutzfachgruppen: inaktiv"
    End If
End Sub

Sub ParameterSheet_Erstellen()
    ' Erstellt oder aktualisiert die Tabelle "Parameter" mit Standardwerten.
    Dim ws  As Worksheet
    Dim cH  As Long
    Dim cSH As Long

    Set ws = SheetByName("Parameter")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(1))
        ws.name = "Parameter"
    End If

    ws.Cells.Clear
    cH = RGB(31, 73, 125)
    cSH = RGB(68, 114, 196)

    ' Titel
    ws.Cells(1, 1).Value = "Optimierungsparameter"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 13
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 3)).Merge
    ws.Cells(1, 1).Interior.Color = cH
    ws.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    ws.rows(1).RowHeight = 22

    ' Spaltenk" & Chr(246) & "pfe
    ws.Cells(2, 1).Value = "Parameter"
    ws.Cells(2, 2).Value = "Wert"
    ws.Cells(2, 3).Value = "Erkl" & Chr(228) & "rung"
    With ws.Range(ws.Cells(2, 1), ws.Cells(2, 3))
        .Font.Bold = True
        .Interior.Color = cSH
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Daten: Param-Name | Wert | Erkl" & Chr(228) & "rung
    Dim rows(11) As String
    Dim vals(11) As Double
    Dim expl(11) As String

    rows(0) = "Toleranz (Soll-" & Chr(220) & "berschreitung)": vals(0) = 2:     expl(0) = "Max. erlaubte Abweichung vom Soll-Wert (z.B. 2 = 2 Std. Puffer)"
    rows(1) = "---":                                             vals(1) = 0:    expl(1) = "--- Scores ---"
    rows(2) = "Klassenleitung":                                  vals(2) = 100: expl(2) = "Bonus wenn Lehrer KlassenleiterIn dieser Klasse ist"
    rows(3) = "Wunsch Prio 3":                                   vals(3) = 80:   expl(3) = "Lehrer hat dringenden Wunsch f" & Chr(252) & "r diese Klasse"
    rows(4) = "Wunsch Prio 2":                                   vals(4) = 40:   expl(4) = "Lehrer hat mittleren Wunsch f" & Chr(252) & "r diese Klasse"
    rows(5) = "Wunsch Prio 1":                                   vals(5) = 15:   expl(5) = "Lehrer hat schwachen Wunsch f" & Chr(252) & "r diese Klasse"
    rows(6) = "Anti-Wunsch Prio 2":                              vals(6) = 40:   expl(6) = "Abzug: Lehrer m" & Chr(246) & "chte diese Klasse nicht (Prio 2)"
    rows(7) = "Anti-Wunsch Prio 1":                              vals(7) = 15:   expl(7) = "Abzug: Lehrer m" & Chr(246) & "chte diese Klasse vermeiden (Prio 1)"
    rows(8) = "Kontinuit" & Chr(228) & "t":                     vals(8) = 8:    expl(8) = "Bonus wenn Lehrer schon anderes Fach in dieser Klasse hat"
    rows(9) = "Faktor freie Stunden":                            vals(9) = 3:    expl(9) = "Freie Stunden x Faktor = Bonus (h" & Chr(246) & "her = gleichm" & Chr(228) & "ssigere Verteilung)"
    rows(10) = "Faktor " & Chr(220) & "berlastung":             vals(10) = 10:  expl(10) = Chr(220) & "berlastung x Faktor = Malus (h" & Chr(246) & "her = " & Chr(220) & "berlastung st" & Chr(228) & "rker vermeiden)"

    rows(11) = "Faktor Unterbesetzung":                          vals(11) = 5:   expl(11) = "Unterbesetzung x Faktor = Malus (h" & Chr(246) & "her = Unterbesetzung st" & Chr(228) & "rker vermeiden)"

    Dim r As Long
    Dim dataRow As Long
    dataRow = 3
    For r = 0 To 11
        If rows(r) = "---" Then
            ' Trennzeile
            ws.Cells(dataRow, 1).Value = ""
            ws.Range(ws.Cells(dataRow, 1), ws.Cells(dataRow, 3)).Interior.Color = RGB(220, 230, 242)
            ws.Cells(dataRow, 3).Value = expl(r)
            ws.Cells(dataRow, 3).Font.Bold = True
            ws.Cells(dataRow, 3).Font.Color = cH
        Else
            ws.Cells(dataRow, 1).Value = rows(r)
            ws.Cells(dataRow, 2).Value = vals(r)
            ws.Cells(dataRow, 2).NumberFormat = "0.##"
            ws.Cells(dataRow, 3).Value = expl(r)
            If dataRow Mod 2 = 0 Then
                ws.Range(ws.Cells(dataRow, 1), ws.Cells(dataRow, 3)).Interior.Color = RGB(242, 242, 242)
            End If
        End If
        dataRow = dataRow + 1
    Next r

    ' Zeitlimit (Zeile 15)
    ws.Cells(dataRow, 1).Value = "Zeitlimit BT-Phase (Sek)"
    ws.Cells(dataRow, 2).Value = 60
    ws.Cells(dataRow, 2).NumberFormat = "0"
    ws.Cells(dataRow, 3).Value = "Max. Sekunden fuer Backtracking-Phase (z.B. 30, 60, 120)"
    If dataRow Mod 2 = 0 Then
        ws.Range(ws.Cells(dataRow, 1), ws.Cells(dataRow, 3)).Interior.Color = RGB(242, 242, 242)
    End If
    dataRow = dataRow + 1

    ' Schutzfachgruppen-Trenner
    ws.Cells(dataRow, 1).Value = "--- Schutzfachgruppen ---"
    ws.Range(ws.Cells(dataRow, 1), ws.Cells(dataRow, 3)).Interior.Color = RGB(220, 230, 242)
    ws.Cells(dataRow, 3).Value = "--- optional: Fachgruppen vor Ueberlast schuetzen ---"
    ws.Cells(dataRow, 3).Font.Bold = True
    ws.Cells(dataRow, 3).Font.Color = cH
    dataRow = dataRow + 1

    ws.Cells(dataRow, 1).Value = "Schutzfachgruppe 1"
    ws.Cells(dataRow, 2).Value = ""
    ws.Cells(dataRow, 3).Value = "Fachgruppen-Kuerzel (z.B. GE, D, SP) oder leer lassen"
    dataRow = dataRow + 1

    ws.Cells(dataRow, 1).Value = "Schutzfachgruppe 2"
    ws.Cells(dataRow, 2).Value = ""
    ws.Cells(dataRow, 3).Value = "Zweite Schutzfachgruppe oder leer lassen"
    dataRow = dataRow + 1

    ws.Cells(dataRow, 1).Value = "Schutz-Malus-Faktor"
    ws.Cells(dataRow, 2).Value = 20
    ws.Cells(dataRow, 2).NumberFormat = "0.##"
    ws.Cells(dataRow, 3).Value = "Malus pro WSt Ueberschreitung fuer Schutzlehrer (hoeher = staerkerer Schutz)"
    dataRow = dataRow + 1

    ' Spalten B einfaerben als editierbar (Zeilen 3-14: Score-Parameter gelb)
    With ws.Range(ws.Cells(3, 2), ws.Cells(15, 2))
        .Interior.Color = RGB(255, 255, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    ' Schutzfachgruppen-Felder orange (dataRow zeigt jetzt auf Zeile nach Malus)
    ' Schutz-Zeilen: Trenner=dataRow-4, FG1=dataRow-3, FG2=dataRow-2, Malus=dataRow-1
    Dim schutzStart As Long: schutzStart = dataRow - 3  ' FG1
    Dim schutzEnd   As Long: schutzEnd = dataRow - 1    ' Malus
    With ws.Range(ws.Cells(schutzStart, 2), ws.Cells(schutzEnd, 2))
        .Interior.Color = RGB(255, 230, 180)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ' Spaltenbreiten
    ws.Columns(1).ColumnWidth = 28
    ws.Columns(2).ColumnWidth = 10
    ws.Columns(3).ColumnWidth = 60

    ' Hinweis
    ws.Cells(dataRow, 1).Value = "Hinweis: Gelbe Felder editieren. Schutzfachgruppe leer lassen = inaktiv."
    ws.Cells(dataRow, 1).Font.Italic = True
    ws.Cells(dataRow, 1).Font.Color = RGB(128, 128, 128)
    ws.Range(ws.Cells(dataRow, 1), ws.Cells(dataRow, 3)).Merge

    ws.Activate
    MsgBox "Parameter-Tabelle erstellt." & vbCrLf & _
           "Passen Sie die gelben Felder an und starten Sie dann die Optimierung.", _
           vbInformation, "Parameter"
End Sub

Sub Debug_KonsistenzPruefen()
    'Prueft ob gleiche Klasse+Fach immer denselben Lehrer hat.
    'Zeigt alle Verletzungen an.
    Dim wsK As Worksheet
    Dim eintraege() As tEintrag
    Dim nE As Long
    Dim e As Long
    Dim e2 As Long
    Dim msg As String
    Dim cnt As Long

    Set wsK = SheetByName("Klassen")
    If wsK Is Nothing Then MsgBox "Kein Klassen-Sheet!": Exit Sub

    nE = EintraegeEinlesen(wsK, eintraege)
    msg = ""
    cnt = 0

    For e = 1 To nE
        If eintraege(e).lehrer = "" Or eintraege(e).lehrer = "?" Then GoTo NaechsterE
        For e2 = e + 1 To nE
            If eintraege(e2).lehrer = "" Or eintraege(e2).lehrer = "?" Then GoTo NaechsterE2
            If LCase(eintraege(e).klasse) <> LCase(eintraege(e2).klasse) Then GoTo NaechsterE2
            If LCase(eintraege(e).fach) <> LCase(eintraege(e2).fach) Then GoTo NaechsterE2
            If eintraege(e).lehrer <> eintraege(e2).lehrer Then
                cnt = cnt + 1
                If cnt <= 10 Then
                    msg = msg & "Zeile " & eintraege(e).zeile & " vs " & eintraege(e2).zeile & ": "
                    msg = msg & eintraege(e).fach & " in " & eintraege(e).klasse & " -> "
                    msg = msg & eintraege(e).lehrer & " vs " & eintraege(e2).lehrer & vbCrLf
                End If
            End If
NaechsterE2:
        Next e2
NaechsterE:
    Next e

    If cnt = 0 Then
        MsgBox "Keine Konsistenz-Verletzungen gefunden!", vbInformation
    Else
        MsgBox cnt & " Verletzung(en):" & vbCrLf & vbCrLf & msg, vbExclamation, "Konsistenz-Fehler"
    End If
End Sub

' ============================================================
' STAMMDATEN-IMPORT: Klassen -> Klassenleitungen in Lehrerliste
' ============================================================
Sub StammdatenKlassen_Import()
    Dim wsSrc    As Worksheet
    Dim wsDest   As Worksheet
    Dim hRow     As Long
    Dim maxR     As Long
    Dim r        As Long
    Dim rL       As Long
    Dim c        As Long
    Dim lastRowL As Long
    Dim cName    As Long
    Dim cKL      As Long
    Dim klassenName  As String
    Dim klLehrer     As String
    Dim ersterL      As String
    Dim lehrerName   As String
    Dim updated      As Long
    Dim notFound     As Long
    Dim notFoundMsg  As String
    Dim gefunden     As Boolean
    Dim hVal         As String
    Dim msg          As String

    Set wsSrc = SheetByName("Stammdaten Klassen")
    Set wsDest = SheetByName("Lehrerliste")
    If wsSrc Is Nothing Then
        MsgBox "Tabelle 'Stammdaten Klassen' nicht gefunden!", vbCritical: Exit Sub
    End If
    If wsDest Is Nothing Then
        MsgBox "Tabelle 'Lehrerliste' nicht gefunden!", vbCritical: Exit Sub
    End If

    ' Kopfzeile in Stammdaten suchen (erste 5 Zeilen)
    hRow = 0: cName = 0: cKL = 0
    For r = 1 To 5
        For c = 1 To 15
            hVal = Trim(CStr(wsSrc.Cells(r, c).Value))
            If LCase(hVal) = "name" And cName = 0 Then
                cName = c: hRow = r
            End If
            If (LCase(hVal) = "klassenlehrer" Or LCase(hVal) = "kl") And cKL = 0 Then
                cKL = c
            End If
        Next c
        If cName > 0 And cKL > 0 Then Exit For
    Next r

    If cName = 0 Then
        MsgBox "Spalte 'Name' in 'Stammdaten Klassen' nicht gefunden!" & vbCrLf & _
               "Erwartet in den ersten 5 Zeilen.", vbCritical
        Exit Sub
    End If
    If cKL = 0 Then
        MsgBox "Spalte 'Klassenlehrer' nicht gefunden!", vbCritical
        Exit Sub
    End If

    maxR = wsSrc.Cells(wsSrc.rows.Count, cName).End(xlUp).row
    lastRowL = wsDest.Cells(wsDest.rows.Count, 1).End(xlUp).row
    updated = 0: notFound = 0: notFoundMsg = ""

    For r = hRow + 1 To maxR
        klassenName = Trim(CStr(wsSrc.Cells(r, cName).Value))
        klLehrer = Trim(CStr(wsSrc.Cells(r, cKL).Value))
        If klassenName = "" Or klLehrer = "" Then GoTo NaechsteZeile

        ' Alle Klassenlehrer verarbeiten (kommasepariert)
        Dim klArr() As String
        Dim kli     As Long
        klArr = Split(klLehrer, ",")
        For kli = 0 To UBound(klArr)
            Dim einL As String
            einL = Trim(klArr(kli))
            If einL = "" Then GoTo NaechsterKL

            ' In Lehrerliste suchen und Spalte 18 (KlassenleiterIn) setzen
            gefunden = False
            For rL = 2 To lastRowL
                lehrerName = Trim(CStr(wsDest.Cells(rL, 1).Value))
                If LCase(lehrerName) = LCase(einL) Then
                    wsDest.Cells(rL, 18).Value = klassenName
                    updated = updated + 1
                    gefunden = True
                    Exit For
                End If
            Next rL

            If Not gefunden Then
                notFound = notFound + 1
                If notFound <= 10 Then
                    notFoundMsg = notFoundMsg & "  - " & einL & _
                                  " (KL von " & klassenName & ")" & vbCrLf
                End If
            End If
NaechsterKL:
        Next kli
NaechsteZeile:
    Next r

    msg = updated & " Klassenleitung(en) in Lehrerliste eingetragen."
    If notFound > 0 Then
        msg = msg & vbCrLf & vbCrLf & notFound & " Lehrer nicht gefunden:" & vbCrLf & notFoundMsg
        If notFound > 10 Then
            msg = msg & "  ... und " & (notFound - 10) & " weitere."
        End If
        MsgBox msg, vbExclamation, "Stammdaten-Import"
    Else
        MsgBox msg, vbInformation, "Stammdaten-Import"
    End If
End Sub

' ============================================================
' GPU002-EXPORT: Lehrer aus Tabelle Klassen in GPU002.TXT eintragen
' ============================================================
Sub GPU002_LehrerExport()
    ' Liest GPU002.TXT, vergleicht UNr+Klasse+Fach mit Tabelle Klassen,
    ' und ersetzt den Lehrer (auch "?") durch den aus Tabelle Klassen.
    Dim wsK      As Worksheet
    Dim filePath As String
    Dim outPath  As String
    Dim fileNr   As Integer
    Dim outNr    As Integer
    Dim zeile    As String
    Dim parts()  As String
    Dim klasse   As String
    Dim fach     As String
    Dim lehrer   As String
    Dim nE       As Long
    Dim eintraege() As tEintrag
    Dim e        As Long
    Dim updated  As Long
    Dim total    As Long
    Dim lines()  As String
    Dim nLines   As Long
    Dim li       As Long

    Set wsK = SheetByName("Klassen")
    If wsK Is Nothing Then
        MsgBox "Tabelle 'Klassen' nicht gefunden!", vbCritical: Exit Sub
    End If

    ' Datei auswaehlen
    filePath = Application.GetOpenFilename( _
        "GPU002-Datei (*.TXT;*.txt),*.TXT;*.txt", , _
        "GPU002.TXT auswaehlen")
    If filePath = "False" Or filePath = "" Then Exit Sub

    ' Speicherort fuer Ausgabedatei
    outPath = left(filePath, Len(filePath) - 4) & "_neu.TXT"
    outPath = Application.GetSaveAsFilename( _
        outPath, "GPU002-Datei (*.TXT;*.txt),*.TXT;*.txt", , _
        "Ausgabedatei waehlen")
    If outPath = "False" Or outPath = "" Then Exit Sub

    ' Eintraege aus Klassen-Tabelle lesen
    nE = EintraegeEinlesen(wsK, eintraege)
    If nE = 0 Then
        MsgBox "Keine Eintraege in Tabelle Klassen!", vbCritical: Exit Sub
    End If

    ' GPU002 komplett einlesen
    nLines = 0
    fileNr = FreeFile
    Open filePath For Input As #fileNr
    Do While Not EOF(fileNr)
        Line Input #fileNr, zeile
        nLines = nLines + 1
        ReDim Preserve lines(1 To nLines)
        lines(nLines) = zeile
    Loop
    Close #fileNr

    ' Zeilenweise verarbeiten und Lehrer ersetzen
    updated = 0: total = nLines
    For li = 1 To nLines
        zeile = lines(li)
        If Trim(zeile) = "" Then GoTo NaechsteZeile

        parts = Split(zeile, ";")
        If UBound(parts) < 6 Then GoTo NaechsteZeile

        ' Felder extrahieren (Anführungszeichen entfernen)
        klasse = Replace(parts(4), Chr(34), "")
        fach = Replace(parts(6), Chr(34), "")

        If klasse = "" Or fach = "" Then GoTo NaechsteZeile

        ' In Klassen-Tabelle suchen: gleiche Klasse + Fach
        ' (UNr existiert in Klassen-Tabelle nicht, daher nur Klasse+Fach)
        For e = 1 To nE
            If LCase(eintraege(e).klasse) <> LCase(klasse) Then GoTo NaechsterEintrag
            If LCase(eintraege(e).fach) <> LCase(fach) Then GoTo NaechsterEintrag

            ' Lehrer aus Klassen-Tabelle nehmen (L1, Spalte 4)
            Dim neuerLehrer As String
            neuerLehrer = Trim(CStr(wsK.Cells(eintraege(e).zeile, 4).Value))
            If neuerLehrer = "" Or neuerLehrer = "?" Then GoTo NaechsterEintrag

            ' Feld 5 (Lehrer) in der Zeile ersetzen
            parts(5) = Chr(34) & neuerLehrer & Chr(34)
            lines(li) = Join(parts, ";")
            updated = updated + 1
            Exit For
NaechsterEintrag:
        Next e
NaechsteZeile:
    Next li

    ' Ausgabedatei schreiben
    outNr = FreeFile
    Open outPath For Output As #outNr
    For li = 1 To nLines
        Print #outNr, lines(li)
    Next li
    Close #outNr

    MsgBox updated & " von " & total & " Zeilen aktualisiert." & vbCrLf & _
           "Ausgabe: " & outPath, vbInformation, "GPU002-Export"
End Sub

' ============================================================
' DEBUG: Warum bleibt ein Eintrag bei SA unbesetzt?
' ============================================================
Sub Debug_SA_LeereBesetzung()
    Dim wsL As Worksheet
    Dim wsK As Worksheet
    Dim wsW As Worksheet
    Dim lehrer()   As tLehrer
    Dim nL         As Long
    Dim wuensche() As tWunsch
    Dim nW         As Long
    Dim eintraege() As tEintrag
    Dim nE         As Long
    Dim e          As Long
    Dim i          As Long
    Dim msg        As String
    Dim fachMsg    As String
    Dim cnt        As Long

    Set wsL = SheetByName("Lehrerliste")
    Set wsK = SheetByName("Klassen")
    Set wsW = SheetByName("Lehrerw" & Chr(252) & "nsche")
    If wsL Is Nothing Or wsK Is Nothing Or wsW Is Nothing Then
        MsgBox "Benoetigt: Lehrerliste, Klassen, Lehrerw" & Chr(252) & "nsche", vbCritical
        Exit Sub
    End If

    ReDim g_lehrer(1 To 1)
    g_nL = LehrerEinlesen(wsL, g_lehrer)
    Call SchutzFlagsSetzen
    ReDim g_wuensche(1 To 1)
    g_nW = WuenscheEinlesen(wsW, g_wuensche)
    nE = EintraegeEinlesen(wsK, eintraege)

    ' Parameter laden
    Call ParameterEinlesen

    msg = "Eintraege ohne qualifizierten Lehrer (offen):" & vbCrLf & vbCrLf
    cnt = 0

    For e = 1 To nE
        If IstFixiert(eintraege(e).UrsprungsLehrer) Then GoTo WeiterD

        ' Kandidaten zaehlen
        Dim nKand As Long: nKand = 0
        Dim nKandOhne As Long: nKandOhne = 0
        Dim kannListe As String: kannListe = ""
        Dim vollListe As String: vollListe = ""

        For i = 1 To g_nL
            If KannFach(g_lehrer(i), eintraege(e).fach, eintraege(e).IstOberstufe) Then
                nKandOhne = nKandOhne + 1
                If g_lehrer(i).sollWst > 0 Then
                    kannListe = kannListe & g_lehrer(i).name & _
                        "(Soll=" & Format(g_lehrer(i).sollWst, "0.0") & ") "
                End If
                If g_lehrer(i).istWSt + eintraege(e).WertUV <= g_lehrer(i).sollWst + g_toleranz Then
                    nKand = nKand + 1
                End If
            End If
        Next i

        If nKandOhne = 0 Then
            cnt = cnt + 1
            If cnt <= 8 Then
                msg = msg & "- " & eintraege(e).fach & " in " & eintraege(e).klasse & _
                      " (Zeile " & eintraege(e).zeile & "): KEIN Lehrer mit KannFach!" & vbCrLf
            End If
        ElseIf nKand = 0 Then
            cnt = cnt + 1
            If cnt <= 8 Then
                msg = msg & "- " & eintraege(e).fach & " in " & eintraege(e).klasse & _
                      " (Zeile " & eintraege(e).zeile & "): " & nKandOhne & " Lehrer qualif. " & _
                      "aber ALLE voll (Toleranz=" & Format(g_toleranz, "0.0") & ")" & vbCrLf
                msg = msg & "  Qualif.: " & kannListe & vbCrLf
            End If
        End If
WeiterD:
    Next e

    If cnt = 0 Then
        MsgBox "Alle offenen Eintraege haben mindestens einen qualifizierten Lehrer mit Kapazitaet.", _
               vbInformation, "Debug SA"
    Else
        If cnt > 8 Then msg = msg & "... und " & (cnt - 8) & " weitere."
        msg = msg & vbCrLf & "Hinweis: IstWSt-Werte sind aus der Klassen-Tabelle (Spalte L1)," & _
              vbCrLf & "nicht aus einem SA-Lauf."
        MsgBox msg, vbExclamation, "Debug SA: " & cnt & " Problem(e)"
    End If
End Sub

' ============================================================
' STUNDENTAFEL-VERGLEICH
' Vergleicht Soll-Stunden (Sheet "Stundentafel") mit
' tatsaechlichen Eintraegen (Sheet "Klassen").
' Schreibt Abweichungen in Sheet "Stundentafel_Pruefung".
' ============================================================
Sub StundentafelVergleich()
    Dim wsST   As Worksheet   ' Stundentafel
    Dim wsK    As Worksheet   ' Klassen
    Dim wsOut  As Worksheet   ' Ausgabe

    Set wsST = SheetByName("Stundentafel")
    Set wsK = SheetByName("Klassen")
    If wsST Is Nothing Then
        MsgBox "Sheet 'Stundentafel' nicht gefunden!" & vbCrLf & _
               "Bitte eine Tabelle mit Spalten: Fach | 5 | 6 | 7 | 8 | 9 | 10 erstellen.", _
               vbCritical: Exit Sub
    End If
    If wsK Is Nothing Then
        MsgBox "Sheet 'Klassen' nicht gefunden!", vbCritical: Exit Sub
    End If

    ' Ausgabe-Sheet erstellen/leeren
    Set wsOut = SheetByName("Stundentafel_Pruefung")
    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsOut.name = "Stundentafel_Pruefung"
    Else
        wsOut.Cells.Clear
        ' Alle Shapes loeschen und Button danach neu erstellen
        Dim shpDel As Shape
        For Each shpDel In wsOut.Shapes
            shpDel.Delete
        Next shpDel
    End If

    ' Button in Spalte F, Zeile 1 neu erstellen
    Dim btnOut As Shape
    Set btnOut = wsOut.Shapes.AddFormControl(xlButtonControl, wsOut.Columns("F").left, wsOut.rows(1).top + 2, 170, 22)
    btnOut.TextFrame.Characters.Text = "> Stundentafel pr" & Chr(252) & "fen"
    btnOut.TextFrame.Characters.Font.Bold = True
    btnOut.TextFrame.Characters.Font.Size = 10
    btnOut.OnAction = "StundentafelVergleich"

    ' ---- Stundentafel einlesen ----
    ' Kopfzeile suchen: erste Zeile mit Jahrgangsstufen-Angaben
    ' Erkennt sowohl "5" als auch "Jhg 5", "Jhg. 5" etc.
    Dim stMaxR As Long: stMaxR = wsST.Cells(wsST.rows.Count, 1).End(xlUp).row
    Dim stHRow As Long: stHRow = 1
    Dim stMaxC As Long: stMaxC = 0
    Dim rr As Long, cc As Long

    For rr = 1 To 10
        For cc = 1 To 20
            Dim hv As String: hv = Trim(CStr(wsST.Cells(rr, cc).Value))
            ' Vorgabe- und Summenzellen ueberspringen
            If InStr(LCase(hv), "vorg") > 0 Then GoTo NaechsteHdrZelle
            If InStr(LCase(hv), "summe") > 0 Then GoTo NaechsteHdrZelle
            If InStr(LCase(hv), "si") > 0 And Len(hv) <= 5 Then GoTo NaechsteHdrZelle
            If InStr(hv, "+") > 0 Then GoTo NaechsteHdrZelle
            ' Zahl aus Zelle extrahieren (auch aus "Jhg 5", "Jhg. 10")
            Dim hvNum As String: hvNum = ""
            Dim hvi As Long: hvi = 1
            Do While hvi <= Len(hv)
                If Mid(hv, hvi, 1) >= "0" And Mid(hv, hvi, 1) <= "9" Then
                    hvNum = hvNum & Mid(hv, hvi, 1)
                ElseIf hvNum <> "" Then
                    hvi = Len(hv) + 1  ' Abbruch
                End If
                hvi = hvi + 1
            Loop
            If IsNumeric(hvNum) Then
                Dim hvJhg As Long: hvJhg = CLng(hvNum)
                If hvJhg >= 5 And hvJhg <= 10 Then
                    stHRow = rr
                    stMaxC = wsST.Cells(rr, wsST.Columns.Count).End(xlToLeft).Column
                    Exit For
                End If
            End If
NaechsteHdrZelle:
        Next cc
        If stMaxC > 0 Then Exit For
    Next rr
    If stMaxC = 0 Then stMaxC = wsST.Cells(stHRow, wsST.Columns.Count).End(xlToLeft).Column

    ' Jahrgangsstufen aus Kopfzeile lesen
    ' Extrahiert Zahl aus "5", "Jhg 5", "Jhg. 5" etc.; Vorgabe/SI-Spalten -> 0
    Dim jhgAnz As Long: jhgAnz = stMaxC - 1
    If jhgAnz <= 0 Then MsgBox "Stundentafel: keine Spalten gefunden!" & vbCrLf & _
        "Bitte sicherstellen dass die Kopfzeile die Jahrgangsstufen (5-10) enth" & Chr(228) & "lt.", vbCritical: Exit Sub

    Dim jhgStufen() As Long
    ReDim jhgStufen(1 To jhgAnz)
    Dim c As Long
    For c = 2 To stMaxC
        Dim jhgVal As String: jhgVal = Trim(CStr(wsST.Cells(stHRow, c).Value))
        ' Vorgabe- und Summenspalten -> 0
        If InStr(LCase(jhgVal), "vorg") > 0 Then
            jhgStufen(c - 1) = 0
        ElseIf InStr(LCase(jhgVal), "summe") > 0 Then
            jhgStufen(c - 1) = 0
        ElseIf InStr(jhgVal, "+") > 0 Then
            jhgStufen(c - 1) = 0
        ElseIf InStr(LCase(jhgVal), "si") > 0 And Len(jhgVal) <= 5 Then
            jhgStufen(c - 1) = 0
        Else
            ' Zahl extrahieren (aus "5", "Jhg 5", "Jhg. 10" etc.)
            Dim jNum As String: jNum = ""
            Dim ji As Long: ji = 1
            Do While ji <= Len(jhgVal)
                If Mid(jhgVal, ji, 1) >= "0" And Mid(jhgVal, ji, 1) <= "9" Then
                    jNum = jNum & Mid(jhgVal, ji, 1)
                ElseIf jNum <> "" Then
                    ji = Len(jhgVal) + 1  ' Abbruch
                End If
                ji = ji + 1
            Loop
            If IsNumeric(jNum) Then
                Dim jn As Long: jn = CLng(jNum)
                If jn >= 5 And jn <= 10 Then jhgStufen(c - 1) = jn
            End If
        End If
    Next c

    ' Faecher und Soll-Stunden einlesen
    Dim stFach() As String
    Dim stWSt()  As Double
    Dim stAnz    As Long: stAnz = 0
    ReDim stFach(1 To (stMaxR - 1) * jhgAnz)
    ReDim stWSt(1 To (stMaxR - 1) * jhgAnz)
    Dim stJhg()  As Long
    ReDim stJhg(1 To (stMaxR - 1) * jhgAnz)

    Dim r As Long
    For r = stHRow + 1 To stMaxR
        Dim fach As String: fach = Trim(CStr(wsST.Cells(r, 1).Value))
        If fach = "" Then GoTo NaechsteZeile
        ' Summen-, Hinweis- und Ergaenzungsstunden-Zeilen ueberspringen
        Dim fachU As String: fachU = UCase(fach)
        If left(fachU, 5) = "SUMME" Then GoTo NaechsteZeile
        If left(fachU, 4) = "KERN" Then GoTo NaechsteZeile
        If left(fachU, 4) = "VORG" Then GoTo NaechsteZeile
        If left(fachU, 11) = "WOCHENSTUND" Then GoTo NaechsteZeile
        If left(fachU, 3) = "ERG" Then GoTo NaechsteZeile
        If left(fachU, 6) = "METHOD" Then GoTo NaechsteZeile
        If left(fachU, 5) = "LIONS" Then GoTo NaechsteZeile
        If left(fachU, 6) = "KLASSE" Then GoTo NaechsteZeile
        If left(fachU, 5) = "SPORT" And InStr(fach, ":") > 0 Then GoTo NaechsteZeile
        If left(fachU, 1) = "0" Or left(fachU, 1) = "1" Then GoTo NaechsteZeile  ' Hinweiszeilen die mit Zahlen beginnen
        If left(fachU, 9) = "AUSNAHME:" Then GoTo NaechsteZeile
        For c = 2 To stMaxC
            Dim wstVal As String: wstVal = Trim(CStr(wsST.Cells(r, c).Value))
            ' Leerzeichen, "-", "0" = kein Unterricht
            If wstVal = "" Or wstVal = "-" Or wstVal = "0" Then GoTo NaechsteSpalte
            If Not IsNumeric(wstVal) Then GoTo NaechsteSpalte
            Dim sollWst As Double: sollWst = CDbl(wstVal)
            If sollWst <= 0 Then GoTo NaechsteSpalte
            stAnz = stAnz + 1
            stFach(stAnz) = fach
            stJhg(stAnz) = jhgStufen(c - 1)
            stWSt(stAnz) = sollWst
NaechsteSpalte:
        Next c
NaechsteZeile:
    Next r

    ' ---- Klassen-Tabelle einlesen: IST-Stunden pro Klasse+Fach ----
    Dim klMaxR As Long: klMaxR = wsK.Cells(wsK.rows.Count, 2).End(xlUp).row
    Dim aktK   As String: aktK = ""
    ' Dictionary: "Klasse|Fach" -> IstWSt
    Dim istKeys()  As String
    Dim istWStArr() As Double
    Dim istAnz     As Long: istAnz = 0
    ReDim istKeys(1 To klMaxR * 2)
    ReDim istWStArr(1 To klMaxR * 2)

    For r = 2 To klMaxR
        Dim zK As String: zK = Trim(CStr(wsK.Cells(r, 1).Value))
        Dim zF As String: zF = Trim(CStr(wsK.Cells(r, 2).Value))
        If zK <> "" Then aktK = zK
        If zF = "" Then GoTo NaechsteKlasse
        ' Foerderkurse und Ergaenzungsstunden ueberspringen (gleiche Liste wie KlassenUV)
        Dim zFU As String: zFU = UCase(zF)
        If InStr(zFU, " LRS") > 0 Then GoTo NaechsteKlasse
        If InStr(zFU, " ERG") > 0 Then GoTo NaechsteKlasse
        If InStr(zFU, "_ERG") > 0 Then GoTo NaechsteKlasse
        If InStr(zFU, "DAZ") > 0 Then GoTo NaechsteKlasse
        If InStr(zFU, " CC") > 0 Then GoTo NaechsteKlasse
        If InStr(zFU, "DELF") > 0 Then GoTo NaechsteKlasse
        If InStr(zFU, " F" & Chr(214)) > 0 Then GoTo NaechsteKlasse
        If zFU = "AG" Or left(zFU, 3) = "AG " Then GoTo NaechsteKlasse
        If InStr(zFU, "NIX") > 0 Then GoTo NaechsteKlasse
        If InStr(zFU, "STRPR") > 0 Then GoTo NaechsteKlasse
        If InStr(zFU, "LOT") > 0 Then GoTo NaechsteKlasse
        If InStr(zFU, "VBER") > 0 Then GoTo NaechsteKlasse
        If InStr(zFU, "VPL") > 0 And Len(zFU) <= 5 Then GoTo NaechsteKlasse
        Dim zW As Double
        If IsNumeric(wsK.Cells(r, 3).Value) Then zW = CDbl(wsK.Cells(r, 3).Value)
        ' Gruppennamen aufloesen: "05A,05B,05C" -> einzelne Klassen
        Dim zKArr() As String: zKArr = Split(aktK, ",")
        Dim zKI As Long
        For zKI = 0 To UBound(zKArr)
            Dim zKOne As String: zKOne = Trim(zKArr(zKI))
            If zKOne = "" Then GoTo NaechsteKlInGrp2
            Dim key As String: key = LCase(zKOne) & "|" & LCase(zF)
            Dim found As Boolean: found = False
            Dim ki As Long
            For ki = 1 To istAnz
                If istKeys(ki) = key Then
                    istWStArr(ki) = istWStArr(ki) + zW
                    found = True: Exit For
                End If
            Next ki
            If Not found Then
                istAnz = istAnz + 1
                istKeys(istAnz) = key
                istWStArr(istAnz) = zW
            End If
NaechsteKlInGrp2:
        Next zKI
NaechsteKlasse:
    Next r

    ' ---- Klassen nach Jahrgangsstufe gruppieren ----
    ' Alle eindeutigen Klassen sammeln
    Dim allKlassen() As String
    Dim allKlAnz     As Long: allKlAnz = 0
    ReDim allKlassen(1 To istAnz)
    aktK = ""
    For r = 2 To klMaxR
        Dim kl As String: kl = Trim(CStr(wsK.Cells(r, 1).Value))
        If kl = "" Then GoTo NaechsteKl2
        ' Gruppennamen aufloesen: "05A,05B,05C" -> einzelne Klassen
        Dim klGrp() As String: klGrp = Split(kl, ",")
        Dim klGrpI As Long
        For klGrpI = 0 To UBound(klGrp)
            Dim klSingle As String: klSingle = Trim(klGrp(klGrpI))
            If klSingle = "" Then GoTo NaechsteKlGrp
            ' Als kl setzen und weiter verarbeiten
            kl = klSingle
            Dim klFound As Boolean: klFound = False
            Dim kli As Long
        For kli = 1 To allKlAnz
            If allKlassen(kli) = kl Then klFound = True: Exit For
        Next kli
        If Not klFound Then
            allKlAnz = allKlAnz + 1
            allKlassen(allKlAnz) = kl
        End If
NaechsteKlGrp:
        Next klGrpI
NaechsteKl2:
    Next r

    ' Jahrgangsstufe einer Klasse bestimmen (erste Ziffer(n))
    ' z.B. "05A" -> 5, "10B" -> 10, "7C" -> 7
    ' ---- Ausgabe schreiben ----
    ' Header
    Dim cH As Long: cH = RGB(31, 73, 125)
    With wsOut.Cells(1, 1)
        .Value = "Stundentafel-Pruefung: Abweichungen"
        .Font.Bold = True: .Font.Size = 13
        .Interior.Color = cH: .Font.Color = RGB(255, 255, 255)
    End With
    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, 5)).Merge

    wsOut.Cells(2, 1).Value = "Klasse"
    wsOut.Cells(2, 2).Value = "Fach"
    wsOut.Cells(2, 3).Value = "Soll (WSt)"
    wsOut.Cells(2, 4).Value = "Ist (WSt)"
    wsOut.Cells(2, 5).Value = "Differenz"
    Dim cSH As Long: cSH = RGB(68, 114, 196)
    Dim col As Long
    For col = 1 To 5
        wsOut.Cells(2, col).Font.Bold = True
        wsOut.Cells(2, col).Interior.Color = cSH
        wsOut.Cells(2, col).Font.Color = RGB(255, 255, 255)
    Next col

    Dim outRow As Long: outRow = 3
    Dim fehlend As Long: fehlend = 0
    Dim abweich As Long: abweich = 0

    ' Fuer jede Klasse: Jahrgangsstufe bestimmen, Soll-Stunden aus Stundentafel
    For kli = 1 To allKlAnz
        Dim klName As String: klName = allKlassen(kli)

        ' Sammelklassen ignorieren: enthalten nur Buchstaben nach der Zahl (z.B. 5a, 5b, 10k)
        ' Echte Klassen haben Format: 05A, 05B, 7C, 10E etc. (Grossbuchstabe nach Zahl)
        ' Sammelklassen: 5a, 5b, EF, Q1, Q2 oder enthalten Sonderzeichen
        Dim klLower As String: klLower = LCase(klName)
        ' Pruefen ob es eine echte Klasse ist: Zahl gefolgt von Grossbuchstabe
        Dim istEchteKlasse As Boolean: istEchteKlasse = False
        Dim cki As Long
        For cki = 1 To Len(klName)
            If Mid(klName, cki, 1) >= "0" And Mid(klName, cki, 1) <= "9" Then
                ' Zahl gefunden - jetzt Buchstabe danach pruefen
                If cki < Len(klName) Then
                    Dim nextCh As String: nextCh = Mid(klName, cki + 1, 1)
                    ' Noch eine Zahl? Weiterlesen
                    If nextCh >= "0" And nextCh <= "9" Then GoTo NaechstesCki
                    ' Grossbuchstabe = echte Klasse (05A, 7C, 10E)
                    If nextCh >= "A" And nextCh <= "Z" Then istEchteKlasse = True: Exit For
                    ' Kleinbuchstabe = Sammelklasse (5a, 5b) -> ignorieren
                    If nextCh >= "a" And nextCh <= "z" Then istEchteKlasse = False: Exit For
                End If
            End If
NaechstesCki:
        Next cki
        If Not istEchteKlasse Then GoTo NaechsteKlasse2

        ' Jahrgangsstufe extrahieren (nur aus erstem Klassennamen vor Komma)
        Dim klFirst As String: klFirst = klName
        Dim commaPos As Long: commaPos = InStr(klName, ",")
        If commaPos > 0 Then klFirst = left(klName, commaPos - 1)
        ' Zusammengesetzte Klassennamen (05A,05B) ueberspringen
        If commaPos > 0 Then GoTo NaechsteKlasse2
        ' Jahrgangsstufe nur aus echten Klassennamen (z.B. 05A, 7B, 10C)
        ' Klassen wie "AG7", "EF", "Q1" werden ignoriert
        If Not (klFirst Like "[0-9]*" Or klFirst Like "0[0-9]*") Then GoTo NaechsteKlasse2
        Dim jhgStr As String: jhgStr = ""
        Dim ci As Long
        For ci = 1 To Len(klFirst)
            If Mid(klFirst, ci, 1) >= "0" And Mid(klFirst, ci, 1) <= "9" Then
                jhgStr = jhgStr & Mid(klFirst, ci, 1)
            Else
                If jhgStr <> "" Then Exit For
            End If
        Next ci
        If Not IsNumeric(jhgStr) Then GoTo NaechsteKlasse2
        Dim jhg As Long: jhg = CLng(jhgStr)
        ' Nur Jahrgangstufen 5-10 beruecksichtigen
        If jhg < 5 Or jhg > 10 Then GoTo NaechsteKlasse2

        ' Alle Soll-Eintraege fuer diese Jahrgangsstufe
        Dim si As Long
        For si = 1 To stAnz
            If stJhg(si) <> jhg Then GoTo NaechsterSoll

            Dim sollFach As String: sollFach = stFach(si)
            Dim soll As Double: soll = stWSt(si)

            ' Wahlpflichtunterricht -> separat als L9 K1, EK K1, BI K1, IF K1, L9 K2, EK K2 ausgeben
            If LCase(sollFach) = LCase("Wahlpflichtunterricht") Then
                Dim wpuFaecher(1 To 6) As String
                wpuFaecher(1) = "L9 K1": wpuFaecher(2) = "EK K1"
                wpuFaecher(3) = "BI K1": wpuFaecher(4) = "IF K1"
                wpuFaecher(5) = "L9 K2": wpuFaecher(6) = "EK K2"
                Dim wpuI As Long
                For wpuI = 1 To 6
                    Dim wpuKey As String: wpuKey = LCase(klName) & "|" & LCase(wpuFaecher(wpuI))
                    Dim wpuIst As Double: wpuIst = 0
                    For ki = 1 To istAnz
                        If istKeys(ki) = wpuKey Then wpuIst = istWStArr(ki): Exit For
                    Next ki
                    ' Auch via StTafelFachName suchen (falls als WPU gespeichert)
                    If wpuIst = 0 Then
                        Dim wpuKey2 As String: wpuKey2 = LCase(klName) & "|wpu"
                        For ki = 1 To istAnz
                            If istKeys(ki) = wpuKey2 Then wpuIst = istWStArr(ki): Exit For
                        Next ki
                    End If
                    Dim wpuDiff As Double: wpuDiff = wpuIst - soll
                    If Abs(wpuDiff) > 0.1 Then
                        Dim wpuColor As Long
                        If wpuDiff < 0 Then
                            wpuColor = RGB(220, 80, 80)
                            If wpuIst = 0 Then fehlend = fehlend + 1 Else abweich = abweich + 1
                        Else
                            wpuColor = RGB(255, 235, 156): abweich = abweich + 1
                        End If
                        wsOut.Cells(outRow, 1).Value = klName
                        wsOut.Cells(outRow, 2).Value = wpuFaecher(wpuI)
                        wsOut.Cells(outRow, 3).Value = soll
                        wsOut.Cells(outRow, 4).Value = wpuIst
                        wsOut.Cells(outRow, 5).Value = wpuDiff
                        wsOut.Cells(outRow, 5).NumberFormat = "+0.0;-0.0;0"
                        For col = 1 To 5
                            wsOut.Cells(outRow, col).Interior.Color = wpuColor
                        Next col
                        outRow = outRow + 1
                    End If
                Next wpuI
                GoTo NaechsterSoll
            End If

            ' Religionslehre/PP -> separat als ER, KR, PP ausgeben (nicht als Religionslehre/PP)
            If LCase(sollFach) = LCase("Religionslehre/PP") Then
                Dim relFaecher(1 To 3) As String
                relFaecher(1) = "ER": relFaecher(2) = "KR": relFaecher(3) = "PP"
                Dim rfi As Long
                For rfi = 1 To 3
                    Dim relKey As String: relKey = LCase(klName) & "|" & LCase(relFaecher(rfi))
                    Dim relIst As Double: relIst = 0
                    For ki = 1 To istAnz
                        If istKeys(ki) = relKey Then relIst = istWStArr(ki): Exit For
                    Next ki
                    ' Auch via StTafelFachName suchen
                    If relIst = 0 Then
                        For ki = 1 To istAnz
                            Dim rKey2 As String: rKey2 = istKeys(ki)
                            Dim rSep As Long: rSep = InStr(rKey2, "|")
                            If rSep = 0 Then GoTo NaechsterRelKey
                            If LCase(left(rKey2, rSep - 1)) <> LCase(klName) Then GoTo NaechsterRelKey
                            Dim rFachK As String: rFachK = Mid(rKey2, rSep + 1)
                            If LCase(StTafelFachName(rFachK)) = LCase(sollFach) Then
                                relIst = relIst + istWStArr(ki)
                            End If
NaechsterRelKey:
                        Next ki
                    End If
                    Dim relDiff As Double: relDiff = relIst - soll
                    If Abs(relDiff) > 0.1 Then
                        Dim relColor As Long
                        If relDiff < 0 Then
                            relColor = RGB(220, 80, 80)   ' Dunkelrot: zu wenig
                            If relIst = 0 Then fehlend = fehlend + 1 Else abweich = abweich + 1
                        Else
                            relColor = RGB(255, 235, 156) ' Gelb: zu viel
                            abweich = abweich + 1
                        End If
                        wsOut.Cells(outRow, 1).Value = klName
                        wsOut.Cells(outRow, 2).Value = relFaecher(rfi)
                        wsOut.Cells(outRow, 3).Value = soll
                        wsOut.Cells(outRow, 4).Value = relIst
                        wsOut.Cells(outRow, 5).Value = relDiff
                        wsOut.Cells(outRow, 5).NumberFormat = "+0.0;-0.0;0"
                        For col = 1 To 5
                            wsOut.Cells(outRow, col).Interior.Color = relColor
                        Next col
                        outRow = outRow + 1
                    End If
                Next rfi
                GoTo NaechsterSoll
            End If

            ' Ist-Stunden fuer diese Klasse+Fach suchen
            ' Klassen-Tabelle kann Kuerzel enthalten (z.B. "D", "BI") oder Langnamen
            ' Stundentafel hat Langnamen -> Kuerzel per StTafelFachName mappen
            Dim istKey As String: istKey = LCase(klName) & "|" & LCase(sollFach)
            Dim ist As Double: ist = 0
            For ki = 1 To istAnz
                If istKeys(ki) = istKey Then ist = istWStArr(ki): Exit For
            Next ki
            ' Falls nicht gefunden: alle Eintraege pruefen ob Kuerzel auf sollFach mappt
            If ist = 0 Then
                For ki = 1 To istAnz
                    ' Schluessel hat Form "klasse|fach"
                    Dim kKey As String: kKey = istKeys(ki)
                    Dim pSep As Long: pSep = InStr(kKey, "|")
                    If pSep = 0 Then GoTo NaechsterKey
                    If LCase(left(kKey, pSep - 1)) <> LCase(klName) Then GoTo NaechsterKey
                    Dim kFach As String: kFach = Mid(kKey, pSep + 1)
                    If LCase(StTafelFachName(kFach)) = LCase(sollFach) Then
                        ist = ist + istWStArr(ki)  ' summieren falls mehrere Kuerzel passen
                    End If
NaechsterKey:
                Next ki
            End If

            Dim diff As Double: diff = ist - soll
            ' Nur Abweichungen ausgeben (Toleranz: 0.1)
            If Abs(diff) > 0.1 Then
                Dim bgColor As Long
                If diff < 0 Then
                    bgColor = RGB(220, 80, 80)    ' Dunkelrot: zu wenig (auch komplett fehlend)
                    If ist = 0 Then fehlend = fehlend + 1 Else abweich = abweich + 1
                Else
                    bgColor = RGB(255, 235, 156)  ' Gelb: zu viel
                    abweich = abweich + 1
                End If

                wsOut.Cells(outRow, 1).Value = klName
                wsOut.Cells(outRow, 2).Value = sollFach
                wsOut.Cells(outRow, 3).Value = soll
                wsOut.Cells(outRow, 4).Value = ist
                wsOut.Cells(outRow, 5).Value = diff
                wsOut.Cells(outRow, 5).NumberFormat = "+0.0;-0.0;0"
                For col = 1 To 5
                    wsOut.Cells(outRow, col).Interior.Color = bgColor
                Next col
                outRow = outRow + 1
            End If
NaechsterSoll:
        Next si
NaechsteKlasse2:
    Next kli

    ' Spaltenbreiten
    wsOut.Columns(1).ColumnWidth = 10
    wsOut.Columns(2).ColumnWidth = 25
    wsOut.Columns(3).ColumnWidth = 12
    wsOut.Columns(4).ColumnWidth = 12
    wsOut.Columns(5).ColumnWidth = 12

    ' Zusammenfassung
    outRow = outRow + 1
    wsOut.Cells(outRow, 1).Value = fehlend & " fehlende Faecher, " & abweich & " Abweichungen"
    wsOut.Cells(outRow, 1).Font.Bold = True
    wsOut.Range(wsOut.Cells(outRow, 1), wsOut.Cells(outRow, 5)).Merge

    wsOut.Activate

    MsgBox fehlend & " fehlende Faecher (rot)" & vbCrLf & _
           abweich & " Stundenabweichungen (gelb/gruen)" & vbCrLf & vbCrLf & _
           "Ergebnis in Sheet 'Stundentafel_Pruefung'.", _
           vbInformation, "Stundentafel-Pruefung"
End Sub

' ============================================================
' Hilfsfunktion: Fachkuerzel -> Stundentafel-Fachname

' ============================================================
' Hilfsfunktion: Fachname -> Engpass-Gruppenname
' Zentrale Zuordnung fuer Diagnose und Engpass-Analyse
' ============================================================
Function FachEngpassKey(fach As String) As String
    Dim f As String: f = Trim(fach)
    Do While InStr(f, "  ") > 0: f = Replace(f, "  ", " "): Loop
    If f = "" Then FachEngpassKey = "": Exit Function
    If g_fgAnz = 0 Then Call FachgruppenLaden
    Dim fi As Long
    For fi = 1 To g_fgAnz
        If LCase(g_fgFach(fi)) = LCase(f) Then
            FachEngpassKey = g_fgGruppe(fi): Exit Function
        End If
    Next fi
    Dim pos As Long: pos = InStr(f, " ")
    If pos > 1 Then
        Dim prx As String: prx = Trim(left(f, pos - 1))
        For fi = 1 To g_fgAnz
            If LCase(g_fgFach(fi)) = LCase(prx) Then
                FachEngpassKey = g_fgGruppe(fi): Exit Function
            End If
        Next fi
    End If
    FachEngpassKey = ""
End Function
Function StTafelFachName(kuerzel As String) As String
    Dim k As String
    Dim pos As Long: pos = InStr(kuerzel, " ")
    If pos > 0 Then
        k = UCase(Trim(left(kuerzel, pos - 1)))
    Else
        k = UCase(Trim(kuerzel))
    End If
    ' Wahlpflichtfach-Kurse: L9 K1, EK K1, BI K1, IF K1, L9 K2, EK K2 etc.
    If InStr(UCase(kuerzel), " K1") > 0 Or InStr(UCase(kuerzel), " K2") > 0 Then
        If InStr(UCase(kuerzel), "ERG") = 0 Then
            StTafelFachName = "Wahlpflichtunterricht": Exit Function
        End If
    End If
    Select Case k
        Case "D":               StTafelFachName = "Deutsch"
        Case "M":               StTafelFachName = "Mathematik"
        Case "E":               StTafelFachName = "Englisch"
        Case "F":               StTafelFachName = "2. Fremdsprache: Franz" & Chr(246) & "sisch"
        Case "BI":              StTafelFachName = "Biologie"
        Case "CH":              StTafelFachName = "Chemie"
        Case "PH":              StTafelFachName = "Physik"
        Case "IF":              StTafelFachName = "Informatik"
        Case "GE":              StTafelFachName = "Geschichte"
        Case "EK":              StTafelFachName = "Erdkunde"
        Case "WIPO", "PO", "WI": StTafelFachName = "Politik"
        Case "KU":              StTafelFachName = "Kunst"
        Case "MU":              StTafelFachName = "Musik"
        Case "ER", "KR", "PP", "RE", "REL": StTafelFachName = "Religionslehre/PP"
        Case "SP", "SWI", "SWI1", "SWI2": StTafelFachName = "Sport"
        Case "WP", "WPU", "WPI": StTafelFachName = "Wahlpflichtunterricht"
        Case Else:              StTafelFachName = kuerzel
    End Select
End Function

' ============================================================
' Button "Stundentafel pruefen" in Sheet Stundentafel_Pruefung erstellen
' ============================================================
Sub StundentafelPruefung_ButtonErstellen()
    Dim wsOut As Worksheet
    Set wsOut = SheetByName("Stundentafel_Pruefung")
    If wsOut Is Nothing Then
        ' Sheet anlegen falls nicht vorhanden
        Set wsOut = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsOut.name = "Stundentafel_Pruefung"
    End If

    ' Alte Buttons entfernen
    Dim shp As Shape
    For Each shp In wsOut.Shapes
        If shp.OnAction = "StundentafelVergleich" Then shp.Delete
    Next shp

    ' Position: Spalte F, Zeile 1
    Dim leftPos  As Double: leftPos = wsOut.Columns("F").left
    Dim topPos   As Double: topPos = wsOut.rows(1).top + 2
    Dim btnW     As Double: btnW = 170
    Dim btnH     As Double: btnH = 22

    Dim btn As Shape
    Set btn = wsOut.Shapes.AddFormControl(xlButtonControl, leftPos, topPos, btnW, btnH)
    btn.TextFrame.Characters.Text = "> Stundentafel pr" & Chr(252) & "fen"
    btn.TextFrame.Characters.Font.Bold = True
    btn.TextFrame.Characters.Font.Size = 10
    btn.OnAction = "StundentafelVergleich"

    wsOut.Activate
    MsgBox "Button erstellt in Sheet 'Stundentafel_Pruefung', Spalte F.", vbInformation
End Sub

' ============================================================
' Fach-Engpass-Analyse (unabhaengig vom Optimierungslauf)
' Berechnet verfuegbare Kapazitaet direkt aus Tabelle Klassen
' ============================================================
Sub FachEngpassAnalyse()
    ' Alle Variablendeklarationen am Anfang (VBA-Regel)
    Dim wsUV  As Worksheet
    Dim wsK   As Worksheet
    Dim wsL   As Worksheet
    Dim wsOut As Worksheet
    Dim shpE  As Shape
    Dim btnE  As Shape
    Dim cH As Long, cSH As Long, cFe As Long, cOr As Long, cWa As Long, cOK As Long
    Dim zRow As Long, lr As Long, kr As Long, li As Long, fj As Long, fj2 As Long
    Dim lN As String, lS As Double
    Dim lName(1 To 300) As String, lSoll(1 To 300) As Double
    Dim lIst(1 To 300)  As Double, lRow(1 To 300)  As Long
    Dim lAnz As Long, lMaxR As Long, kMaxR As Long
    Dim kKE As String, kFE As String, kWE As Double, kL1 As String, kL2 As String
    Dim aktKE As String
    Dim uvHR As Long, uvCFF As Long, uvCWF As Long, uvCUF As Long, uvCKF As Long
    Dim uvMR As Long, uvR As Long, uvC As Long, uvHV As String
    Dim uvF As String, uvFU2 As String, uvKlE As String, uvWE As Double
    Dim uvUN As String, sk As String, skFound As Boolean, eaSi As Long
    Dim uvFN2 As String, uvPx As String, uvKy As String, ufi As Long, ufk As Long
    Dim spPosUV As Long
    Dim ergPrx As String, ergSp As Long
    Dim verfE As Double, lpE As String, lfE As String, lkE As String
    Dim rfiE As Long, rfkE As Long, acE As Boolean
    Dim lf2E As String, lp2E As String, lk2E As String
    Dim eaI As Long, eaJ As Long, eaSc1 As Double, eaSc2 As Double
    Dim eaTF As String, eaTB As Double, eaTK As Double, eaTQ As Long
    Dim scE As Double, engE As String, ecol As Long
    Dim wpuI As Long, wpuBed As Double, wpuQual As Long, wpuKap As Double
    Dim wpuKurse(1 To 6) As String, wpuFaecherMap(1 To 6) As String
    Dim wpuSeen(1 To 2000) As String, wpuSeenAnz As Long
    Dim wpuScore As Double, wpuEng As String, wpuColor As Long
    Dim NC3dummy As Long  ' placeholder

    Set wsUV = SheetByName("KlassenUV")
    Set wsK = SheetByName("Klassen")
    Set wsL = SheetByName("Lehrerliste")
    If wsUV Is Nothing Then MsgBox "Sheet 'KlassenUV' fehlt!", vbCritical: Exit Sub
    If wsK Is Nothing Then MsgBox "Sheet 'Klassen' fehlt!", vbCritical: Exit Sub
    If wsL Is Nothing Then MsgBox "Sheet 'Lehrerliste' fehlt!", vbCritical: Exit Sub

    ' ---- Ausgabe-Sheet ----
    Set wsOut = SheetByName("Engpass_Analyse")
    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsOut.name = "Engpass_Analyse"
    Else
        wsOut.Cells.Clear
        For Each shpE In wsOut.Shapes: shpE.Delete: Next shpE
    End If

    ' Button neu erstellen
    Set btnE = wsOut.Shapes.AddFormControl(xlButtonControl, wsOut.Columns("F").left, wsOut.rows(1).top + 2, 170, 22)
    btnE.TextFrame.Characters.Text = "> Engpass neu berechnen"
    btnE.TextFrame.Characters.Font.Bold = True
    btnE.TextFrame.Characters.Font.Size = 10
    btnE.OnAction = "FachEngpassAnalyse"

    cH = RGB(31, 73, 125)
    cSH = RGB(68, 114, 196)
    cFe = RGB(220, 80, 80)
    cOr = RGB(255, 150, 80)
    cWa = RGB(255, 235, 156)
    cOK = RGB(198, 239, 206)
    zRow = 1

    Call SectionHeader(wsOut, zRow, "FACH-ENGPASS-ANALYSE", cH, 8): zRow = zRow + 2
    wsOut.Cells(zRow, 1).Value = "Verfuegbare Kapazitaet = Soll-WSt minus bereits in Tabelle Klassen eingetragene Stunden"
    wsOut.Cells(zRow, 1).Font.Italic = True: zRow = zRow + 2

    ' ---- Schritt 1: IstWSt + Bedarf aus Tabelle Klassen ----
    ' IstWSt: Summe Wert (Sp.9) pro Lehrer wenn in Sp.6 (Fix) eingetragen
    ' Bedarf:  Summe Wert (Sp.9) pro Fach aus allen Zeilen
    lMaxR = wsL.Cells(wsL.rows.Count, 1).End(xlUp).row
    kMaxR = wsK.Cells(wsK.rows.Count, 2).End(xlUp).row

    ' Lehrer-Namen und SollWSt aus Lehrerliste einlesen
    Const MAXL As Long = 300
    lAnz = 0
    For lr = 2 To lMaxR
        lN = Trim(CStr(wsL.Cells(lr, 1).Value))
        If lN = "" Then GoTo NaechsterLR
        lS = 0
        If IsNumeric(wsL.Cells(lr, 19).Value) Then lS = CDbl(wsL.Cells(lr, 19).Value)
        If lS <= 0 Then GoTo NaechsterLR
        lAnz = lAnz + 1
        lName(lAnz) = lN
        lSoll(lAnz) = lS
        lIst(lAnz) = 0
        lRow(lAnz) = lr
NaechsterLR:
    Next lr

    ' IstWSt aus Fix-Spalte (Sp.6), Bedarf aus Wert-Spalte (Sp.9)
    Const MAXF As Long = 200
    Dim fFach(1 To MAXF)  As String
    Dim fBed(1 To MAXF)   As Double
    Dim fKap(1 To MAXF)   As Double
    Dim fQual(1 To MAXF)  As Long
    Dim fAnz As Long: fAnz = 0

    aktKE = ""
    For kr = 2 To kMaxR
        kKE = Trim(CStr(wsK.Cells(kr, 1).Value))
        kFE = Trim(CStr(wsK.Cells(kr, 2).Value))
        If kKE <> "" Then aktKE = kKE
        If kFE = "" Then GoTo NaechsteKRE
        ' Wert aus Spalte 9
        Dim kWert As Double: kWert = 0
        If IsNumeric(wsK.Cells(kr, 9).Value) Then kWert = CDbl(wsK.Cells(kr, 9).Value)
        ' Fix-Lehrer aus Sp.6
        Dim kFix As String: kFix = Trim(CStr(wsK.Cells(kr, 6).Value))
        Dim kIstFix As Boolean: kIstFix = (kFix <> "" And kFix <> "?")
        ' IstWSt: nur wenn Fix-Lehrer eingetragen
        If kIstFix And kWert > 0 Then
            For li = 1 To lAnz
                If LCase(lName(li)) = LCase(kFix) Then lIst(li) = lIst(li) + kWert: Exit For
            Next li
        End If
        ' Bedarf: nur nicht-fixierte Zeilen (Sp.6 leer oder ?)
        Dim kFKey As String: kFKey = FachEngpassKey(kFE)
        If kFKey <> "" And kWert > 0 And Not kIstFix Then
            Dim kFI As Long: kFI = 0
            Dim kFK As Long
            For kFK = 1 To fAnz
                If LCase(fFach(kFK)) = LCase(kFKey) Then kFI = kFK: Exit For
            Next kFK
            If kFI = 0 Then fAnz = fAnz + 1: kFI = fAnz: fFach(kFI) = kFKey
            fBed(kFI) = fBed(kFI) + kWert
        End If
NaechsteKRE:
    Next kr

    ' ---- Schritt 2 (Bedarf bereits aus Klassen geladen) -> Schritt 3: Kapazitaet ----
    ' ---- Schritt 3: Kapazitaet pro Fach aus Lehrerliste ----
    For li = 1 To lAnz
        verfE = lSoll(li) - lIst(li)
        If verfE < 0 Then verfE = 0
        For fj = 1 To 16
            lfE = Trim(wsL.Cells(lRow(li), 1 + fj).Value)
            If lfE = "" Then GoTo NaechsterLFE
            Do While InStr(lfE, "  ") > 0: lfE = Join(Split(lfE, "  "), " "): Loop
            lpE = FachGruppenPraefix(lfE)
            If lpE = "" Then lpE = FachLevelPraefix(lfE)
            If lpE = "" Then lpE = FachSuffixPraefix(lfE)
            lkE = IIf(lpE <> "", lpE, lfE)
            rfiE = 0
            For rfkE = 1 To fAnz
                If LCase(fFach(rfkE)) = LCase(lkE) Then rfiE = rfkE: Exit For
            Next rfkE
            If rfiE > 0 Then
                ' Lehrer nur einmal pro Fach zaehlen
                acE = False
                For fj2 = 1 To fj - 1
                    lf2E = Trim(wsL.Cells(lRow(li), 1 + fj2).Value)
                    If lf2E = "" Then GoTo NC2
                    Do While InStr(lf2E, "  ") > 0: lf2E = Join(Split(lf2E, "  "), " "): Loop
                    lp2E = FachGruppenPraefix(lf2E)
                    If lp2E = "" Then lp2E = FachLevelPraefix(lf2E)
                    If lp2E = "" Then lp2E = FachSuffixPraefix(lf2E)
                    lk2E = IIf(lp2E <> "", lp2E, lf2E)
                    If LCase(lk2E) = LCase(lkE) Then acE = True: Exit For
NC2:
                Next fj2
                If Not acE Then
                    fQual(rfiE) = fQual(rfiE) + 1
                    fKap(rfiE) = fKap(rfiE) + verfE
                End If
            End If
NaechsterLFE:
        Next fj
    Next li

    ' ---- Schritt 4: Sortieren nach Score = Bedarf / max(0.1, Kapazitaet) ----
    For eaI = 1 To fAnz - 1
        For eaJ = eaI + 1 To fAnz
            eaSc1 = fBed(eaI) / IIf(fKap(eaI) > 0.1, fKap(eaI), 0.1)
            eaSc2 = fBed(eaJ) / IIf(fKap(eaJ) > 0.1, fKap(eaJ), 0.1)
            If eaSc2 > eaSc1 Then
                eaTF = fFach(eaI): fFach(eaI) = fFach(eaJ): fFach(eaJ) = eaTF
                eaTB = fBed(eaI):  fBed(eaI) = fBed(eaJ):   fBed(eaJ) = eaTB
                eaTK = fKap(eaI):  fKap(eaI) = fKap(eaJ):   fKap(eaJ) = eaTK
                eaTQ = fQual(eaI): fQual(eaI) = fQual(eaJ): fQual(eaJ) = eaTQ
            End If
        Next eaJ
    Next eaI

    ' ---- Schritt 5: Ausgabe ----
    Call TableHeader(wsOut, zRow, Array("Fach", "Bedarf (WSt)", "Qual. Lehrer", "Freie Kap. (WSt)", "Score (B/K)", "Engpass"), cSH)
    zRow = zRow + 1

    For eaI = 1 To fAnz
        scE = fBed(eaI) / IIf(fKap(eaI) > 0.1, fKap(eaI), 0.1)
        If scE > 2 Then
            engE = "Kritisch (" & Format(scE, "0.0") & "x)"
        ElseIf scE > 1 Then
            engE = "Knapp (" & Format(scE, "0.0") & "x)"
        ElseIf scE > 0.5 Then
            engE = "OK (" & Format(scE, "0.0") & "x)"
        Else
            engE = "Puffer (" & Format(scE, "0.0") & "x)"
        End If
        wsOut.Cells(zRow, 1).Value = fFach(eaI)
        wsOut.Cells(zRow, 2).Value = fBed(eaI)
        wsOut.Cells(zRow, 3).Value = fQual(eaI)
        wsOut.Cells(zRow, 4).Value = fKap(eaI)
        wsOut.Cells(zRow, 5).Value = scE
        wsOut.Cells(zRow, 6).Value = engE
        wsOut.Cells(zRow, 5).NumberFormat = "0.00"
        If scE > 2 Then
            ecol = cFe
        ElseIf scE > 1 Then
            ecol = cOr
        ElseIf scE > 0.5 Then
            ecol = cWa
        Else
            ecol = cOK
        End If
        Call FarbeZeile(wsOut, zRow, 1, 6, ecol)
        zRow = zRow + 1
    Next eaI

    wsOut.Columns(1).ColumnWidth = 14: wsOut.Columns(2).ColumnWidth = 13
    wsOut.Columns(3).ColumnWidth = 13: wsOut.Columns(4).ColumnWidth = 16
    wsOut.Columns(5).ColumnWidth = 12: wsOut.Columns(6).ColumnWidth = 22

    ' ---- WPU-Detailanalyse ----
    zRow = zRow + 2
    Call SectionHeader(wsOut, zRow, "WAHLPFLICHTUNTERRICHT - Detailanalyse (L9/EK/BI/IF/SP inkl. K1/K2/ERG)", cH, 8): zRow = zRow + 1
    Call TableHeader(wsOut, zRow, Array("WPU-Kurs", "Bedarf (WSt)", "Qual. Lehrer", "Freie Kap. (WSt)", "Score (B/K)", "Engpass"), cSH)
    zRow = zRow + 1

    ' WPU-Kurse und ihr Fach-Mapping (Kurs -> Fach das Lehrer koennen muessen)
    wpuKurse(1) = "L9 K1":  wpuFaecherMap(1) = "L9"
    wpuKurse(2) = "EK K1":  wpuFaecherMap(2) = "EK"
    wpuKurse(3) = "BI K1":  wpuFaecherMap(3) = "BI"
    wpuKurse(4) = "IF K1":  wpuFaecherMap(4) = "IF"
    wpuKurse(5) = "L9 K2":  wpuFaecherMap(5) = "L9"
    wpuKurse(6) = "EK K2":  wpuFaecherMap(6) = "EK"

    For wpuI = 1 To 6
        ' Bedarf: Wst aus KlassenUV fuer diesen Kurs
        wpuBed = 0
        wpuQual = 0
        wpuKap = 0

        ' Bedarf aus KlassenUV
        wpuSeenAnz = 0
        For uvR = uvHR + 1 To uvMR
            uvF = Trim(CStr(wsUV.Cells(uvR, uvCFF).Value))
            If LCase(uvF) <> LCase(wpuKurse(wpuI)) Then GoTo NaechsteWPUZeile
            uvKlE = ""
            If uvCKF > 0 Then uvKlE = Trim(CStr(wsUV.Cells(uvR, uvCKF).Value))
            If uvKlE = "" Then GoTo NaechsteWPUZeile
            If Not IsNumeric(wsUV.Cells(uvR, uvCWF).Value) Then GoTo NaechsteWPUZeile
            uvWE = CDbl(wsUV.Cells(uvR, uvCWF).Value)
            If uvWE <= 0 Then GoTo NaechsteWPUZeile
            ' Deduplizierung
            If uvCUF > 0 Then
                uvUN = Trim(CStr(wsUV.Cells(uvR, uvCUF).Value))
                If uvUN <> "" Then
                    sk = uvUN & "|" & LCase(uvKlE) & "|" & LCase(uvF)
                    skFound = False
                    For eaSi = 1 To wpuSeenAnz
                        If wpuSeen(eaSi) = sk Then skFound = True: Exit For
                    Next eaSi
                    If skFound Then GoTo NaechsteWPUZeile
                    wpuSeenAnz = wpuSeenAnz + 1: wpuSeen(wpuSeenAnz) = sk
                End If
            End If
            wpuBed = wpuBed + uvWE
NaechsteWPUZeile:
        Next uvR

        ' Qualifizierte Lehrer: koennen das Fach wpuFaecherMap(wpuI)
        For li = 1 To lAnz
            verfE = lSoll(li) - lIst(li)
            If verfE < 0 Then verfE = 0
            For fj = 1 To 16
                lfE = Trim(wsL.Cells(lRow(li), 1 + fj).Value)
                If lfE = "" Then GoTo NaechsterWPULF
                Do While InStr(lfE, "  ") > 0: lfE = Join(Split(lfE, "  "), " "): Loop
                lpE = FachGruppenPraefix(lfE)
                If lpE = "" Then lpE = FachLevelPraefix(lfE)
                If lpE = "" Then lpE = FachSuffixPraefix(lfE)
                lkE = IIf(lpE <> "", lpE, lfE)
                If LCase(lkE) = LCase(wpuFaecherMap(wpuI)) Then
                    ' Lehrer nur einmal zaehlen
                    acE = False
                    For fj2 = 1 To fj - 1
                        lf2E = Trim(wsL.Cells(lRow(li), 1 + fj2).Value)
                        If lf2E = "" Then GoTo NC3
                        Do While InStr(lf2E, "  ") > 0: lf2E = Join(Split(lf2E, "  "), " "): Loop
                        lp2E = FachGruppenPraefix(lf2E)
                        If lp2E = "" Then lp2E = FachLevelPraefix(lf2E)
                        If lp2E = "" Then lp2E = FachSuffixPraefix(lf2E)
                        lk2E = IIf(lp2E <> "", lp2E, lf2E)
                        If LCase(lk2E) = LCase(wpuFaecherMap(wpuI)) Then acE = True: Exit For
NC3:
                    Next fj2
                    If Not acE Then
                        wpuQual = wpuQual + 1
                        wpuKap = wpuKap + verfE
                    End If
                    Exit For
                End If
NaechsterWPULF:
            Next fj
        Next li

        ' Ausgabe
        wpuScore = wpuBed / IIf(wpuKap > 0.1, wpuKap, 0.1)
        If wpuScore > 2 Then
            wpuEng = "Kritisch (" & Format(wpuScore, "0.0") & "x)"
        ElseIf wpuScore > 1 Then
            wpuEng = "Knapp (" & Format(wpuScore, "0.0") & "x)"
        ElseIf wpuScore > 0.5 Then
            wpuEng = "OK (" & Format(wpuScore, "0.0") & "x)"
        Else
            wpuEng = "Puffer (" & Format(wpuScore, "0.0") & "x)"
        End If
        wsOut.Cells(zRow, 1).Value = wpuKurse(wpuI)
        wsOut.Cells(zRow, 2).Value = wpuBed
        wsOut.Cells(zRow, 3).Value = wpuQual
        wsOut.Cells(zRow, 4).Value = wpuKap
        wsOut.Cells(zRow, 5).Value = wpuScore
        wsOut.Cells(zRow, 6).Value = wpuEng
        wsOut.Cells(zRow, 5).NumberFormat = "0.00"
        If wpuScore > 2 Then
            wpuColor = cFe
        ElseIf wpuScore > 1 Then
            wpuColor = cOr
        ElseIf wpuScore > 0.5 Then
            wpuColor = cWa
        Else
            wpuColor = cOK
        End If
        Call FarbeZeile(wsOut, zRow, 1, 6, wpuColor)
        zRow = zRow + 1
    Next wpuI

    wsOut.Activate
    wsOut.Cells(1, 1).Select
    MsgBox "Engpass-Analyse abgeschlossen (" & fAnz & " F" & Chr(228) & "cher).", vbInformation
End Sub
