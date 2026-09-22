Option Explicit

Private Const RunName As String = "run"
Private Const ResetSpaltenName As String = "ResetSpalten"
Private Const ResetZeilenName As String = "ResetZeilen"

Private Const ConfigBlattName As String = "Config"
Private Const RunsTabellenName As String = "tblRuns"


'==========================================================
' Reagiert auf Änderungen der Steuerzelle
'==========================================================
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim rngRun As Range

    Set rngRun = GetNamedRange(Me, RunName)

    If rngRun Is Nothing Then Exit Sub
    If Intersect(Target, rngRun) Is Nothing Then Exit Sub

    RunAnsichtAktualisieren

End Sub


'==========================================================
' Aktualisiert beim Aktivieren des Blatts
'==========================================================
Private Sub Worksheet_Activate()

    RunAnsichtAktualisieren

End Sub


'==========================================================
' Hauptsteuerung
'==========================================================
Public Sub RunAnsichtAktualisieren()

    Dim wsConfig As Worksheet
    Dim tblRuns As ListObject
    Dim runZeile As ListRow

    Dim rngRun As Range
    Dim rngResetSpalten As Range
    Dim rngResetZeilen As Range

    Dim runWert As Long

    Dim spaltenListe As String
    Dim zeilenListe As String
    Dim gruppenZu As String
    Dim gruppenAuf As String
    Dim startGruppen As String

    Dim eventsWarenAktiv As Boolean
    Dim screenUpdatingWarAktiv As Boolean
    Dim alterBerechnungsmodus As XlCalculation

    On Error GoTo Fehler

    eventsWarenAktiv = Application.EnableEvents
    screenUpdatingWarAktiv = Application.ScreenUpdating
    alterBerechnungsmodus = Application.Calculation

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set rngRun = GetNamedRange(Me, RunName)
    Set rngResetSpalten = GetNamedRange(Me, ResetSpaltenName)
    Set rngResetZeilen = GetNamedRange(Me, ResetZeilenName)

    If rngRun Is Nothing Then
        Err.Raise vbObjectError + 1000, _
                  "RunAnsichtAktualisieren", _
                  "Der benannte Bereich '" & RunName & _
                  "' wurde nicht gefunden."
    End If

    If rngRun.Cells.CountLarge <> 1 Then
        Err.Raise vbObjectError + 1001, _
                  "RunAnsichtAktualisieren", _
                  "Der benannte Bereich '" & RunName & _
                  "' muss auf genau eine Zelle zeigen."
    End If

    If rngResetSpalten Is Nothing Then
        Err.Raise vbObjectError + 1002, _
                  "RunAnsichtAktualisieren", _
                  "Der benannte Bereich '" & ResetSpaltenName & _
                  "' wurde nicht gefunden."
    End If

    If rngResetZeilen Is Nothing Then
        Err.Raise vbObjectError + 1003, _
                  "RunAnsichtAktualisieren", _
                  "Der benannte Bereich '" & ResetZeilenName & _
                  "' wurde nicht gefunden."
    End If

    If Not rngResetSpalten.Worksheet Is Me Then
        Err.Raise vbObjectError + 1004, _
                  "RunAnsichtAktualisieren", _
                  "'" & ResetSpaltenName & _
                  "' liegt nicht auf dem Blatt '" & Me.Name & "'."
    End If

    If Not rngResetZeilen.Worksheet Is Me Then
        Err.Raise vbObjectError + 1005, _
                  "RunAnsichtAktualisieren", _
                  "'" & ResetZeilenName & _
                  "' liegt nicht auf dem Blatt '" & Me.Name & "'."
    End If

    Set wsConfig = ThisWorkbook.Worksheets(ConfigBlattName)

    On Error Resume Next
    Set tblRuns = wsConfig.ListObjects(RunsTabellenName)
    On Error GoTo Fehler

    If tblRuns Is Nothing Then
        Err.Raise vbObjectError + 1006, _
                  "RunAnsichtAktualisieren", _
                  "Die Excel-Tabelle '" & RunsTabellenName & _
                  "' wurde auf dem Blatt '" & ConfigBlattName & _
                  "' nicht gefunden."
    End If

    TabellenSpalte tblRuns, "Run"
    TabellenSpalte tblRuns, "Name"
    TabellenSpalte tblRuns, "Spalten"
    TabellenSpalte tblRuns, "Zeilen"
    TabellenSpalte tblRuns, "Gruppen zu"
    TabellenSpalte tblRuns, "Gruppen auf"
    TabellenSpalte tblRuns, "Start Gruppen"

    '------------------------------------------------------
    ' Leere Steuerzelle = vollständiger Grundzustand
    '------------------------------------------------------
    If Len(Trim$(CStr(rngRun.Value2))) = 0 Then

        GrundzustandOeffnen rngResetSpalten, rngResetZeilen
        GoTo Aufraeumen

    End If

    If Not IsNumeric(rngRun.Value2) Then
        Err.Raise vbObjectError + 1007, _
                  "RunAnsichtAktualisieren", _
                  "Der Wert in '" & RunName & _
                  "' ist keine gültige Zahl."
    End If

    runWert = CLng(rngRun.Value2)

    '------------------------------------------------------
    ' Run 0 = vollständiger Grundzustand
    '------------------------------------------------------
    If runWert = 0 Then

        GrundzustandOeffnen rngResetSpalten, rngResetZeilen
        GoTo Aufraeumen

    End If

    '------------------------------------------------------
    ' Config-Zeile suchen
    '------------------------------------------------------
    Set runZeile = RunZeileSuchen(tblRuns, runWert)

    If runZeile Is Nothing Then
        Err.Raise vbObjectError + 1008, _
                  "RunAnsichtAktualisieren", _
                  "Für Run " & runWert & _
                  " wurde in '" & RunsTabellenName & _
                  "' keine Konfiguration gefunden."
    End If

    '------------------------------------------------------
    ' Config lesen
    '------------------------------------------------------
    spaltenListe = TabellenWertLesen( _
        tblRuns, runZeile, "Spalten")

    zeilenListe = TabellenWertLesen( _
        tblRuns, runZeile, "Zeilen")

    gruppenZu = TabellenWertLesen( _
        tblRuns, runZeile, "Gruppen zu")

    gruppenAuf = TabellenWertLesen( _
        tblRuns, runZeile, "Gruppen auf")

    startGruppen = UCase$(Trim$(TabellenWertLesen( _
        tblRuns, runZeile, "Start Gruppen")))

    If Len(startGruppen) = 0 Then
        startGruppen = "OPEN"
    End If

    '------------------------------------------------------
    ' Neutraler Ausgangszustand
    '------------------------------------------------------
    rngResetSpalten.EntireColumn.Hidden = False
    rngResetZeilen.EntireRow.Hidden = False

    '------------------------------------------------------
    ' Gruppierungs-Grundzustand
    '------------------------------------------------------
    Select Case startGruppen

        Case "OPEN", "OFFEN", "AUF"
            AlleGruppenOeffnen

        Case "CLOSED", "CLOSE", "GESCHLOSSEN", "ZU"
            AlleGruppenSchliessen

        Case Else
            Err.Raise vbObjectError + 1009, _
                      "RunAnsichtAktualisieren", _
                      "Ungültiger Wert bei 'Start Gruppen': " & _
                      startGruppen & vbCrLf & _
                      "Erlaubt sind OPEN oder CLOSED."

    End Select

    '------------------------------------------------------
    ' Run-Ansicht anwenden
    '------------------------------------------------------
    SpaltenAnwenden spaltenListe, rngResetSpalten
    ZeilenAnwenden zeilenListe, rngResetZeilen

    GruppenAnwenden gruppenZu, False
    GruppenAnwenden gruppenAuf, True

Aufraeumen:
    Application.Calculation = alterBerechnungsmodus
    Application.ScreenUpdating = screenUpdatingWarAktiv
    Application.EnableEvents = eventsWarenAktiv
    Exit Sub

Fehler:
    Application.Calculation = alterBerechnungsmodus
    Application.ScreenUpdating = screenUpdatingWarAktiv
    Application.EnableEvents = eventsWarenAktiv

    MsgBox "Die Run-Ansicht konnte nicht angepasst werden:" & _
           vbCrLf & Err.Description, _
           vbExclamation, _
           "Run-Konfiguration"

End Sub


'==========================================================
' Grundzustand bei run leer oder 0
'==========================================================
Private Sub GrundzustandOeffnen( _
    ByVal rngResetSpalten As Range, _
    ByVal rngResetZeilen As Range)

    rngResetSpalten.EntireColumn.Hidden = False
    rngResetZeilen.EntireRow.Hidden = False

    AlleGruppenOeffnen

End Sub


'==========================================================
' Sucht die passende Run-Zeile
'==========================================================
Private Function RunZeileSuchen( _
    ByVal tbl As ListObject, _
    ByVal runWert As Long) As ListRow

    Dim zeile As ListRow
    Dim colRun As Long
    Dim wert As Variant

    colRun = TabellenSpalte(tbl, "Run")

    For Each zeile In tbl.ListRows

        wert = zeile.Range.Cells(1, colRun).Value2

        If IsNumeric(wert) Then

            If CLng(wert) = runWert Then
                Set RunZeileSuchen = zeile
                Exit Function
            End If

        End If

    Next zeile

End Function


'==========================================================
' Liest einen Tabellenwert
'==========================================================
Private Function TabellenWertLesen( _
    ByVal tbl As ListObject, _
    ByVal zeile As ListRow, _
    ByVal spaltenName As String) As String

    Dim colIndex As Long

    colIndex = TabellenSpalte(tbl, spaltenName)

    TabellenWertLesen = Trim$(CStr( _
        zeile.Range.Cells(1, colIndex).Value2))

End Function


'==========================================================
' Liefert die relative Tabellenspalte
'==========================================================
Private Function TabellenSpalte( _
    ByVal tbl As ListObject, _
    ByVal spaltenName As String) As Long

    Dim spalte As listColumn

    On Error Resume Next
    Set spalte = tbl.ListColumns(spaltenName)
    On Error GoTo 0

    If spalte Is Nothing Then
        Err.Raise vbObjectError + 1100, _
                  "TabellenSpalte", _
                  "In der Tabelle '" & tbl.Name & _
                  "' fehlt die Spalte '" & spaltenName & "'."
    End If

    TabellenSpalte = spalte.Index

End Function


'==========================================================
' Spalten anwenden
'==========================================================
Private Sub SpaltenAnwenden( _
    ByVal liste As String, _
    ByVal rngResetSpalten As Range)

    Dim elemente As Variant
    Dim element As Variant
    Dim bereichsName As String
    Dim rng As Range

    If Len(Trim$(liste)) = 0 Then Exit Sub

    rngResetSpalten.EntireColumn.Hidden = True

    liste = NormalisiereNamenListe(liste)
    elemente = Split(liste, ";")

    For Each element In elemente

        bereichsName = Trim$(CStr(element))

        If Len(bereichsName) > 0 Then

            Set rng = GetNamedRange(Me, bereichsName)

            If rng Is Nothing Then
                Err.Raise vbObjectError + 1110, _
                          "SpaltenAnwenden", _
                          "Der benannte Spaltenbereich '" & _
                          bereichsName & "' wurde nicht gefunden."
            End If

            If Not rng.Worksheet Is Me Then
                Err.Raise vbObjectError + 1111, _
                          "SpaltenAnwenden", _
                          "Der benannte Bereich '" & bereichsName & _
                          "' liegt nicht auf dem Blatt '" & Me.Name & "'."
            End If

            rng.EntireColumn.Hidden = False

        End If

    Next element

End Sub


'==========================================================
' Zeilen anwenden
'==========================================================
Private Sub ZeilenAnwenden( _
    ByVal liste As String, _
    ByVal rngResetZeilen As Range)

    Dim elemente As Variant
    Dim element As Variant
    Dim wert As String
    Dim zeilenNummer As Long

    If Len(Trim$(liste)) = 0 Then Exit Sub

    rngResetZeilen.EntireRow.Hidden = True

    liste = NormalisiereZahlenListe(liste)
    elemente = Split(liste, ";")

    For Each element In elemente

        wert = Trim$(CStr(element))

        If Len(wert) > 0 Then

            If Not IsNumeric(wert) Then
                Err.Raise vbObjectError + 1120, _
                          "ZeilenAnwenden", _
                          "Ungültige Zeilenangabe: " & wert
            End If

            zeilenNummer = CLng(wert)

            If zeilenNummer < 1 Or _
               zeilenNummer > Me.Rows.Count Then

                Err.Raise vbObjectError + 1121, _
                          "ZeilenAnwenden", _
                          "Ungültige Zeilennummer: " & zeilenNummer
            End If

            Me.Rows(zeilenNummer).Hidden = False

        End If

    Next element

End Sub


'==========================================================
' Gruppen öffnen oder schließen
'==========================================================
Private Sub GruppenAnwenden( _
    ByVal liste As String, _
    ByVal oeffnen As Boolean)

    Dim elemente As Variant
    Dim element As Variant
    Dim wert As String

    Dim ersteDetailZeile As Long
    Dim zusammenfassungsZeile As Long

    If Len(Trim$(liste)) = 0 Then Exit Sub

    liste = NormalisiereZahlenListe(liste)
    elemente = Split(liste, ";")

    For Each element In elemente

        wert = Trim$(CStr(element))

        If Len(wert) > 0 Then

            If Not IsNumeric(wert) Then
                Err.Raise vbObjectError + 1130, _
                          "GruppenAnwenden", _
                          "Ungültige Gruppenangabe: " & wert
            End If

            ersteDetailZeile = CLng(wert)

            zusammenfassungsZeile = _
                GruppenZusammenfassungsZeile(ersteDetailZeile)

            If zusammenfassungsZeile = 0 Then
                Err.Raise vbObjectError + 1131, _
                          "GruppenAnwenden", _
                          "Bei Zeile " & ersteDetailZeile & _
                          " wurde keine Gruppierung gefunden."
            End If

            On Error GoTo GruppenFehler

            Me.Rows(zusammenfassungsZeile).ShowDetail = oeffnen

            On Error GoTo 0

        End If

    Next element

    Exit Sub

GruppenFehler:
    Err.Raise vbObjectError + 1132, _
              "GruppenAnwenden", _
              "Die Gruppierung ab Zeile " & _
              ersteDetailZeile & _
              " konnte nicht geändert werden."

End Sub


'==========================================================
' Ermittelt die Zusammenfassungszeile
'==========================================================
Private Function GruppenZusammenfassungsZeile( _
    ByVal ersteDetailZeile As Long) As Long

    Dim gruppenEbene As Long
    Dim aktuelleZeile As Long

    If ersteDetailZeile < 1 Or _
       ersteDetailZeile > Me.Rows.Count Then
        Exit Function
    End If

    gruppenEbene = Me.Rows(ersteDetailZeile).OutlineLevel

    If gruppenEbene <= 1 Then Exit Function

    If Me.Outline.SummaryRow = xlBelow Then

        aktuelleZeile = ersteDetailZeile

        Do While aktuelleZeile < Me.Rows.Count

            aktuelleZeile = aktuelleZeile + 1

            If Me.Rows(aktuelleZeile).OutlineLevel < gruppenEbene Then
                GruppenZusammenfassungsZeile = aktuelleZeile
                Exit Function
            End If

        Loop

    Else

        aktuelleZeile = ersteDetailZeile

        Do While aktuelleZeile > 1

            aktuelleZeile = aktuelleZeile - 1

            If Me.Rows(aktuelleZeile).OutlineLevel < gruppenEbene Then
                GruppenZusammenfassungsZeile = aktuelleZeile
                Exit Function
            End If

        Loop

    End If

End Function


'==========================================================
' Alle Gruppen öffnen
'==========================================================
Private Sub AlleGruppenOeffnen()

    On Error GoTo Fehler

    Me.Outline.ShowLevels RowLevels:=2
    Exit Sub

Fehler:
    Err.Raise vbObjectError + 1140, _
              "AlleGruppenOeffnen", _
              "Die Gruppierungen konnten nicht geöffnet werden."

End Sub


'==========================================================
' Alle Gruppen schließen
'==========================================================
Private Sub AlleGruppenSchliessen()

    On Error GoTo Fehler

    Me.Outline.ShowLevels RowLevels:=1
    Exit Sub

Fehler:
    Err.Raise vbObjectError + 1141, _
              "AlleGruppenSchliessen", _
              "Die Gruppierungen konnten nicht geschlossen werden."

End Sub


'==========================================================
' Normalisiert eine Liste von Namen
'==========================================================
Private Function NormalisiereNamenListe( _
    ByVal wert As String) As String

    wert = Trim$(wert)
    wert = Replace(wert, ",", ";")
    wert = Replace(wert, "|", ";")

    NormalisiereNamenListe = ListeBereinigen(wert)

End Function


'==========================================================
' Normalisiert eine Zahlenliste
'==========================================================
Private Function NormalisiereZahlenListe( _
    ByVal wert As String) As String

    wert = Trim$(wert)
    wert = Replace(wert, ",", ";")
    wert = Replace(wert, ".", ";")
    wert = Replace(wert, "|", ";")

    NormalisiereZahlenListe = ListeBereinigen(wert)

End Function


'==========================================================
' Bereinigt Trennzeichen
'==========================================================
Private Function ListeBereinigen( _
    ByVal wert As String) As String

    Do While InStr(wert, ";;") > 0
        wert = Replace(wert, ";;", ";")
    Loop

    Do While Len(wert) > 0 And Left$(wert, 1) = ";"
        wert = Mid$(wert, 2)
    Loop

    Do While Len(wert) > 0 And Right$(wert, 1) = ";"
        wert = Left$(wert, Len(wert) - 1)
    Loop

    ListeBereinigen = wert

End Function


'==========================================================
' Findet einen benannten Bereich
'==========================================================
Private Function GetNamedRange( _
    ByVal ws As Worksheet, _
    ByVal nameString As String) As Range

    Dim nm As Name

    Set GetNamedRange = Nothing

    nameString = Trim$(nameString)

    If Len(nameString) = 0 Then Exit Function

    On Error Resume Next
    Set GetNamedRange = ws.Range(nameString)
    On Error GoTo 0

    If Not GetNamedRange Is Nothing Then Exit Function

    On Error Resume Next
    Set nm = ThisWorkbook.Names(nameString)
    On Error GoTo 0

    If Not nm Is Nothing Then

        On Error Resume Next
        Set GetNamedRange = nm.RefersToRange
        On Error GoTo 0

    End If

End Function

