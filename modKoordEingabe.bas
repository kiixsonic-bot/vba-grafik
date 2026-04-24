Attribute VB_Name = "modKoordEingabe"
Public Sub KoordinatenHinzufügen()
    Dim ws As Worksheet
    Dim KoordNeu As clsReKoord
    Dim StartKoord As clsReKoord
    Dim tmpY As Double, tmpX As Double
    Dim Distanz As Double, Richtung As Double

    Set ws = ThisWorkbook.Worksheets("Grafik")

    ' Toggle-Zustand prüfen
    Dim IsPolar As Boolean
    IsPolar = ws.OLEObjects("SwitchButton").Object.Value

    ' Neuen Punkt erstellen
    If IsPolar = False Then
        ' --- Rechtwinklig ---
        Set KoordNeu = New clsReKoord
        KoordNeu.name = "Custom_" & Timer
        KoordNeu.y = CDbl(ws.OLEObjects("EingabeYorDist").Object.Text)
        KoordNeu.x = CDbl(ws.OLEObjects("EingabeXorDir").Object.Text)
    Else
        ' --- Polar: Basis-Koordinate wählen ---
        If SelectedKoordinate Is Nothing Then
            Call SetInfoText("Wähle eine bestehende Koordinate aus!")
            Exit Sub
        End If

        Set StartKoord = SelectedKoordinate
        Distanz = CDbl(ws.OLEObjects("EingabeYorDist").Object.Text)
        Richtung = CDbl(ws.OLEObjects("EingabeXorDir").Object.Text)

        ' Polarkoordinate umrechnen
        tmpY = StartKoord.y + Distanz * Cos(Richtung * WorksheetFunction.Pi() / 180)
        tmpX = StartKoord.x + Distanz * Sin(Richtung * WorksheetFunction.Pi() / 180)

        Set KoordNeu = New clsReKoord
        KoordNeu.name = "Custom_" & Timer
        KoordNeu.y = tmpY
        KoordNeu.x = tmpX
    End If

    ' Zur Grafik und Collection hinzufügen
    Call AddKoordButton(ws, KoordNeu, RGB(0, 150, 0), "Custom")
    Call AddInfoText("Neue Koordinate hinzugefügt: " & KoordNeu.name)
End Sub
