Attribute VB_Name = "genEingabefenster"
Public Sub EingabebereichErstellen()

    Dim ws As Worksheet
    Dim obj As OLEObject
    Dim clsKoordCtrl As clsKoordEingabe
    Set ws = ThisWorkbook.Worksheets("Grafik")

    ' Entferne alte Controls, falls vorhanden
    On Error Resume Next
    ws.OLEObjects("EingabeYorDist").Delete
    ws.OLEObjects("EingabeXorDir").Delete
    On Error GoTo 0

    ' --- TextBox: Y oder Distanz ---
    Set obj = ws.OLEObjects.Add(ClassType:="Forms.TextBox.1", _
                                Left:=700, Top:=50, Width:=120, Height:=25)
    obj.name = "EingabeYorDist"
    obj.Object.Text = "Y-Wert"  ' Startwert

    ' --- TextBox: X oder Richtung ---
    Set obj = ws.OLEObjects.Add(ClassType:="Forms.TextBox.1", _
                                Left:=700, Top:=85, Width:=120, Height:=25)
    obj.name = "EingabeXorDir"
    obj.Object.Text = "X-Wert"  ' Startwert

    ' --- Toggle-Button erstellen ---
    Call ToggleButtonErstellen

    ' --- Klassensteuerung mit Events ---
    Set clsKoordCtrl = New clsKoordEingabe
    clsKoordCtrl.Init ws.OLEObjects("SwitchButton").Object, _
                      ws.OLEObjects("EingabeYorDist").Object, _
                      ws.OLEObjects("EingabeXorDir").Object

End Sub

