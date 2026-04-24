Attribute VB_Name = "genToggleButton"
Public Sub ToggleButtonErstellen()

    Dim ws As Worksheet
    Dim toggleObj As OLEObject
    Set ws = ThisWorkbook.Worksheets("Grafik")

    ' Entferne vorhandene Schalter (falls nötig)
    On Error Resume Next
    ws.OLEObjects("SwitchButton").Delete
    On Error GoTo 0

    ' Toggle-Button hinzufügen
    Set toggleObj = ws.OLEObjects.Add(ClassType:="Forms.ToggleButton.1", _
                                      Left:=700, Top:=10, Width:=120, Height:=30)
    toggleObj.name = "SwitchButton"
    toggleObj.Object.Caption = "Rechtwinklig"
    toggleObj.Object.Value = False  ' Standardwert: Polar deaktiviert
End Sub


