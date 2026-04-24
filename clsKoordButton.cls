VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKoordButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Btn As MSForms.CommandButton
Attribute Btn.VB_VarHelpID = -1

Private pName As String
Private pX As Double
Private pY As Double

' Eigenschaften setzen
Public Property Let name(ByVal v As String)
    pName = v
End Property

Public Property Get name() As String
    name = pName
End Property

Public Property Let x(ByVal v As Double)
    pX = v
End Property

Public Property Get x() As Double
    x = pX
End Property

Public Property Let y(ByVal v As Double)
    pY = v
End Property

Public Property Get y() As Double
    y = pY
End Property

' Button-Klick-Ereignis



Private Sub Btn_Click()
    ' Koordinate auswählen, falls Polar aktiviert ist
    Dim ws As Worksheet
    Dim IsPolar As Boolean
    Set ws = ThisWorkbook.Worksheets("Grafik")
    IsPolar = ws.OLEObjects("SwitchButton").Object.Value

    If IsPolar Then
        Set SelectedKoordinate = Me
        Call AddInfoText("Gewählte Startkoordinate: " & Me.name)
    Else
        MsgBox "Punkt: " & Me.name & vbCrLf & _
               "X = " & Me.x & vbCrLf & _
               "Y = " & Me.y, vbInformation, "Koordinateninfo"
    End If
End Sub
