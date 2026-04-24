VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKoordEingabe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ActiveXButton As OLEObject
Attribute ActiveXButton.VB_VarHelpID = -1
Private WithEvents EingabeYorDist As OLEObject
Attribute EingabeYorDist.VB_VarHelpID = -1
Private WithEvents EingabeXorDir As OLEObject
Attribute EingabeXorDir.VB_VarHelpID = -1

' Initialisiert die ActiveX-Steuerelemente
Public Sub Init( _
    ByRef toggleCtrl As OLEObject, _
    ByRef txtYorDist As OLEObject, _
    ByRef txtXorDir As OLEObject)

    Set ActiveXButton = toggleCtrl
    Set EingabeYorDist = txtYorDist
    Set EingabeXorDir = txtXorDir

    ' Standardbeschriftung setzen
    UpdateBeschriftung False
End Sub

' Ereignishandler für den Toggle-Button (Click-Event)
Private Sub ActiveXButton_Click()

    Dim IsPolar As Boolean
    IsPolar = ActiveXButton.Object.Value

    ' Beschriftungen aktualisieren
    UpdateBeschriftung IsPolar
End Sub

' Beschriftungen für die Textfelder und den Button ändern
Private Sub UpdateBeschriftung(ByVal IsPolar As Boolean)

    If IsPolar Then
        ActiveXButton.Object.Caption = "Polar"
        EingabeYorDist.Object.Text = ""
        EingabeYorDist.Object.Value = "Distanz"
        EingabeXorDir.Object.Value = "Richtung"
    Else
        ActiveXButton.Object.Caption = "Rechtwinklig"
        EingabeYorDist.Object.Text = ""
        EingabeYorDist.Object.Value = "Y-Wert"
        EingabeXorDir.Object.Value = "X-Wert"
    End If

End Sub


