VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPolarkoordinaten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pName As String
Private pStrecke As Variant
Private pRichtung As Variant

Public Property Let name(ByVal v As String)
    pName = v
End Property

Public Property Get name() As String
    name = pName
End Property

Public Property Let Strecke(ByVal v As Variant)
    pStrecke = v
End Property

Public Property Get Strecke() As Variant
    Strecke = pStrecke
End Property

Public Property Let Richtung(ByVal v As Variant)
    pRichtung = v
End Property

Public Property Get Richtung() As Variant
    Richtung = pRichtung
End Property

