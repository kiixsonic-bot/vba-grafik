VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReKoord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pName As String
Private pY As Variant
Private pX As Variant

Public Property Let name(ByVal v As String)
    pName = v
End Property

Public Property Get name() As String
    name = pName
End Property

Public Property Let y(ByVal v As Variant)
    pY = v
End Property

Public Property Get y() As Variant
    y = pY
End Property

Public Property Let x(ByVal v As Variant)
    pX = v
End Property

Public Property Get x() As Variant
    x = pX
End Property

