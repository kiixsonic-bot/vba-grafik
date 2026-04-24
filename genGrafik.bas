Attribute VB_Name = "genGrafik"
'Attribute VB_Name = "modKoordGrafik"
Option Explicit

Public KoordButtons As Collection

' ============================================================
'   Aus den beiden Arrays die ActiveX-Buttons erzeugen
' ============================================================
Public Sub ErzeugeKoordButtons(ByRef arrAlt() As clsReKoord, _
                               ByRef arrNeu() As clsReKoord)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Grafik")

    ' --- alte Buttons entfernen ---
    Dim obj As OLEObject
    For Each obj In ws.OLEObjects
        obj.Delete
    Next obj

    Set KoordButtons = New Collection

    ' --- Min/Max ermitteln ---
    Dim yMin As Double, yMax As Double
    Dim xMin As Double, xMax As Double
    Call GetMinMax(arrAlt, arrNeu, yMin, yMax, xMin, xMax)

    ' --- Skalierung ---
    Dim margin As Double: margin = 40
    Dim plotW  As Double: plotW = 600
    Dim plotH  As Double: plotH = 500

    If yMax - yMin = 0 Then yMax = yMin + 1
    If xMax - xMin = 0 Then xMax = xMin + 1

    Dim s As Double
    s = WorksheetFunction.Min(plotW / (yMax - yMin), _
                              plotH / (xMax - xMin))

    ' --- Alt-Punkte (rot) ---
    Dim i As Long
    For i = LBound(arrAlt) To UBound(arrAlt)
        If Not arrAlt(i) Is Nothing Then
            Call AddKoordButton(ws, arrAlt(i), s, yMin, xMax, margin, _
                                RGB(220, 50, 50), "_alt")
        End If
    Next i

    ' --- Neu-Punkte (blau) ---
    For i = LBound(arrNeu) To UBound(arrNeu)
        If Not arrNeu(i) Is Nothing Then
            Call AddKoordButton(ws, arrNeu(i), s, yMin, xMax, margin, _
                                RGB(0, 120, 255), "_neu")
        End If
    Next i

    MsgBox KoordButtons.Count & " Koordinaten-Buttons erzeugt.", vbInformation

End Sub

' ============================================================
'   Einen clsKoordButton aus einem clsReKoord erzeugen
' ============================================================
Private Sub AddKoordButton(ws As Worksheet, _
                           pt As clsReKoord, _
                           ByVal s As Double, _
                           ByVal yMin As Double, _
                           ByVal xMax As Double, _
                           ByVal margin As Double, _
                           ByVal btnColor As Long, _
                           ByVal suffix As String)

    ' Geodätisch: Y ? Left (horizontal), X ? Top (vertikal, invertiert)
    Dim leftPt As Double
    Dim topPt  As Double
    leftPt = margin + (pt.y - yMin) * s
    topPt = margin + (xMax - pt.x) * s

    Dim btnObj As OLEObject
    Set btnObj = ws.OLEObjects.Add( _
        ClassType:="Forms.CommandButton.1", _
        Left:=leftPt, Top:=topPt, _
        Width:=14, Height:=14)

    btnObj.Object.Caption = ""
    btnObj.Object.BackColor = btnColor
    btnObj.Object.TakeFocusOnClick = False
    btnObj.name = "btn_" & Replace(pt.name, " ", "_") & suffix

    ' clsKoordButton-Instanz
    Dim c As clsKoordButton
    Set c = New clsKoordButton
    Set c.Btn = btnObj.Object
    c.name = pt.name
    c.x = pt.x
    c.y = pt.y

    KoordButtons.Add c

End Sub

' ============================================================
'   Min/Max über beide Arrays (Nothing-Einträge überspringen)
' ============================================================
Private Sub GetMinMax(ByRef arrA() As clsReKoord, _
                      ByRef arrN() As clsReKoord, _
                      ByRef yMin As Double, ByRef yMax As Double, _
                      ByRef xMin As Double, ByRef xMax As Double)

    Dim first As Boolean: first = True
    Dim i As Long

    For i = LBound(arrA) To UBound(arrA)
        If Not arrA(i) Is Nothing Then
            If first Then
                yMin = arrA(i).y: yMax = arrA(i).y
                xMin = arrA(i).x: xMax = arrA(i).x
                first = False
            Else
                If arrA(i).y < yMin Then yMin = arrA(i).y
                If arrA(i).y > yMax Then yMax = arrA(i).y
                If arrA(i).x < xMin Then xMin = arrA(i).x
                If arrA(i).x > xMax Then xMax = arrA(i).x
            End If
        End If
    Next i

    For i = LBound(arrN) To UBound(arrN)
        If Not arrN(i) Is Nothing Then
            If first Then
                yMin = arrN(i).y: yMax = arrN(i).y
                xMin = arrN(i).x: xMax = arrN(i).x
                first = False
            Else
                If arrN(i).y < yMin Then yMin = arrN(i).y
                If arrN(i).y > yMax Then yMax = arrN(i).y
                If arrN(i).x < xMin Then xMin = arrN(i).x
                If arrN(i).x > xMax Then xMax = arrN(i).x
            End If
        End If
    Next i

End Sub

