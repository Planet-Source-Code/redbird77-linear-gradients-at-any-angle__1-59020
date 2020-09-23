Attribute VB_Name = "mGradient"
' mGradient.bas
' 2005 February 19
' redbird77@earthlink.net
' http://home.earthlink.net/~redbird77

' To render a gradient all you need is either:
'
' the DrawGradient sub and the 6 included API declares - OR -
' the DrawGradientVB sub and nothing else.
'
' Neither must be in a module, they can go anywhere - form, class, module...

Option Explicit

Private Declare Sub RtlMoveMemory Lib "kernel32" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function MoveTo Lib "gdi32" Alias "MoveToEx" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Sub DrawGradient(ByVal hDC As Long, _
                        ByVal lWidth As Long, ByVal lHeight As Long, _
                        ByVal lCol1 As Long, ByVal lCol2 As Long, _
                        ByVal zAngle As Single)

Dim xStart  As Long, yStart As Long
Dim xEnd    As Long, yEnd   As Long
Dim x1      As Long, y1     As Long
Dim x2      As Long, y2     As Long
Dim lRange  As Long
Dim iQ      As Integer
Dim bVert   As Boolean
Dim lPtr    As Long, lInc   As Long
Dim lCols() As Long
Dim hPO     As Long, hPN    As Long
Dim r       As Long
Dim x       As Long, xUp    As Long
Dim b1(2)   As Byte, b2(2)  As Byte
Dim p       As Single, ip   As Single

    lInc = 1: xEnd = lWidth - 1: yEnd = lHeight - 1
    
    ' Positive angles are measured counter-clockwise; negative angles clockwise.
    zAngle = zAngle Mod 360
    If zAngle < 0 Then zAngle = 360 + zAngle
    
    ' Get angle's quadrant (0 - 3).
    iQ = zAngle \ 90
    
    ' Is angle more horizontal or vertical?
    bVert = ((iQ + 1) * 90) - zAngle > 45
    If (iQ Mod 2 = 0) Then bVert = Not bVert
    
    ' Convert angle in degrees to radians.
    zAngle = zAngle * Atn(1) / 45
    
    ' Get start and end y-positions (if vertical), x-positions (if horizontal).
    If bVert Then
        If zAngle Then xStart = lHeight / Abs(Tan(zAngle))
        lRange = lWidth + xStart - 1

        y1 = IIf(iQ Mod 2, 0, yEnd)
        y2 = IIf(y1, -1, lHeight)
        
        If iQ > 1 Then
            lPtr = lRange: lInc = -1
        End If
    Else
        yStart = lWidth * Abs(Tan(zAngle))
        lRange = lHeight + yStart - 1

        x1 = IIf(iQ Mod 2, 0, xEnd)
        x2 = IIf(x1, -1, lWidth)
        
        If iQ = 1 Or iQ = 2 Then
            lPtr = lRange: lInc = -1
        End If
    End If
    
' -------------------------------------------------------------------
' Fill in the color array with the interpolated color values.
' -------------------------------------------------------------------
    ReDim lCols(lRange)

    ' Get the r, g, b components of each color.
    RtlMoveMemory b1(0), lCol1, 3
    RtlMoveMemory b2(0), lCol2, 3

    xUp = UBound(lCols)
    
    For x = 0 To xUp
        ' Get the position and the 1 - position.
        p = x / xUp
        ip = 1 - p
        
        ' Interpolate the value at the current position.
        lCols(x) = RGB(b1(0) * ip + b2(0) * p, b1(1) * ip + b2(1) * p, b1(2) * ip + b2(2) * p)
    Next
   
' -------------------------------------------------------------------
' Draw the lines of the gradient at user-specified angle.
' -------------------------------------------------------------------
    If bVert Then
        For x1 = -xStart To xEnd
            hPN = CreatePen(0, 1, lCols(lPtr))
            hPO = SelectObject(hDC, hPN)
            MoveTo hDC, x1, y1, ByVal 0&
            LineTo hDC, x2, y2
            r = SelectObject(hDC, hPO): r = DeleteObject(hPN)
            lPtr = lPtr + lInc
            x2 = x2 + 1
        Next
    Else
        For y1 = -yStart To yEnd
            hPN = CreatePen(0, 1, lCols(lPtr))
            hPO = SelectObject(hDC, hPN)
            MoveTo hDC, x1, y1, ByVal 0&
            LineTo hDC, x2, y2
            r = SelectObject(hDC, hPO): r = DeleteObject(hPN)
            lPtr = lPtr + lInc
            y2 = y2 + 1
        Next
    End If
    
End Sub

Public Sub DrawGradientVB(ByRef oCanvas As Object, _
                          ByVal lCol1 As Long, ByVal lCol2 As Long, _
                          ByVal zAngle As Single)
                          
' If you are going to draw the gradient on a form or picture box, use
' this version instead because you needn't specify the height and
' width.

Dim xStart  As Long, yStart As Long
Dim xEnd    As Long, yEnd   As Long
Dim x1      As Long, y1     As Long
Dim x2      As Long, y2     As Long
Dim lRange  As Long
Dim iQ      As Integer
Dim bVert   As Boolean
Dim lPtr    As Long, lInc   As Long
Dim lCols() As Long
Dim hPO     As Long, hPN    As Long
Dim r       As Long, hDC    As Long
Dim x       As Long, xUp    As Long
Dim b1(2)   As Byte, b2(2)  As Byte
Dim p       As Single, ip   As Single
Dim lWid    As Long, lHgt   As Long

    lWid = oCanvas.ScaleWidth: lHgt = oCanvas.ScaleHeight
    lInc = 1: xEnd = lWid - 1: yEnd = lHgt - 1
    
    ' Positive angles are measured counter-clockwise; negative angles clockwise.
    zAngle = zAngle Mod 360
    If zAngle < 0 Then zAngle = 360 + zAngle
    
    ' Get angle's quadrant (0 - 3).
    iQ = zAngle \ 90
    
    ' Is angle more horizontal or vertical?
    bVert = ((iQ + 1) * 90) - zAngle > 45
    If (iQ Mod 2 = 0) Then bVert = Not bVert
    
    ' Convert angle in degrees to radians.
    zAngle = zAngle * Atn(1) / 45
    
    ' Get start and end y-positions (if vertical), x-positions (if horizontal).
    If bVert Then
        If zAngle Then xStart = lHgt / Abs(Tan(zAngle))
        lRange = lWid + xStart - 1

        y1 = IIf(iQ Mod 2, 0, yEnd)
        y2 = IIf(y1, -1, lHgt)
        
        If iQ > 1 Then
            lPtr = lRange: lInc = -1
        End If
    Else
        yStart = lWid * Abs(Tan(zAngle))
        lRange = lHgt + yStart - 1

        x1 = IIf(iQ Mod 2, 0, xEnd)
        x2 = IIf(x1, -1, lWid)
        
        If iQ = 1 Or iQ = 2 Then
            lPtr = lRange: lInc = -1
        End If
    End If
    
' -------------------------------------------------------------------
' Fill in the color array with the interpolated color values.
' -------------------------------------------------------------------
    ReDim lCols(lRange)

    ' Get the r, g, b components of each color.
    RtlMoveMemory b1(0), lCol1, 3
    RtlMoveMemory b2(0), lCol2, 3

    xUp = UBound(lCols)
    
    For x = 0 To xUp
        ' Get the position and the 1 - position.
        p = x / xUp
        ip = 1 - p
        
        ' Interpolate the value at the current position.
        lCols(x) = RGB(b1(0) * ip + b2(0) * p, b1(1) * ip + b2(1) * p, b1(2) * ip + b2(2) * p)
    Next
   
' -------------------------------------------------------------------
' Draw the lines of the gradient at user-specified angle.
' -------------------------------------------------------------------
    hDC = oCanvas.hDC
    
    If bVert Then
        For x1 = -xStart To xEnd
            hPN = CreatePen(0, 1, lCols(lPtr))
            hPO = SelectObject(hDC, hPN)
            MoveTo hDC, x1, y1, ByVal 0&
            LineTo hDC, x2, y2
            r = SelectObject(hDC, hPO): r = DeleteObject(hPN)
            lPtr = lPtr + lInc
            x2 = x2 + 1
        Next
    Else
        For y1 = -yStart To yEnd
            hPN = CreatePen(0, 1, lCols(lPtr))
            hPO = SelectObject(hDC, hPN)
            MoveTo hDC, x1, y1, ByVal 0&
            LineTo hDC, x2, y2
            r = SelectObject(hDC, hPO): r = DeleteObject(hPN)
            lPtr = lPtr + lInc
            y2 = y2 + 1
        Next
    End If
    
End Sub
