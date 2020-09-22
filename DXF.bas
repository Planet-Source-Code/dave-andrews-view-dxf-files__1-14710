Attribute VB_Name = "DXF"
Option Explicit
Const pi = 3.14159265358979

Type RECT
    X1 As Single
    Y1 As Single
    X2 As Single
    Y2 As Single
End Type

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
' lfFaceName(LF_FACESIZE)
  lfFacename As String * 33
End Type

Type DataSet
    Key As Integer
    Value As Variant
End Type

Type Geometry
    Type As String
    Data() As DataSet
End Type

Type Block
    Name As String
    Entities() As Geometry
End Type

Type DXFData
    Blocks() As Block
    Entities() As Geometry
End Type

Dim Section() As String

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Sub ClearKeys(ByRef Geo As Geometry)
Dim i As Integer
For i = 0 To UBound(Geo.Data)
    Geo.Data(i).Key = i
Next i
End Sub

Function dAngle(Angle As Single) As Single
If Angle > 360 Then
    dAngle = Angle - 360
ElseIf Angle < 0 Then
    dAngle = Angle + 360
Else
    dAngle = Angle
End If
End Function


Function FindStart(sArray() As String, Start As Long)
Dim i As Long
For i = Start To UBound(sArray)
    If sArray(i) = "10" And sArray(i + 2) = "20" Then
        FindStart = i
        Exit Function
    End If
Next i
FindStart = -1
End Function

Sub PrepareEntity(ByRef Geo As Geometry)
'This may take a little more time during the "load"
'and it may take a little more memory, but in the end
'it will draw much faster
Dim i As Integer
Select Case Geo.Type
    Case "LINE"
        If UBound(Geo.Data) < 3 Then ReDim Preserve Geo.Data(3) As DataSet
        Geo.Data(0).Value = kVal(Geo.Data(), 10)
        Geo.Data(1).Value = kVal(Geo.Data(), 20)
        Geo.Data(2).Value = kVal(Geo.Data(), 11)
        Geo.Data(3).Value = kVal(Geo.Data(), 21)
        ReDim Preserve Geo.Data(3) As DataSet
    Case "ARC"
        If UBound(Geo.Data) < 4 Then ReDim Preserve Geo.Data(4) As DataSet
        Geo.Data(0).Value = kVal(Geo.Data(), 10)
        Geo.Data(1).Value = kVal(Geo.Data(), 20)
        Geo.Data(2).Value = kVal(Geo.Data(), 40)
        Geo.Data(3).Value = kVal(Geo.Data(), 50)
        Geo.Data(4).Value = kVal(Geo.Data(), 51)
        ReDim Preserve Geo.Data(4) As DataSet
    Case "CIRCLE"
        If UBound(Geo.Data) < 2 Then ReDim Preserve Geo.Data(2) As DataSet
        Geo.Data(0).Value = kVal(Geo.Data(), 10)
        Geo.Data(1).Value = kVal(Geo.Data(), 20)
        Geo.Data(2).Value = kVal(Geo.Data(), 40)
        ReDim Preserve Geo.Data(2) As DataSet
    Case "ELLIPSE"
        If UBound(Geo.Data) < 6 Then ReDim Preserve Geo.Data(6) As DataSet
        Geo.Data(0).Value = kVal(Geo.Data(), 10)
        Geo.Data(1).Value = kVal(Geo.Data(), 20)
        Geo.Data(2).Value = kVal(Geo.Data(), 11)
        Geo.Data(3).Value = kVal(Geo.Data(), 21)
        Geo.Data(4).Value = kVal(Geo.Data(), 40)
        Geo.Data(5).Value = kVal(Geo.Data(), 41)
        Geo.Data(6).Value = kVal(Geo.Data(), 42)
        ReDim Preserve Geo.Data(6) As DataSet
    Case "VERTEX"
        If UBound(Geo.Data) < 1 Then ReDim Preserve Geo.Data(1) As DataSet
        Geo.Data(0).Value = kVal(Geo.Data(), 10)
        Geo.Data(1).Value = kVal(Geo.Data(), 20)
        ReDim Preserve Geo.Data(1) As DataSet
    Case "TEXT"
        If UBound(Geo.Data) < 4 Then ReDim Preserve Geo.Data(4) As DataSet
        Geo.Data(0).Value = kVal(Geo.Data(), 10)
        Geo.Data(1).Value = kVal(Geo.Data(), 20)
        Geo.Data(2).Value = kVal(Geo.Data(), 40)
        Geo.Data(3).Value = kVal(Geo.Data(), 50)
        Geo.Data(4).Value = kVal(Geo.Data(), 1)
        ReDim Preserve Geo.Data(4) As DataSet
    Case "INSERT"
        If UBound(Geo.Data) < 5 Then ReDim Preserve Geo.Data(5) As DataSet
        Geo.Data(0).Value = kVal(Geo.Data(), 2)
        Geo.Data(1).Value = kVal(Geo.Data(), 10)
        Geo.Data(2).Value = kVal(Geo.Data(), 20)
        Geo.Data(3).Value = kVal(Geo.Data(), 41)
        Geo.Data(4).Value = kVal(Geo.Data(), 42)
        Geo.Data(5).Value = kVal(Geo.Data(), 50)
        ReDim Preserve Geo.Data(5) As DataSet
    Case "DIMENSION"
        If UBound(Geo.Data) < 10 Then ReDim Preserve Geo.Data(10) As DataSet
        Geo.Data(0).Value = kVal(Geo.Data(), 2)
        Geo.Data(1).Value = kVal(Geo.Data(), 10)
        Geo.Data(2).Value = kVal(Geo.Data(), 20)
        Geo.Data(3).Value = kVal(Geo.Data(), 11)
        Geo.Data(4).Value = kVal(Geo.Data(), 21)
        Geo.Data(5).Value = kVal(Geo.Data(), 12)
        Geo.Data(6).Value = kVal(Geo.Data(), 22)
        Geo.Data(7).Value = kVal(Geo.Data(), 13)
        Geo.Data(8).Value = kVal(Geo.Data(), 23)
        Geo.Data(9).Value = kVal(Geo.Data(), 14)
        Geo.Data(10).Value = kVal(Geo.Data(), 24)
        ReDim Preserve Geo.Data(10) As DataSet
End Select
ClearKeys Geo
End Sub
Function PtAng(X1 As Single, Y1 As Single) As Single
If X1 = 0 Then
    If Y1 >= 0 Then
        PtAng = 90
    Else
        PtAng = 270
    End If
    PtAng = PtAng * pi / 180
    Exit Function
ElseIf Y1 = 0 Then
    If X1 >= 0 Then
        PtAng = 0
    Else
        PtAng = 180
    End If
    PtAng = PtAng * pi / 180
    Exit Function
Else
    PtAng = Atn(Y1 / X1)
    PtAng = PtAng * 180 / pi
    If PtAng < 0 Then PtAng = PtAng + 360
    If PtAng > 360 Then PtAng = PtAng - 360
    '----------Test for direction-(quadrant check)-------
    If X1 < 0 Then PtAng = PtAng + 180
    If Y1 < 0 And PtAng < 90 Then PtAng = PtAng + 180
    'If X1 < 0 And PtAng <> 180 Then PtAng = PtAng + 180
    'If Y1 < 0 And PtAng = 90 Then PtAng = PtAng + 180
    
    'One final check
    If PtAng < 0 Then PtAng = PtAng + 360
    If PtAng > 360 Then PtAng = PtAng - 360
    PtAng = PtAng * pi / 180
End If
End Function
Function cHyp(X1 As Single, Y1 As Single) As Single
cHyp = Sqr((X1 * X1) + (Y1 * Y1))
End Function

Sub DrawDXF(Canvas As PictureBox, DXF As DXFData)
On Error GoTo exitMe
Canvas.Cls
Canvas.Picture = LoadPicture()
Dim i As Integer
For i = 0 To UBound(DXF.Entities)
    DrawDXFGeometry Canvas, DXF, DXF.Entities(), i, 0, 0, 1, 1, 0
Next i
Canvas.Picture = Canvas.Image
exitMe:
End Sub

Sub DrawBlock(Canvas As PictureBox, DXF As DXFData, BlockNum As Integer)
On Error GoTo exitMe
Canvas.Cls
Canvas.Picture = LoadPicture()
Dim i As Integer
For i = 0 To UBound(DXF.Blocks(BlockNum).Entities)
    DrawDXFGeometry Canvas, DXF, DXF.Blocks(BlockNum).Entities(), i, 0, 0, 1, 1, 0
Next i
Canvas.Picture = Canvas.Image
exitMe:
End Sub
Sub DrawDXFBlock(Canvas As PictureBox, DXF As DXFData, Name As String, cX As Single, cY As Single, ScaleX As Single, ScaleY As Single, Angle As Single)
Dim i As Integer
Dim bNum As Integer
bNum = GetBlock(DXF, Name)
For i = 0 To UBound(DXF.Blocks(bNum).Entities)
    DrawDXFGeometry Canvas, DXF, DXF.Blocks(bNum).Entities(), i, cX, cY, ScaleX, ScaleY, Angle
Next i
End Sub
Sub DrawDXFDImension(Canvas As PictureBox, DXF As DXFData, Name As String)
Dim i As Integer
Dim bNum As Integer
bNum = GetBlock(DXF, Name)
For i = 0 To UBound(DXF.Blocks(bNum).Entities)
    DrawDXFGeometry Canvas, DXF, DXF.Blocks(bNum).Entities(), i, 0, 0, 1, 1, 0
Next i
End Sub
Sub DrawDXFLine(Canvas As PictureBox, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, Color As Long)
Canvas.Line (X1, -Y1)-(X2, -Y2), Color
End Sub

Sub DrawDXFText(Canvas As PictureBox, X1 As Single, Y1 As Single, Angle As Single, Size As Single, Text As String, Color As Long)
Dim F As LOGFONT
Dim hPrevFont As Long
Dim hFont As Long
Dim FontName As String
Dim XSIZE As Integer
Dim YSIZE As Integer
F.lfEscapement = 10 * Val(Angle) 'rotation angle, in tenths
FontName = "Arial Black" + Chr$(0) 'null terminated
F.lfFacename = FontName
XSIZE = Canvas.ScaleX(Size, 0, 2)
YSIZE = Canvas.ScaleY(Size, 0, 2)
If XSIZE = 0 Then XSIZE = 1
If YSIZE = 0 Then YSIZE = 1
F.lfWidth = (XSIZE * -15) / Screen.TwipsPerPixelY
F.lfHeight = (YSIZE * -20) / Screen.TwipsPerPixelY
hFont = CreateFontIndirect(F)
hPrevFont = SelectObject(Canvas.hdc, hFont)
Canvas.ForeColor = Color
Canvas.CurrentX = X1
Canvas.CurrentY = -Y1 - Size
Canvas.Print Text

'  Clean up, restore original font
hFont = SelectObject(Canvas.hdc, hPrevFont)
DeleteObject hFont
End Sub

Sub DrawDXFArc(Canvas As PictureBox, X1 As Single, Y1 As Single, rad As Single, Angle1 As Single, Angle2 As Single, Color As Long)
Angle1 = dAngle(Angle1)
Angle2 = dAngle(Angle2)
Dim i As Single
Dim interval As Single
If Angle1 > Angle2 Then
    If Angle1 <> 360 Then Canvas.Circle (X1, -Y1), rad, Color, Angle1 * pi / 180, 2 * pi
    If Angle2 <> 0 Then Canvas.Circle (X1, -Y1), rad, Color, 0, Angle2 * pi / 180
Else
    'It's a good practice to ALWAYS split your arcs into sections
    'this method may not draw it properly
    'if the arc ever ends up being close to a circle (CLOSED)
    interval = (Angle2 - Angle1) / pi
    For i = Angle1 To Angle2 - interval Step interval
        Canvas.Circle (X1, -Y1), rad, Color, i * pi / 180, (i + interval) * pi / 180
    Next i
    Canvas.Circle (X1, -Y1), rad, Color, i * pi / 180, (Angle2) * pi / 180
End If
End Sub
Sub DrawDXFCircle(Canvas As PictureBox, X1 As Single, Y1 As Single, rad As Single, Color As Long)
Canvas.Circle (X1, -Y1), rad, Color
End Sub
Sub DrawDXFPoint(Canvas As PictureBox, X1 As Single, Y1 As Single, Color As Long)
Canvas.DrawWidth = 3
Canvas.PSet (X1, -Y1), Color
Canvas.DrawWidth = 1
End Sub

Sub DrawDXFEllipse(Canvas As PictureBox, cX As Single, cY As Single, mX As Single, mY As Single, Ratio As Single, Angle1 As Single, Angle2 As Single, NumPoints As Integer, Color)
'This was the HARDEST part of this project
'I don't know why . . it all seems simple now,
'but I had the hardest time figuring out how to rotate the ellipse
'none the less rotate an ellipse that isn't "closed"
'you CAN NOT simply use the windows API for drawing ellipses,
'because Windows does not allow rotation of the ellipse
Dim A As Single, B As Single
Dim RotAngle As Single
Dim A1 As Single, A2 As Single
Dim X1 As Single, Y1 As Single
Dim X2 As Single, Y2 As Single
Dim X3 As Single, Y3 As Single
Dim Hyp As Single
Dim j As Single
Dim U As Single
Dim Count As Integer
A = Sqr((mX * mX) + (mY * mY))
If mX < 0 Then A = -A
B = Ratio * A
If mX = 0 Then
    RotAngle = pi / 2
Else
    RotAngle = Atn(mY / mX)
End If
For U = Angle1 To Angle2 + (pi / (NumPoints * 2)) Step pi / NumPoints
    X1 = A * Cos(U)
    Y1 = B * Sin(U)
    Hyp = Sqr((X1 * X1) + (Y1 * Y1))
    If X1 = 0 Then
        j = pi / 2
    Else
        j = Atn(Y1 / X1)
    End If
    If X1 < 0 Then Hyp = -Hyp
    If (j * 180 / pi) + (RotAngle * 180 / pi) > 360 Then j = j + (2 * pi)
    X2 = (Hyp * Cos(RotAngle + j))
    Y2 = (Hyp * Sin(RotAngle + j))
    If Count > 0 Then Canvas.Line (cX + X3, -cY - Y3)-(cX + X2, -cY - Y2), Color
    X3 = X2
    Y3 = Y2
    Count = Count + 1
Next U
End Sub
Sub DrawDXFGeometry(Canvas As PictureBox, DXF As DXFData, Geo() As Geometry, Start As Integer, cX As Single, cY As Single, ScaleX As Single, ScaleY As Single, Angle As Single)
'When drawing geometry, and a 'modifier' is applied such as origin,scale or rotation
'you should follow the following order to draw geometry properly (when modified)
'--------
'SCALE
'ROTATION
'ORIGIN
'--------
On Error Resume Next
Dim Color As Long
Dim i As Integer
Dim X1 As Single
Dim Y1 As Single
Dim X2 As Single
Dim Y2 As Single
Dim X3 As Single
Dim Y3 As Single
Dim Angle1 As Single
Dim Angle2 As Single
Dim Angle3 As Single
Dim Ratio As Single
Dim rad As Single
Dim PCount As Integer
Dim Text As String
Dim Size As Single
Dim Name As String
Dim EndPoly As Boolean
Canvas.DrawWidth = 1
Canvas.DrawStyle = vbSolid
Color = vbBlack
Select Case Geo(Start).Type
    Case "LINE"
        'Get the values
        X1 = Geo(Start).Data(0).Value
        Y1 = Geo(Start).Data(1).Value
        X2 = Geo(Start).Data(2).Value
        Y2 = Geo(Start).Data(3).Value
        'Scale them relative to their origin
        X1 = X1 * ScaleX
        Y1 = Y1 * ScaleY
        X2 = X2 * ScaleX
        Y2 = Y2 * ScaleY
        'Rotate them relative to their origin
        If Angle <> 0 Then
            X3 = RotX(X1, Y1, Angle)
            Y3 = RotY(X1, Y1, Angle)
            X1 = X3
            Y1 = Y3
            X3 = RotX(X2, Y2, Angle)
            Y3 = RotY(X2, Y2, Angle)
            X2 = X3
            Y2 = Y3
        End If
        'Move the origin
        X1 = X1 + cX
        Y1 = Y1 + cY
        X2 = X2 + cX
        Y2 = Y2 + cY
        'Draw the line
        DrawDXFLine Canvas, X1, Y1, X2, Y2, Color
    Case "ARC"
        'Circles and arc's AUTOMATICALLY become ELLIPSES when scaled
        X1 = Geo(Start).Data(0).Value
        Y1 = Geo(Start).Data(1).Value
        rad = Geo(Start).Data(2).Value
        Angle1 = Geo(Start).Data(3).Value
        Angle2 = Geo(Start).Data(4).Value
        X1 = X1 * ScaleX
        Y1 = Y1 * ScaleY
        'You can't "STRETCH' an arc . . . or any BLOCK for that matter
        'If you stretch and ARC or a circle in the PV . . .it becomes an ellipse
        If ScaleX <> 1 Then
            rad = rad * ScaleX
        ElseIf ScaleY <> 1 Then
            rad = rad * ScaleY
        End If
        If Angle <> 0 Then
            X3 = RotX(X1, Y1, Angle)
            Y3 = RotY(X1, Y1, Angle)
            X1 = X3
            Y1 = Y3
        End If
        If ScaleX < 0 Or ScaleY < 0 Then
            'the ARC is mirrored
            Swap Angle1, Angle2
            Angle1 = 180 - Angle1
            Angle2 = 180 - Angle2
        End If
        Angle1 = Angle1 + (Angle * 180 / pi)
        Angle2 = Angle2 + (Angle * 180 / pi)
        X1 = X1 + cX
        Y1 = Y1 + cY
        DrawDXFArc Canvas, X1, Y1, Abs(rad), Angle1, Angle2, Color
    Case "CIRCLE"
        'Circles and arc's AUTOMATICALLY become ELLIPSES when scaled
        X1 = Geo(Start).Data(0).Value
        Y1 = Geo(Start).Data(1).Value
        rad = Geo(Start).Data(2).Value
        X1 = X1 * ScaleX
        Y1 = Y1 * ScaleY
        If ScaleX <> 1 Then
            rad = rad * ScaleX
        ElseIf ScaleY <> 1 Then
            rad = rad * ScaleY
        End If
        If Angle <> 0 Then
            X3 = RotX(X1, Y1, Angle)
            Y3 = RotY(X1, Y1, Angle)
            X1 = X3
            Y1 = Y3
        End If
        X1 = X1 + cX
        Y1 = Y1 + cY
        DrawDXFCircle Canvas, X1, Y1, Abs(rad), Color
    Case "ELLIPSE"
        X1 = Geo(Start).Data(0).Value
        Y1 = Geo(Start).Data(1).Value
        X2 = Geo(Start).Data(2).Value
        Y2 = Geo(Start).Data(3).Value
        Ratio = Geo(Start).Data(4).Value
        Angle1 = Geo(Start).Data(5).Value
        Angle2 = Geo(Start).Data(6).Value
        X1 = X1 * ScaleX
        Y1 = Y1 * ScaleY
        X2 = X2 * ScaleX
        Y2 = Y2 * ScaleY
        If Angle <> 0 Then
            X3 = RotX(X1, Y1, Angle)
            Y3 = RotY(X1, Y1, Angle)
            X1 = X3
            Y1 = Y3
            X3 = RotX(X2, Y2, Angle)
            Y3 = RotY(X2, Y2, Angle)
            X2 = X3
            Y2 = Y3
        End If
        If ScaleX < 0 Or ScaleY < 0 Then Ratio = -Ratio 'the ELLIPSE is mirrored
        X1 = X1 + cX
        Y1 = Y1 + cY
        DrawDXFEllipse Canvas, X1, Y1, X2, Y2, Ratio, Angle1, Angle2, 32, Color
    Case "POLYLINE"
        'a POLYLINE is a list of "VERTEX" points that are strung together
        PCount = 1
        EndPoly = False
        Do While Not EndPoly
            X1 = Geo(Start + PCount).Data(0).Value
            Y1 = Geo(Start + PCount).Data(1).Value
            X2 = Geo(Start + PCount + 1).Data(0).Value
            Y2 = Geo(Start + PCount + 1).Data(1).Value
            'Scale them relative to their origin
            X1 = X1 * ScaleX
            X2 = X2 * ScaleX
            Y1 = Y1 * ScaleY
            Y2 = Y2 * ScaleY
            'Rotate them relative to their origin
            If Angle <> 0 Then
                X3 = RotX(X1, Y1, Angle)
                Y3 = RotY(X1, Y1, Angle)
                X1 = X3
                Y1 = Y3
                X3 = RotX(X2, Y2, Angle)
                Y3 = RotY(X2, Y2, Angle)
                X2 = X3
                Y2 = Y3
            End If
            'Move the origin
            X1 = X1 + cX
            Y1 = Y1 + cY
            X2 = X2 + cX
            Y2 = Y2 + cY
            'Dray the line
            DrawDXFLine Canvas, X1, Y1, X2, Y2, Color
            PCount = PCount + 1
            If Start + PCount + 1 > UBound(Geo) Then
                EndPoly = True
            ElseIf Geo(Start + PCount + 1).Type <> "VERTEX" Then
                EndPoly = True
            End If
        Loop
    Case "TEXT"
        'there is no scaling for TEXT entities
        X1 = Geo(Start).Data(0).Value
        Y1 = Geo(Start).Data(1).Value
        Size = Geo(Start).Data(2).Value
        Angle1 = Geo(Start).Data(3).Value + Angle
        Text = Geo(Start).Data(4).Value
        'Move the origin
        X1 = X1 + cX
        Y1 = Y1 + cY
        DrawDXFText Canvas, X1, Y1, Angle1, Size, Text, Color
    Case "INSERT"
        'Just a note: BLOCKS can not be "Stretched" but if they are mirrored . . that
        'shows up in the "scale" dataset for BLOCKS
        Name = Geo(Start).Data(0).Value
        X1 = Geo(Start).Data(1).Value
        Y1 = Geo(Start).Data(2).Value
        X2 = Geo(Start).Data(3).Value
        Y2 = Geo(Start).Data(4).Value
        '"0" scale = scale of "1"
        If X2 = 0 Then X2 = 1
        If Y2 = 0 Then Y2 = 1
        Angle1 = Geo(Start).Data(5).Value * pi / 180
        DrawDXFBlock Canvas, DXF, Name, X1, Y1, X2, Y2, Angle1
    Case "DIMENSION"
        'Just a note: BLOCKS can not be "Stretched" but if they are mirrored . . that
        'shows up in the "scale" dataset for BLOCKS
        Name = Geo(Start).Data(0).Value
        X1 = Geo(Start).Data(1).Value
        Y1 = Geo(Start).Data(2).Value
        DrawDXFDImension Canvas, DXF, Name
End Select
End Sub

Sub FindCommand(FileNum As Integer, Command As String)
Dim X As String
Do While UCase(Trim(X)) <> UCase(Command)
    Line Input #FileNum, X
Loop
End Sub

Function GetBlock(DXF As DXFData, Name As String) As Integer
Dim i As Integer
For i = 0 To UBound(DXF.Blocks)
    If DXF.Blocks(i).Name = Name Then
        GetBlock = i
        Exit Function
    End If
Next i
End Function

Function GetSection(FileNum As Integer, Start As String, Finish As String, EndString As String, sArray() As String) As Boolean
ReDim sArray(0) As String
Dim Temp As String
Dim i As Long
Do While Temp <> Start
    Line Input #FileNum, Temp
    Temp = UCase(Trim(Temp))
    If Temp = EndString Then
        GetSection = False
        Exit Function
    End If
Loop
Do While Temp <> Finish
    Line Input #FileNum, Temp
    Temp = UCase(Trim(Temp))
    If Temp <> Finish Then
        ReDim Preserve sArray(i) As String
        sArray(i) = Temp
        i = i + 1
    End If
Loop
GetSection = True
End Function

Sub ImportDXF(FileDXF As String, ByRef DXF As DXFData)
Dim FF As Integer
Dim DXFLine As String
Dim bCount As Integer
Dim eCount As Integer
Dim ENDSEC As Boolean

ReDim DXF.Blocks(0) As Block
ReDim DXF.Entities(0) As Geometry
FF = FreeFile
Open FileDXF For Input As #FF
'First we skip all the header stuff and get to the section called 'BLOCKS'
FindCommand FF, "BLOCKS"
'---------------------------
'BLOCKS are groups of geometry that
'are re-useable within the drawing
'they may appear several times within one drawing
'and if the block is modified it automatically
'modifies each time wherev it's used within the drawing
Do While Not ENDSEC
    'First we load in a SECTION into an array (BLOCK) to (ENDBLK)
    'we do this until we come across the "ENDSEC" command
    If GetSection(FF, "BLOCK", "ENDBLK", "ENDSEC", Section()) Then
        'We have a "BLOCK" in the array
        'So we have to advance our array of BLOCKS
        ReDim Preserve DXF.Blocks(bCount) As Block
        ReDim Preserve DXF.Blocks(bCount).Entities(eCount) As Geometry
        If ParseBlock(Section(), DXF.Blocks(bCount)) Then
            bCount = bCount + 1
            eCount = 0
        End If
    Else
        ENDSEC = True
    End If
Loop
'Now we go after the 'Primary View Entities
ENDSEC = False
eCount = 0
GetSection FF, "ENTITIES", "ENDSEC", "ENDSEC", Section()
'This grabs ALL PV ENTITIES . . . kind of like one huge block
Close #FF 'We can close the file because we're finished with it
'Next we fill the array with geometry data
ParsePV Section(), DXF.Entities()

End Sub
Function IsCommand(InText As String)
Select Case UCase(InText)
    Case "LINE", "VERTEX", "POLYLINE", "CIRCLE", "ARC", "ELLIPSE", "TEXT", "INSERT", "DIMENSION"
        'These are the basic ENTITY COMMANDS available in the DXF language
        IsCommand = True
    Case Else
        IsCommand = False
End Select
End Function
Function kVal(Data() As DataSet, Key As Integer) As Variant
Dim i As Integer
For i = 0 To UBound(Data)
    If Data(i).Key = Key Then
        kVal = Data(i).Value
        Exit Function
    End If
Next i
kVal = 0
End Function

Function ParseBlock(sArray() As String, ByRef tBlock As Block) As Boolean
'On Local Error GoTo exitMe:
Dim i As Long
Dim j As Long
Dim k As Long
Dim p As Long
'first we have to look for a section "6" to determine if this BLOCK section is "important"
i = SearchSection(sArray(), i, "6")
If i = -1 Then
    ParseBlock = False
    Exit Function
End If
i = SearchSection(sArray(), i, "2") + 1
tBlock.Name = sArray(i)
For j = i To UBound(sArray)
    If IsCommand(sArray(j)) Then 'We Found an ENTITY COMMAND
        ReDim Preserve tBlock.Entities(k) As Geometry
        tBlock.Entities(k).Type = sArray(j)
        'I am not sure if a BLOCK can use a block.
        'Either way, this is designed to work even if you can
        Select Case tBlock.Entities(k).Type
            Case "INSERT", "DIMENSION"
                'KEY "2" on an INSERT provides the BLOCK name to be inserted
                j = SearchSection(sArray(), j, "2")
            Case Else
                j = FindStart(sArray(), j)
                'j = SearchSection(sArray(), j, "10")
        End Select
        Do While sArray(j) <> "0"
            ReDim Preserve tBlock.Entities(k).Data(p)
            tBlock.Entities(k).Data(p).Key = sArray(j)
            tBlock.Entities(k).Data(p).Value = sArray(j + 1)
            p = p + 1
            j = j + 2
        Loop
        PrepareEntity tBlock.Entities(k)
        k = k + 1
        p = 0
    End If
Next j
ParseBlock = True
Exit Function
exitMe:
MsgBox "ERROR  " & Err.Description
End Function

Function ParsePV(sArray() As String, ByRef tGeo() As Geometry) As Boolean
Dim i As Long
Dim j As Long
Dim k As Long
Dim p As Long
For j = i To UBound(sArray)
    If IsCommand(sArray(j)) Then 'we found an ENTITY COMMAND
        ReDim Preserve tGeo(k) As Geometry
        tGeo(k).Type = sArray(j)
        Select Case tGeo(k).Type
            Case "INSERT", "DIMENSION"
                'KEY "2" on an INSERT provides the BLOCK name to be inserted to the PV
                j = SearchSection(sArray(), j, "2")
            Case Else
                j = FindStart(sArray(), j)
                'j = SearchSection(sArray(), j, "10")
        End Select
        Do While sArray(j) <> "0"
            ReDim Preserve tGeo(k).Data(p)
            tGeo(k).Data(p).Key = sArray(j)
            tGeo(k).Data(p).Value = sArray(j + 1)
            p = p + 1
            j = j + 2
        Loop
        PrepareEntity tGeo(k)
        k = k + 1
        p = 0
    End If
Next j
ParsePV = True
End Function
Function RotX(X1 As Single, Y1 As Single, Angle As Single) As Single
RotX = cHyp(X1, Y1) * Cos(PtAng(X1, Y1) + Angle)
End Function

Function RotY(X1 As Single, Y1 As Single, Angle As Single) As Single
RotY = cHyp(X1, Y1) * Sin(PtAng(X1, Y1) + Angle)
End Function
Function SearchSection(sArray() As String, Start As Long, Value As String) As Long
Dim i As Long
For i = Start To UBound(sArray)
    If sArray(i) = Value Then
        SearchSection = i
        Exit Function
    End If
Next i
SearchSection = -1
End Function


Sub Swap(ByRef A As Variant, ByRef B As Variant)
Dim C As Variant
C = A
A = B
B = C
End Sub


