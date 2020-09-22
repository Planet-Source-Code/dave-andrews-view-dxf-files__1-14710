VERSION 5.00
Begin VB.Form frmBrowser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DXF Browser"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picControls 
      Align           =   4  'Align Right
      Height          =   7380
      Left            =   7410
      ScaleHeight     =   7320
      ScaleWidth      =   2460
      TabIndex        =   1
      Top             =   0
      Width           =   2520
      Begin VB.Frame frmView 
         Caption         =   "View"
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   6600
         Width           =   2295
         Begin VB.TextBox txtView 
            Height          =   285
            Left            =   1560
            TabIndex        =   16
            Top             =   215
            Width           =   495
         End
         Begin VB.VScrollBar vscrView 
            Height          =   450
            Left            =   2040
            TabIndex        =   15
            Top             =   120
            Width           =   190
         End
         Begin VB.OptionButton optView 
            Caption         =   "Blocks"
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   14
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optView 
            Caption         =   "PV"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.Frame frameMouse 
         Caption         =   "Mouse"
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   5880
         Width           =   2295
         Begin VB.CommandButton cmdZoomIn 
            Caption         =   "+"
            Height          =   255
            Left            =   1680
            TabIndex        =   11
            Top             =   240
            Width           =   255
         End
         Begin VB.CommandButton cmdZoomOut 
            Caption         =   "-"
            Height          =   255
            Left            =   1920
            TabIndex        =   10
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton optMouse 
            Caption         =   " Zoom"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   9
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optMouse 
            Caption         =   "Center"
            Height          =   255
            Index           =   4
            Left            =   3600
            TabIndex        =   8
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optMouse 
            Caption         =   "Pan"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.ListBox List1 
         Height          =   1980
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   5
         Top             =   3840
         Width           =   2295
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   2295
      End
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
      Begin VB.FileListBox File1 
         Height          =   1650
         Left            =   120
         Pattern         =   "*.dxf"
         TabIndex        =   2
         Top             =   2160
         Width           =   2295
      End
   End
   Begin VB.PictureBox picDXF 
      AutoRedraw      =   -1  'True
      Height          =   7380
      Left            =   0
      ScaleHeight     =   100
      ScaleLeft       =   -5
      ScaleMode       =   0  'User
      ScaleTop        =   -95
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   0
      Width           =   7380
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MyDXF As DXFData
Dim dragX As Single
Dim dragY As Single
Dim SelGroup As RECT
Dim Pan As Boolean
Dim Zoom As Boolean
Sub RedrawPic()
If optView(0) Then
    DrawDXF picDXF, MyDXF
Else
    DrawBlock picDXF, MyDXF, vscrView.Value
End If
End Sub

Private Sub cmdZoomIn_Click()
picDXF.ScaleHeight = 0.75 * picDXF.ScaleHeight
picDXF.ScaleWidth = 0.75 * picDXF.ScaleWidth
RedrawPic
End Sub

Private Sub cmdZoomOut_Click()
picDXF.ScaleHeight = 1.25 * picDXF.ScaleHeight
picDXF.ScaleWidth = 1.25 * picDXF.ScaleWidth
RedrawPic
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub


Private Sub File1_Click()
ImportDXF File1.Path & "\" & File1.filename, MyDXF
vscrView.Value = 0
vscrView.Max = UBound(MyDXF.Blocks)
RedrawPic
'Exit Sub
'This next part is merely a view of the dxf data
On Error Resume Next
List1.Clear
Dim i As Integer
Dim j As Integer
Dim k As Integer
For i = 0 To UBound(MyDXF.Blocks)
    List1.AddItem "-" & MyDXF.Blocks(i).Name
    For j = 0 To UBound(MyDXF.Blocks(i).Entities)
        List1.AddItem "--" & MyDXF.Blocks(i).Entities(j).Type
        For k = 0 To UBound(MyDXF.Blocks(i).Entities(j).Data)
            List1.AddItem "---" & MyDXF.Blocks(i).Entities(j).Data(k).Key & " = " & MyDXF.Blocks(i).Entities(j).Data(k).Value
        Next k
    Next j
Next i
List1.AddItem "--------------"
For i = 0 To UBound(MyDXF.Entities)
    List1.AddItem "PV -" & MyDXF.Entities(i).Type
    For k = 0 To UBound(MyDXF.Entities(i).Data)
        List1.AddItem "---" & MyDXF.Entities(i).Data(k).Key & " = " & MyDXF.Entities(i).Data(k).Value
    Next k
Next i
End Sub

Private Sub optView_Click(Index As Integer)
txtView = MyDXF.Blocks(vscrView.Value).Name
RedrawPic
End Sub

Private Sub picDXF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
dragX = X
dragY = Y
SelGroup.X1 = X
SelGroup.Y1 = Y
SelGroup.X2 = X
SelGroup.Y2 = Y
If optMouse(0) Then Pan = True
If optMouse(1) Then Zoom = True
End Sub


Private Sub picDXF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Caption = "DXF Browser -- (" & Format(X, "0.000") & "," & Format(-Y, "0.000") & ")"
If Pan Then
    picDXF.ScaleTop = picDXF.ScaleTop + (dragY - Y)
    picDXF.ScaleLeft = picDXF.ScaleLeft + (dragX - X)
    picDXF.Cls
    picDXF.Picture = LoadPicture()
    RedrawPic
    Exit Sub
End If
If Zoom Then
    picDXF.DrawMode = 6
    picDXF.DrawStyle = 1
    picDXF.DrawWidth = 1
    picDXF.Line (SelGroup.X1, SelGroup.Y1)-(SelGroup.X2, SelGroup.Y2), vbBlack, B
    picDXF.Line (SelGroup.X1, SelGroup.Y1)-(X, Y), vbBlack, B
    SelGroup.X2 = X
    SelGroup.Y2 = Y
End If
End Sub


Private Sub picDXF_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoEvents
picDXF.DrawMode = 13
picDXF.DrawStyle = 0
picDXF.DrawWidth = 1
If Zoom Then
    If SelGroup.X2 < SelGroup.X1 Then Swap SelGroup.X1, SelGroup.X2
    If SelGroup.Y2 < SelGroup.Y1 Then Swap SelGroup.Y1, SelGroup.Y2
    SelGroup.Y2 = SelGroup.Y1 + Abs(SelGroup.X2 - SelGroup.X1)
    If SelGroup.X2 = SelGroup.X1 Then Exit Sub
    If SelGroup.Y2 = SelGroup.Y1 Then Exit Sub
    picDXF.ScaleWidth = Abs(SelGroup.X2 - SelGroup.X1)
    picDXF.ScaleLeft = SelGroup.X1
    picDXF.ScaleHeight = Abs(SelGroup.Y1 - SelGroup.Y2)
    picDXF.ScaleTop = SelGroup.Y1
    RedrawPic
End If
Pan = False
Zoom = False
End Sub


Private Sub txtView_Change()
RedrawPic
End Sub

Private Sub vscrView_Change()
txtView = MyDXF.Blocks(vscrView.Value).Name
End Sub


Private Sub vscrView_Scroll()
txtView = MyDXF.Blocks(vscrView.Value).Name

End Sub


