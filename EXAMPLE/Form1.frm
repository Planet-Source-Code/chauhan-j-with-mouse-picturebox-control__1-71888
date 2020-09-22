VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox SHOWPicture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1980
      Left            =   720
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   0
      Top             =   720
      Width           =   1980
   End
   Begin VB.Label GRPPicture 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   13000
      Width           =   1065
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim XGrip(2) As Long, YGrip(2) As Long
Dim bMoving As Boolean, bSizing As Boolean
Dim xStart As Long, yStart As Long
Const GripSize = 90
Private Sub MoveGrips()
   On Error Resume Next
   XGrip(0) = SHOWPicture.Left - GripSize
   XGrip(1) = SHOWPicture.Left + SHOWPicture.Width / 2 - GripSize / 2
   XGrip(2) = SHOWPicture.Left + SHOWPicture.Width
   YGrip(0) = SHOWPicture.Top - GripSize
   YGrip(1) = SHOWPicture.Top + SHOWPicture.Height / 2 - GripSize / 2
   YGrip(2) = SHOWPicture.Top + SHOWPicture.Height
   GRPPicture(0).Move XGrip(0), YGrip(0)
   GRPPicture(1).Move XGrip(0), YGrip(1)
   GRPPicture(2).Move XGrip(0), YGrip(2)
   GRPPicture(3).Move XGrip(1), YGrip(2)
   GRPPicture(4).Move XGrip(2), YGrip(2)
   GRPPicture(5).Move XGrip(2), YGrip(1)
   GRPPicture(6).Move XGrip(2), YGrip(0)
   GRPPicture(7).Move XGrip(1), YGrip(0)
End Sub
Private Sub ShowGrip(bShow As Boolean)
   On Error Resume Next
   Dim i As Integer
   SHOWPicture.Move 100, 100, 600, 600
   SHOWPicture.Visible = bShow
   For i = 0 To 7
      GRPPicture(i).Visible = bShow
   Next i
  SHOWPicture.Height = 1900
 SHOWPicture.Width = 1900
 MoveGrips
End Sub

Private Sub InitGrip()
   On Error Resume Next
   Dim i As Integer
   GRPPicture(0).Width = GripSize
   GRPPicture(0).Height = GripSize
   For i = 1 To 7
      Load GRPPicture(i)
      GRPPicture(i).MousePointer = i + 4 * Int((9 - i) / 4)
   Next i
   GRPPicture(0).MousePointer = 8
   ShowGrip False
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()
 AddINFORM SHOWPicture, "Qkby  u;k   [kksyks  j[kks caa| djkks", Form1
End Sub

Private Sub Form_Load()
 On Error Resume Next


  
 

SHOWPicture.Picture = LoadPicture(App.Path & "\G-1.JPG")


InitGrip
picStrech
ShowGrip True
 
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim lft As Long, tp As Long
      For i = 0 To 7
 GRPPicture(i).Visible = False
Next i
End Sub

Private Sub GRPPicture_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
   bSizing = False
   SHOWPicture.Enabled = True

End Sub
Private Sub SHOWPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
  Dim lft As Long, tp As Long
      For i = 0 To 7
 GRPPicture(i).Visible = True
Next i
   If Button = vbLeftButton Then
      bMoving = True
      xStart = X: yStart = Y
      SHOWPicture.MousePointer = 5
End If
End Sub

Private Sub SHOWPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   Dim lft As Long, tp As Long
   
   If bMoving Then
      lft = SHOWPicture.Left + X - xStart
      tp = SHOWPicture.Top + Y - yStart
      If lft <= 0 Then lft = 0
      If tp <= 0 Then tp = 0
      
      SHOWPicture.Move lft, tp
      MoveGrips
   End If
End Sub
Private Sub GRPPicture_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
   
   
   If Button = vbLeftButton Then
      bSizing = True
      xStart = X: yStart = Y
      SHOWPicture.Enabled = False
   
   End If
End Sub

Private Sub GRPPicture_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
   On Error Resume Next
   Dim lft As Long, tp As Long, wdt As Long, hgt As Long
   If bSizing Then
     
picStrech
 
      
      Select Case Index
         Case 0
              lft = SHOWPicture.Left + X - xStart
              tp = SHOWPicture.Top + Y - yStart
              wdt = SHOWPicture.Width - X + xStart
              hgt = SHOWPicture.Height - Y + yStart
         Case 1
              lft = SHOWPicture.Left + X - xStart
              tp = SHOWPicture.Top
              wdt = SHOWPicture.Width - X + xStart
              hgt = SHOWPicture.Height
         Case 2
              lft = SHOWPicture.Left + X - xStart
              tp = SHOWPicture.Top
              wdt = SHOWPicture.Width - X + xStart
              hgt = SHOWPicture.Height + Y - yStart
         Case 3
              lft = SHOWPicture.Left
              tp = SHOWPicture.Top
              wdt = SHOWPicture.Width
              hgt = SHOWPicture.Height + Y - yStart
         Case 4
              lft = SHOWPicture.Left
              tp = SHOWPicture.Top
              wdt = SHOWPicture.Width + X - xStart
              hgt = SHOWPicture.Height + Y - yStart
         Case 5
              lft = SHOWPicture.Left
              tp = SHOWPicture.Top
              wdt = SHOWPicture.Width + X - xStart
              hgt = SHOWPicture.Height
         Case 6
              lft = SHOWPicture.Left
              tp = SHOWPicture.Top + Y - yStart
              wdt = SHOWPicture.Width + X - xStart
              hgt = SHOWPicture.Height - Y + yStart
         Case 7
              lft = SHOWPicture.Left
              tp = SHOWPicture.Top + Y - yStart
              wdt = SHOWPicture.Width
              hgt = SHOWPicture.Height - Y + yStart
   
      End Select
      
      SHOWPicture.Move lft, tp, wdt, hgt
      MoveGrips
   
   End If
End Sub
Private Sub SHOWPicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
   bMoving = False
   SHOWPicture.MousePointer = 0

End Sub
Sub picStrech()
    On Error Resume Next
    SHOWPicture.ScaleMode = 3
    SHOWPicture.AutoRedraw = True
    SHOWPicture.PaintPicture SHOWPicture.Picture, _
        0, 0, SHOWPicture.ScaleWidth, SHOWPicture.ScaleHeight, _
        0, 0, _
        SHOWPicture.Picture.Width / 26.46, _
        SHOWPicture.Picture.Height / 26.46
   End Sub

Private Sub lblShape_Click()

End Sub
