VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Alpha Blend Menu Example - © 2000 Aaron Wilkes"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   3360
      Max             =   255
      TabIndex        =   11
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reset Form"
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show Picture"
      Height          =   615
      Left            =   3600
      TabIndex        =   9
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4080
      Top             =   1560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Background Color"
      Height          =   615
      Left            =   3600
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3600
      Top             =   1560
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   0
      Picture         =   "Transparent Menu.frx":0000
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   1
      Top             =   730
      Visible         =   0   'False
      Width           =   915
      Begin VB.Label MnuLab 
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label MnuLab 
         BackStyle       =   0  'Transparent
         Caption         =   "Print"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label MnuLab 
         BackStyle       =   0  'Transparent
         Caption         =   "Save"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   495
      End
      Begin VB.Label MnuLab 
         BackStyle       =   0  'Transparent
         Caption         =   "Open"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Tag             =   "32"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label MnuLab 
         BackStyle       =   0  'Transparent
         Caption         =   "New"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Tag             =   "8"
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   0
      Picture         =   "Transparent Menu.frx":6C12
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   0
      Width           =   5625
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   465
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Opacity Value: "
      Height          =   315
      Left            =   3480
      TabIndex        =   12
      Top             =   840
      Width           =   2010
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   5160
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   4560
      Picture         =   "Transparent Menu.frx":148A4
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
Const AC_SRC_OVER = &H0
Dim I As Long, A As Long
Dim Opacity As Integer
Dim BF As BLENDFUNCTION, lBF As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Dim MnuFileOpened As Boolean

Private Sub Command1_Click()
If MnuFileOpened Then
Timer2.Enabled = True
MnuFileOpened = False
End If
Form1.BackColor = vbRed
Picture1.Cls
With BF
.BlendOp = AC_SRC_OVER
.BlendFlags = 0
.SourceConstantAlpha = Opacity
.AlphaFormat = 0
End With
RtlMoveMemory lBF, BF, 4
AlphaBlend Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Form1.hdc, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, lBF
Picture1.Refresh
End Sub

Private Sub Command2_Click()
If MnuFileOpened Then
Timer2.Enabled = True
MnuFileOpened = False
End If
Form1.Picture = Image1.Picture
Picture1.Cls
With BF
.BlendOp = AC_SRC_OVER
.BlendFlags = 0
.SourceConstantAlpha = Opacity
.AlphaFormat = 0
End With
RtlMoveMemory lBF, BF, 4
AlphaBlend Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Form1.hdc, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, lBF
Picture1.Refresh
End Sub

Private Sub Command3_Click()
If MnuFileOpened Then
Timer2.Enabled = True
MnuFileOpened = False
End If
Form1.BackColor = &H8000000F
Form1.Picture = Image2.Picture
Picture1.Cls
With BF
.BlendOp = AC_SRC_OVER
.BlendFlags = 0
.SourceConstantAlpha = Opacity
.AlphaFormat = 0
End With
RtlMoveMemory lBF, BF, 4
AlphaBlend Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Form1.hdc, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, lBF
Picture1.Refresh
End Sub

Private Sub Form_Click()
If MnuFileOpened Then
Timer2.Enabled = True
MnuFileOpened = False
End If
End Sub

Private Sub Form_Load()
HScroll1.Value = 128
For I = 0 To MnuLab.Count - 1
MnuLab(I).Tag = MnuLab(I).Top
MnuLab(I).Visible = False
Next I
Picture1.Visible = True
Form1.AutoRedraw = True
Picture1.AutoRedraw = True
Form1.ScaleMode = vbPixels
Picture1.ScaleMode = vbPixels
With BF
.BlendOp = AC_SRC_OVER
.BlendFlags = 0
.SourceConstantAlpha = Opacity
.AlphaFormat = 0
End With
RtlMoveMemory lBF, BF, 4
AlphaBlend Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Form1.hdc, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, lBF
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = RGB(0, 0, 0)
For A = 0 To MnuLab.Count - 1
MnuLab(A).ForeColor = RGB(0, 0, 0)
Next A
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub HScroll1_Change()
Label2.Caption = "Opacity Value" & Str$(HScroll1.Value)
Opacity = HScroll1.Value
Picture1.Cls
With BF
.BlendOp = AC_SRC_OVER
.BlendFlags = 0
.SourceConstantAlpha = Opacity
.AlphaFormat = 0
End With
RtlMoveMemory lBF, BF, 4
AlphaBlend Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Form1.hdc, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, lBF
Picture1.Refresh
End Sub

Private Sub HScroll1_Scroll()
Label2.Caption = "Opacity Value" & Str$(HScroll1.Value)
Opacity = HScroll1.Value
Picture1.Cls
With BF
.BlendOp = AC_SRC_OVER
.BlendFlags = 0
.SourceConstantAlpha = Opacity
.AlphaFormat = 0
End With
RtlMoveMemory lBF, BF, 4
AlphaBlend Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Form1.hdc, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, lBF
Picture1.Refresh
End Sub

Private Sub Label1_Click()
If Not MnuFileOpened Then
Picture2.Visible = True
Form1.AutoRedraw = True
Picture2.AutoRedraw = True
Form1.ScaleMode = vbPixels
Picture2.ScaleMode = vbPixels
With BF
.BlendOp = AC_SRC_OVER
.BlendFlags = 0
.SourceConstantAlpha = 255
.AlphaFormat = 0
End With
RtlMoveMemory lBF, BF, 4
AlphaBlend Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Form1.hdc, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, lBF
Timer1.Enabled = True
MnuFileOpened = True
End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = RGB(0, 255, 0)
If Label1.Top < 9 Then
Label1.Top = Label1.Top + 8
End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = RGB(0, 255, 0)
For A = 0 To MnuLab.Count - 1
MnuLab(A).ForeColor = RGB(0, 0, 0)
Next A
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = RGB(0, 255, 0)
If Label1.Top > 14 Then
Label1.Top = Label1.Top - 8
End If
End Sub

Private Sub MnuLab_Click(Index As Integer)
Dim Ret As Long
Select Case Index
Case 0
MsgBox "New", vbOKOnly + vbInformation, "You Chose..."
Case 1
MsgBox "Open", vbOKOnly + vbInformation, "You Chose..."
Case 2
MsgBox "Save", vbOKOnly + vbInformation, "You Chose..."
Case 3
MsgBox "Print", vbOKOnly + vbInformation, "You Chose..."
Case 4
MsgBox "Quit", vbOKOnly + vbInformation, "You Chose..."
Ret = MsgBox("Do you want to quit?", vbYesNo + vbInformation, "Quit?!?")
If Ret = vbYes Then
Unload Me
End If
End Select
End Sub

Private Sub MnuLab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If MnuLab(Index).Top < MnuLab(Index).Tag + 1 Then
MnuLab(Index).Top = MnuLab(Index).Top + 4
End If
End Sub

Private Sub MnuLab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For A = 0 To MnuLab.Count - 1
MnuLab(A).ForeColor = RGB(0, 0, 0)
Next A
MnuLab(Index).ForeColor = RGB(0, 255, 0)
End Sub

Private Sub MnuLab_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If MnuLab(Index).Top > MnuLab(Index).Tag + 1 Then
MnuLab(Index).Top = MnuLab(Index).Top - 4
End If
End Sub

Private Sub Picture1_Click()
If MnuFileOpened Then
Timer2.Enabled = True
MnuFileOpened = False
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = RGB(0, 0, 0)
For A = 0 To MnuLab.Count - 1
MnuLab(A).ForeColor = RGB(0, 0, 0)
Next A
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For A = 0 To MnuLab.Count - 1
MnuLab(A).ForeColor = RGB(0, 0, 0)
Next A
End Sub

Private Sub Timer1_Timer()
For I = 255 To Opacity Step -30
If I > Opacity - 1 Then
For A = 0 To MnuLab.Count - 1
MnuLab(A).Visible = True
Next A
End If
Timer1.Enabled = False
Picture2.Cls
With BF
.BlendOp = AC_SRC_OVER
.BlendFlags = 0
.SourceConstantAlpha = I
.AlphaFormat = 0
End With
RtlMoveMemory lBF, BF, 4
AlphaBlend Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Form1.hdc, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, lBF
Picture2.Refresh
Next I
End Sub

Private Sub Timer2_Timer()
For I = Opacity To 255 Step 30
If I > 254 - 30 Then
Picture2.Visible = False
Timer2.Enabled = False
For A = 0 To MnuLab.Count - 1
MnuLab(A).Visible = False
Next A
End If
Picture2.Cls
With BF
.BlendOp = AC_SRC_OVER
.BlendFlags = 0
.SourceConstantAlpha = I
.AlphaFormat = 0
End With
RtlMoveMemory lBF, BF, 4
AlphaBlend Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Form1.hdc, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, lBF
Picture2.Refresh
Next I
End Sub
