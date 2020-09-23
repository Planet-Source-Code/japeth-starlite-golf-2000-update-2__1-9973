VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Caption         =   "Course Editor"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox ToolWater 
      Height          =   495
      Left            =   9120
      Picture         =   "Form2.frx":030A
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   59
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame4 
      Caption         =   "Properties"
      Height          =   1815
      Left            =   5400
      TabIndex        =   25
      Top             =   3480
      Width           =   2895
      Begin VB.PictureBox ColorSel 
         Height          =   495
         Left            =   1440
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   27
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox ToolSel 
         Height          =   495
         Left            =   1440
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Coordinates:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Coordinate 
         BackStyle       =   0  'Transparent
         Caption         =   "0 X 0"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1080
         TabIndex        =   30
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Selected Color:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Selected Tool:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.PictureBox ToolCircle 
      Height          =   495
      Left            =   9120
      Picture         =   "Form2.frx":0D44
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   44
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox ToolBox 
      Height          =   495
      Left            =   9120
      Picture         =   "Form2.frx":177E
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   43
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox ToolDot 
      Height          =   495
      Left            =   9120
      Picture         =   "Form2.frx":21B8
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   42
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox ToolGreen 
      Height          =   495
      Left            =   9120
      Picture         =   "Form2.frx":2BF2
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   41
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox ToolFill 
      Height          =   495
      Left            =   9120
      Picture         =   "Form2.frx":362C
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   40
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame5 
      Caption         =   "Tools"
      Height          =   3375
      Left            =   6240
      TabIndex        =   34
      Top             =   0
      Width           =   2055
      Begin VB.TextBox WaterSize2 
         Height          =   285
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   61
         Text            =   "4"
         Top             =   2620
         Width           =   375
      End
      Begin VB.CheckBox LeftOption2 
         Caption         =   "Left"
         Height          =   255
         Left            =   1440
         TabIndex        =   57
         Top             =   2280
         Width           =   570
      End
      Begin VB.TextBox WaterSize 
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   56
         Text            =   "3"
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton CmdWater 
         Height          =   375
         Left            =   120
         Picture         =   "Form2.frx":4066
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton CmdGreen 
         Height          =   375
         Left            =   120
         Picture         =   "Form2.frx":428C
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2880
         Width           =   375
      End
      Begin VB.PictureBox ColorSel2 
         Height          =   255
         Left            =   1080
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   54
         Top             =   1850
         Width           =   735
      End
      Begin VB.TextBox BoxL 
         Height          =   285
         Left            =   760
         MaxLength       =   3
         TabIndex        =   50
         Text            =   "12"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox BoxW 
         Height          =   285
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   49
         Text            =   "12"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox CircleSize 
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   47
         Text            =   "1"
         Top             =   840
         Width           =   375
      End
      Begin VB.CheckBox LeftOption 
         Caption         =   "Left"
         Height          =   255
         Left            =   1440
         TabIndex        =   46
         Top             =   840
         Width           =   570
      End
      Begin VB.CommandButton CmdFill 
         Height          =   375
         Left            =   120
         Picture         =   "Form2.frx":445A
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton CmdBox 
         Height          =   375
         Left            =   120
         Picture         =   "Form2.frx":48C4
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox AutoFinish 
         Caption         =   "Auto Finish"
         Height          =   255
         Left            =   600
         TabIndex        =   45
         Top             =   400
         Width           =   1095
      End
      Begin VB.CommandButton CmdCircle 
         Height          =   375
         Left            =   120
         Picture         =   "Form2.frx":4CE2
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton CmdDot 
         Height          =   375
         Left            =   120
         Picture         =   "Form2.frx":4EB0
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   360
         Width           =   375
      End
      Begin VB.Line Line6 
         X1              =   600
         X2              =   1800
         Y1              =   320
         Y2              =   320
      End
      Begin VB.Line Line5 
         X1              =   600
         X2              =   1800
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line4 
         X1              =   600
         X2              =   1800
         Y1              =   2200
         Y2              =   2200
      End
      Begin VB.Line Line3 
         X1              =   600
         X2              =   1800
         Y1              =   1720
         Y2              =   1720
      End
      Begin VB.Line Line2 
         X1              =   600
         X2              =   1800
         Y1              =   1240
         Y2              =   1240
      End
      Begin VB.Line Line1 
         X1              =   600
         X2              =   1800
         Y1              =   760
         Y2              =   760
      End
      Begin VB.Label Label12 
         Caption         =   "Wave Size:"
         Height          =   255
         Left            =   600
         TabIndex        =   60
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Size:"
         Height          =   255
         Left            =   600
         TabIndex        =   58
         Top             =   2310
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Color:"
         Height          =   255
         Left            =   600
         TabIndex        =   53
         Top             =   1850
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "L:"
         Height          =   255
         Left            =   600
         TabIndex        =   52
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label5 
         Caption         =   "W:"
         Height          =   255
         Left            =   1200
         TabIndex        =   51
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Size:"
         Height          =   255
         Left            =   600
         TabIndex        =   48
         Top             =   870
         Width           =   375
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Redo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   33
      Top             =   8040
      Width           =   1335
   End
   Begin VB.PictureBox CourseRedo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10920
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   55
      TabIndex        =   32
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Undo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   24
      Top             =   8040
      Width           =   1335
   End
   Begin VB.PictureBox CourseUndo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   10920
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   55
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9720
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox CourseFinal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10920
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   55
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colors"
      Height          =   3375
      Left            =   5400
      TabIndex        =   1
      Top             =   0
      Width           =   735
      Begin VB.PictureBox ColorRough 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox ColorFairWay 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   5
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox ColorGreen 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   4
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox ColorWater 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   3
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox ColorSand 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   2
         Top             =   2760
         Width           =   495
      End
   End
   Begin VB.PictureBox CourseEdit 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   8400
      Left            =   120
      Picture         =   "Form2.frx":5366
      ScaleHeight     =   556
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   346
      TabIndex        =   0
      Top             =   120
      Width           =   5250
      Begin VB.Shape BallPlace 
         BackStyle       =   1  'Opaque
         Height          =   105
         Left            =   120
         Shape           =   3  'Circle
         Top             =   8160
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   3135
      Left            =   5400
      TabIndex        =   8
      Top             =   5400
      Width           =   2895
      Begin VB.TextBox StartPosY 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   22
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox StartPosX 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   21
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Play Course"
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Sel Flag"
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Sel Start"
         Height          =   375
         Left            =   1440
         TabIndex        =   14
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox FlagPosY 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   19
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox FlagPosX 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   18
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Load"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clear"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox ParNum 
         Height          =   285
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   10
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Start Pos:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Flag Pos:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Par:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Image IconGreen 
      Height          =   480
      Left            =   9720
      Picture         =   "Form2.frx":92668
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IconFill 
      Height          =   480
      Left            =   9720
      Picture         =   "Form2.frx":93332
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IconPencil 
      Height          =   480
      Left            =   9720
      Picture         =   "Form2.frx":93FFC
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Undo As Integer
Dim xx, yy As Integer
Dim sX, sY As Long
Dim MD As Boolean
Dim i, p As Integer
Dim StartX, StartY As Integer
Dim FlagX, FlagY As Integer
Dim SelStart, SelFlag As Boolean
Dim ShiftDown As Boolean
Dim DotX, DotY As Integer

Private Sub CmdBox_Click()
ToolSel.Picture = ToolBox.Picture
CourseEdit.MouseIcon = IconPencil.Picture
CourseEdit.MousePointer = 99
End Sub

Private Sub CmdCircle_Click()
ToolSel.Picture = ToolCircle.Picture
CourseEdit.MouseIcon = IconPencil.Picture
CourseEdit.MousePointer = 99
End Sub

Private Sub CmdDot_Click()
ToolSel.Picture = ToolDot.Picture
CourseEdit.MouseIcon = IconPencil.Picture
CourseEdit.MousePointer = 99
End Sub

Private Sub CmdFill_Click()
ToolSel.Picture = ToolFill.Picture
CourseEdit.MouseIcon = IconFill.Picture
CourseEdit.MousePointer = 99
End Sub

Private Sub CmdGreen_Click()
ToolSel.Picture = ToolGreen.Picture
CourseEdit.MouseIcon = IconGreen.Picture
CourseEdit.MousePointer = 99
End Sub

Private Sub CmdWater_Click()
ToolSel.Picture = ToolWater.Picture
CourseEdit.MouseIcon = IconPencil.Picture
CourseEdit.MousePointer = 99
End Sub

Private Sub ColorFairWay_Click()
ColorSel.BackColor = ColorFairWay.BackColor
ColorSel2.BackColor = ColorFairWay.BackColor
End Sub

Private Sub ColorGreen_Click()
ColorSel.BackColor = ColorGreen.BackColor
ColorSel2.BackColor = ColorGreen.BackColor
End Sub

Private Sub ColorRough_Click()
ColorSel.BackColor = ColorRough.BackColor
ColorSel2.BackColor = ColorRough.BackColor
End Sub

Private Sub ColorSand_Click()
ColorSel.BackColor = ColorSand.BackColor
ColorSel2.BackColor = ColorSand.BackColor
End Sub

Private Sub ColorWater_Click()
ColorSel.BackColor = ColorWater.BackColor
ColorSel2.BackColor = ColorWater.BackColor
End Sub

Private Sub Command1_Click()
CommonDialog1.Filter = "Bitmaps|*.bmp"
CommonDialog1.ShowSave
If CommonDialog1.FileName = "" Then
    Exit Sub
    End If
CourseEdit.ForeColor = RGB(255, 255, 255)
CourseEdit.PSet (StartX, StartY)
CourseFinal.Height = CourseEdit.Height
CourseFinal.Width = CourseEdit.Width
Call BitBlt(CourseFinal.hdc, 0, 0, CourseEdit.ScaleWidth, CourseEdit.ScaleHeight, CourseEdit.hdc, 0, 0, SRCAND)
Call BitBlt(CourseFinal.hdc, 0, 0, CourseEdit.ScaleWidth, CourseEdit.ScaleHeight, CourseEdit.hdc, 0, 0, SRCPAINT)
SavePicture CourseFinal.Image, CommonDialog1.FileName
End Sub

Private Sub Command2_Click()
CourseEdit.Cls
BallPlace.Visible = False
StartX = 0
StartY = 0
FlagX = 0
FlagY = 0
End Sub

Private Sub Command3_Click()
If FlagX = 0 Or StartX = 0 Then
    MsgBox "Must Determine Flag Position and Start Position", vbCritical, "JeroGolf"
    Exit Sub
    End If
CourseFinal.Height = CourseEdit.Height
CourseFinal.Width = CourseEdit.Width
Call BitBlt(CourseFinal.hdc, 0, 0, CourseEdit.ScaleWidth, CourseEdit.ScaleHeight, CourseEdit.hdc, 0, 0, SRCAND)
Call BitBlt(CourseFinal.hdc, 0, 0, CourseEdit.ScaleWidth, CourseEdit.ScaleHeight, CourseEdit.hdc, 0, 0, SRCPAINT)
Dim ran As Integer
Form1.Ball.Left = StartX
Form1.Ball.Top = StartY
CourseEdit.ForeColor = GetPixel(CourseEdit.hdc, StartX + 1, StartY + 2)
CourseEdit.PSet (BallPlace.Left + 2, BallPlace.Top + 2)
Form1.CoursePic.Picture = CourseFinal.Image
Form1.WindSpeed.Text = Int(Rnd * 16)
ran = Int(Rnd * 4)
If ran = 0 Then
    Form1.WindPic.Picture = Form1.WindLeft.Picture
    End If
If ran = 1 Then
    Form1.WindPic.Picture = Form1.WindRight.Picture
    End If
If ran = 2 Then
    Form1.WindPic.Picture = Form1.WindUp.Picture
    End If
If ran = 3 Then
    Form1.WindPic.Picture = Form1.WindDown.Picture
    End If
Form1.ShotLine.X1 = Form1.Ball.Left + Form1.Ball.Width / 2
Form1.ShotLine.X2 = Form1.ShotLine.X1

Form1.ShotLine.Y1 = Form1.Ball.Top + Form1.Ball.Height / 2
Form1.ShotLine.Y2 = Form1.Ball.Top - 50

Dim aa, bb As Integer
Form1.MeaLine.X1 = Form1.Ball.Left
Form1.MeaLine.Y1 = Form1.Ball.Top
If Form1.MeaLine.Y1 > Form1.MeaLine.Y2 Then
    aa = Form1.MeaLine.Y1 - Form1.MeaLine.Y2
Else
    aa = Form1.ShotLine.Y2 - Form1.MeaLine.Y1
End If
If Form1.MeaLine.X1 > Form1.MeaLine.X2 Then
    bb = Form1.MeaLine.X1 - Form1.ShotLine.X2
Else
    bb = Form1.MeaLine.X2 - Form1.MeaLine.X1
End If
Dim i As Integer
Form1.TotalDist.Text = Sqr(aa * aa + bb * bb)
If InStr(1, Form1.TotalDist.Text, ".", vbTextCompare) <> 0 Then
    i = InStr(1, Form1.TotalDist.Text, ".", vbTextCompare)
    Form1.TotalDist.Text = Mid(Form1.TotalDist.Text, 1, i - 1)
    End If
Form1.LeftDist.Text = Form1.TotalDist.Text
Form1.HolePar.Text = ParNum.Text
Form1.HoleNumber.Text = "Custom"
Form1.HoleScore.Text = "0"
Form1.BallLocate.Text = "Tee Off"
'Set Flag Position
Form1.Flag1.X1 = FlagX
Form1.Flag1.X2 = FlagX
Form1.Flag1.Y1 = FlagY
Form1.Flag1.Y2 = FlagY - 40

Form1.Flag2.X1 = FlagX
Form1.Flag2.X2 = FlagX + 18
Form1.Flag2.Y1 = FlagY - 40
Form1.Flag2.Y2 = FlagY - 35

Form1.Flag3.X1 = FlagX + 18
Form1.Flag3.X2 = FlagX
Form1.Flag3.Y1 = FlagY - 35
Form1.Flag3.Y2 = FlagY - 30
Form2.Visible = False

Form1.IronBox.ListIndex = 0
Form1.WoodBox.ListIndex = 0

Form1.Show
End Sub

Private Sub Command4_Click()
SelStart = True
MsgBox "Click where the Start Is!", vbExclamation, "JeroGolf"
CourseEdit.MousePointer = 0
End Sub

Private Sub Command5_Click()
SelFlag = True
MsgBox "Click where the Hole Is!", vbExclamation, "JeroGolf"
CourseEdit.MousePointer = 0
End Sub

Private Sub Command6_Click()
Dim X, Y As Integer
Dim r, g, b As Integer
CommonDialog1.Filter = "Bitmaps|*.bmp"
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then
    Exit Sub
    End If
CourseEdit.Picture = LoadPicture(CommonDialog1.FileName)
CourseEdit.ForeColor = ColorRough.BackColor
Form2.Caption = "Loading..."
StartX = 0
StartY = 0
FlagX = 0
FlagY = 0
For X = 0 To CourseEdit.ScaleWidth
    For Y = 0 To CourseEdit.ScaleHeight
        If GetPixel(CourseEdit.hdc, X, Y) = ColorRough.BackColor Or GetPixel(CourseEdit.hdc, X, Y) = ColorFairWay.BackColor Or GetPixel(CourseEdit.hdc, X, Y) = ColorGreen.BackColor Or GetPixel(CourseEdit.hdc, X, Y) = ColorSand.BackColor Or GetPixel(CourseEdit.hdc, X, Y) = ColorWater.BackColor Then
        Else
            If GetPixel(CourseEdit.hdc, X, Y) = RGB(0, 0, 0) Then
                FlagX = X - 5
                FlagY = Y - 1
                FlagPosX.Text = FlagX
                FlagPosY.Text = FlagY
            Else
                If GetPixel(CourseEdit.hdc, X, Y) = RGB(255, 255, 255) Then
                    StartX = X
                    StartY = Y
                    BallPlace.Left = X - BallPlace.Width / 2
                    BallPlace.Top = Y - BallPlace.Height / 2
                    BallPlace.Visible = True
                    CourseEdit.ForeColor = GetPixel(CourseEdit.hdc, X - 1, Y)
                    CourseEdit.PSet (X, Y)
                    StartPosX.Text = StartX
                    StartPosY.Text = StartY
                Else
                    CourseEdit.PSet (X, Y)
                End If
            End If
        End If
    Next Y
Next X
Form2.Caption = "Course Editor"
MsgBox "Done Loading and Changing Colors...", vbExclamation, "JeroGolf"
End Sub

Private Sub Command7_Click()

CourseRedo.Height = CourseEdit.Height
CourseRedo.Width = CourseEdit.Width
Call BitBlt(CourseRedo.hdc, 0, 0, CourseEdit.ScaleWidth, CourseEdit.ScaleHeight, CourseEdit.hdc, 0, 0, SRCAND)
Call BitBlt(CourseRedo.hdc, 0, 0, CourseEdit.ScaleWidth, CourseEdit.ScaleHeight, CourseEdit.hdc, 0, 0, SRCPAINT)

Call UndoMove
Command8.Enabled = True
End Sub

Public Sub UndoMove()
Dim UndoNum As Integer
On Error GoTo error
CourseEdit.Cls
UndoNum = CourseUndo.Count - 1

CourseEdit.Picture = CourseUndo(UndoNum).Image

Unload CourseUndo(UndoNum)
Undo = Undo - 1
Exit Sub
error:
MsgBox "Cannot Undo", vbCritical, "JeroGolf"
End Sub

Public Sub RedoMove()
CourseEdit.Cls
CourseEdit.Picture = CourseRedo.Image
End Sub

Private Sub Command8_Click()

Undo = Undo + 1
Load CourseUndo(Undo)
CourseUndo(Undo).Left = CourseUndo(0).Left
CourseUndo(Undo).Top = CourseUndo(Undo - 1).Top + 50

CourseUndo(Undo).Height = CourseEdit.Height
CourseUndo(Undo).Width = CourseEdit.Width
Call BitBlt(CourseUndo(Undo).hdc, 0, 0, CourseEdit.ScaleWidth, CourseEdit.ScaleHeight, CourseEdit.hdc, 0, 0, SRCAND)
Call BitBlt(CourseUndo(Undo).hdc, 0, 0, CourseEdit.ScaleWidth, CourseEdit.ScaleHeight, CourseEdit.hdc, 0, 0, SRCPAINT)

Call RedoMove
Command8.Enabled = False
End Sub

Private Sub CourseEdit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyControl Then
    ShiftDown = True
    End If
If ShiftDown = True And KeyCode = vbKeyZ Then
    Call UndoMove
    ShiftDown = False
    End If
End Sub

Private Sub CourseEdit_KeyUp(KeyCode As Integer, Shift As Integer)
ShiftDown = False
End Sub

Private Sub CourseEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command7.Enabled = True

Undo = Undo + 1
Load CourseUndo(Undo)
CourseUndo(Undo).Left = CourseUndo(0).Left
CourseUndo(Undo).Top = CourseUndo(Undo - 1).Top + 50

CourseUndo(Undo).Height = CourseEdit.Height
CourseUndo(Undo).Width = CourseEdit.Width
Call BitBlt(CourseUndo(Undo).hdc, 0, 0, CourseEdit.ScaleWidth, CourseEdit.ScaleHeight, CourseEdit.hdc, 0, 0, SRCAND)
Call BitBlt(CourseUndo(Undo).hdc, 0, 0, CourseEdit.ScaleWidth, CourseEdit.ScaleHeight, CourseEdit.hdc, 0, 0, SRCPAINT)

MD = True
CourseEdit.ForeColor = ColorSel.BackColor
If SelFlag = True Then
    FlagX = X
    FlagY = Y
    SelFlag = False
    Exit Sub
    End If
If SelStart = True Then
    StartX = X
    StartY = Y
    CourseEdit.ForeColor = GetPixel(CourseEdit.hdc, X + 1, Y + 2)
    CourseEdit.PSet (BallPlace.Left + 2, BallPlace.Top + 2)
    BallPlace.Left = X
    BallPlace.Top = Y
    CourseEdit.ForeColor = RGB(255, 255, 255)
    CourseEdit.PSet (X + 2, Y + 2)
    CourseEdit.ForeColor = ColorSel.BackColor
    BallPlace.Visible = True
    SelStart = False
    Exit Sub
    End If
If ToolSel.Picture = ToolCircle.Picture Then
    Call CircleTool(X, Y)
    End If
If ToolSel.Picture = ToolBox.Picture Then
    Call BoxTool(X, Y)
    End If
If ToolSel.Picture = ToolDot.Picture Then
    sX = X
    sY = Y
    DotX = X
    DotY = Y
    End If
If ToolSel.Picture = ToolGreen.Picture Then
    Call GreenTool(X, Y)
    End If
If ToolSel.Picture = ToolFill.Picture Then
    Call Filling(CourseEdit.Point(X, Y), 0, X, Y)
    End If
If ToolSel.Picture = ToolWater.Picture Then
    Call WaterTool(X, Y)
    End If
End Sub

Private Sub CourseEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo error
Coordinate.Caption = X & " X " & Y
If MD = True Then
    CourseEdit.ForeColor = ColorSel.BackColor
    If ToolSel.Picture = ToolCircle.Picture Then
        Call CircleTool(X, Y)
        End If
    If ToolSel.Picture = ToolBox.Picture Then
        Call BoxTool(X, Y)
        End If
    If ToolSel.Picture = ToolWater.Picture Then
        Call WaterTool(X, Y)
        End If
    If ToolSel.Picture = ToolDot.Picture Then
        CourseEdit.Line (sX, sY)-(X, Y)
        sX = X
        sY = Y
        End If
End If
error:
End Sub

Private Sub CourseEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MD = False
If AutoFinish.Value = 1 And ToolSel.Picture = ToolDot.Picture Then
    CourseEdit.Line (X, Y)-(DotX, DotY)
    End If
End Sub

Private Sub Form_Load()
SelStart = False
SelFlag = False
FlagX = 0
FlagY = 0
StartX = 0
StartY = 0
End Sub


Public Function BoxTool(X As Single, Y As Single)
Dim boxll, boxww As String
boxll = BoxL.Text / 2
If InStr(1, boxll, ".", vbTextCompare) <> 0 Then
    BoxL.Text = BoxL.Text - -1
    End If
boxww = BoxW.Text / 2
If InStr(1, boxww, ".", vbTextCompare) <> 0 Then
    BoxW.Text = BoxW.Text - -1
    End If
    
For xx = X - BoxL.Text / 2 To X + BoxL.Text / 2
    For yy = Y - BoxW.Text / 2 To Y + BoxW.Text / 2
        CourseEdit.PSet (xx, yy)
    Next yy
Next xx
End Function

Public Function CircleTool(X As Single, Y As Single)
Dim ii As Integer
For p = 1 To CircleSize.Text
    For i = 0 To 1
        If LeftOption.Value = False Then
            X = X + i
        Else
            X = X - i
        End If
        For ii = 0 To 1
            Y = Y + ii
            CourseEdit.PSet (X + 2, Y)
            CourseEdit.PSet (X + 3, Y)
            CourseEdit.PSet (X + 4, Y)
            
            CourseEdit.PSet (X + 1, Y + 1)
            CourseEdit.PSet (X + 2, Y + 1)
            CourseEdit.PSet (X + 3, Y + 1)
            CourseEdit.PSet (X + 4, Y + 1)
            CourseEdit.PSet (X + 5, Y + 1)
            
            CourseEdit.PSet (X, Y + 2)
            CourseEdit.PSet (X + 1, Y + 2)
            CourseEdit.PSet (X + 2, Y + 2)
            CourseEdit.PSet (X + 3, Y + 2)
            CourseEdit.PSet (X + 4, Y + 2)
            CourseEdit.PSet (X + 5, Y + 2)
            CourseEdit.PSet (X + 6, Y + 2)
            
            CourseEdit.PSet (X, Y + 3)
            CourseEdit.PSet (X + 1, Y + 3)
            CourseEdit.PSet (X + 2, Y + 3)
            CourseEdit.PSet (X + 3, Y + 3)
            CourseEdit.PSet (X + 4, Y + 3)
            CourseEdit.PSet (X + 5, Y + 3)
            CourseEdit.PSet (X + 6, Y + 3)
            
            CourseEdit.PSet (X, Y + 4)
            CourseEdit.PSet (X + 1, Y + 4)
            CourseEdit.PSet (X + 2, Y + 4)
            CourseEdit.PSet (X + 3, Y + 4)
            CourseEdit.PSet (X + 4, Y + 4)
            CourseEdit.PSet (X + 5, Y + 4)
            CourseEdit.PSet (X + 6, Y + 4)
            
            CourseEdit.PSet (X + 1, Y + 5)
            CourseEdit.PSet (X + 2, Y + 5)
            CourseEdit.PSet (X + 3, Y + 5)
            CourseEdit.PSet (X + 4, Y + 5)
            CourseEdit.PSet (X + 5, Y + 5)
            
            CourseEdit.PSet (X + 2, Y + 6)
            CourseEdit.PSet (X + 3, Y + 6)
            CourseEdit.PSet (X + 4, Y + 6)
        Next ii
    Next i
Next p
End Function

Public Function WaterTool(X As Single, Y As Single)
Dim ii As Integer
For p = 1 To WaterSize.Text
    For i = 0 To WaterSize2.Text
        If LeftOption2.Value = False Then
            X = X + i
        Else
            X = X - i
        End If
        For ii = 0 To 1
            Y = Y + ii
            CourseEdit.PSet (X + 2, Y)
            CourseEdit.PSet (X + 3, Y)
            CourseEdit.PSet (X + 4, Y)
            
            CourseEdit.PSet (X + 1, Y + 1)
            CourseEdit.PSet (X + 2, Y + 1)
            CourseEdit.PSet (X + 3, Y + 1)
            CourseEdit.PSet (X + 4, Y + 1)
            CourseEdit.PSet (X + 5, Y + 1)
            
            CourseEdit.PSet (X, Y + 2)
            CourseEdit.PSet (X + 1, Y + 2)
            CourseEdit.PSet (X + 2, Y + 2)
            CourseEdit.PSet (X + 3, Y + 2)
            CourseEdit.PSet (X + 4, Y + 2)
            CourseEdit.PSet (X + 5, Y + 2)
            CourseEdit.PSet (X + 6, Y + 2)
            
            CourseEdit.PSet (X, Y + 3)
            CourseEdit.PSet (X + 1, Y + 3)
            CourseEdit.PSet (X + 2, Y + 3)
            CourseEdit.PSet (X + 3, Y + 3)
            CourseEdit.PSet (X + 4, Y + 3)
            CourseEdit.PSet (X + 5, Y + 3)
            CourseEdit.PSet (X + 6, Y + 3)
            
            CourseEdit.PSet (X, Y + 4)
            CourseEdit.PSet (X + 1, Y + 4)
            CourseEdit.PSet (X + 2, Y + 4)
            CourseEdit.PSet (X + 3, Y + 4)
            CourseEdit.PSet (X + 4, Y + 4)
            CourseEdit.PSet (X + 5, Y + 4)
            CourseEdit.PSet (X + 6, Y + 4)
            
            CourseEdit.PSet (X + 1, Y + 5)
            CourseEdit.PSet (X + 2, Y + 5)
            CourseEdit.PSet (X + 3, Y + 5)
            CourseEdit.PSet (X + 4, Y + 5)
            CourseEdit.PSet (X + 5, Y + 5)
            
            CourseEdit.PSet (X + 2, Y + 6)
            CourseEdit.PSet (X + 3, Y + 6)
            CourseEdit.PSet (X + 4, Y + 6)
        Next ii
    Next i
Next p
End Function

Public Function GreenTool(X As Single, Y As Single)
CourseEdit.ForeColor = RGB(0, 0, 0)
'Do the Hole Drawing
For xx = X To X + 10
    For yy = Y To Y + 10
        CourseEdit.PSet (xx, yy)
    Next yy
Next xx
CourseEdit.ForeColor = ColorGreen.BackColor
'Do the Green Drawing
CourseEdit.PSet (X, Y)
CourseEdit.PSet (X + 1, Y)
CourseEdit.PSet (X, Y + 1)
CourseEdit.PSet (X + 1, Y + 1)
CourseEdit.PSet (X + 2, Y)
CourseEdit.PSet (X, Y + 2)

CourseEdit.PSet (X + 10, Y + 10)
CourseEdit.PSet (X + 9, Y + 10)
CourseEdit.PSet (X + 10, Y + 9)
CourseEdit.PSet (X + 9, Y + 9)
CourseEdit.PSet (X + 10, Y + 8)
CourseEdit.PSet (X + 8, Y + 10)

CourseEdit.PSet (X + 8, Y)
CourseEdit.PSet (X + 9, Y)
CourseEdit.PSet (X + 10, Y)
CourseEdit.PSet (X + 9, Y + 1)
CourseEdit.PSet (X + 10, Y + 1)
CourseEdit.PSet (X + 10, Y + 2)

CourseEdit.PSet (X, Y + 8)
CourseEdit.PSet (X, Y + 9)
CourseEdit.PSet (X, Y + 10)
CourseEdit.PSet (X + 1, Y + 9)
CourseEdit.PSet (X + 1, Y + 10)
CourseEdit.PSet (X + 2, Y + 10)
End Function

