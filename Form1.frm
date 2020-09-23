VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "JeroGolf"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   573
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      Caption         =   "Options"
      Height          =   7095
      Left            =   8760
      TabIndex        =   24
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton Command1 
         Caption         =   "Custom Hole"
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   3120
         Width           =   1455
      End
      Begin VB.PictureBox Course5 
         Height          =   375
         Left            =   1320
         Picture         =   "Form1.frx":030A
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   53
         Top             =   2640
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Course5but 
         Caption         =   "Hole #5"
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   2640
         Width           =   1095
      End
      Begin VB.PictureBox Sample 
         BackColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   48
         Top             =   6000
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox TypeRough 
         AutoRedraw      =   -1  'True
         Height          =   350
         Left            =   480
         Picture         =   "Form1.frx":8D60C
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   47
         Top             =   6000
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.PictureBox TypeFairWay 
         AutoRedraw      =   -1  'True
         Height          =   350
         Left            =   840
         Picture         =   "Form1.frx":8F38E
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   46
         Top             =   6000
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.PictureBox TypeSand 
         AutoRedraw      =   -1  'True
         Height          =   350
         Left            =   840
         Picture         =   "Form1.frx":8FA24
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   45
         Top             =   6360
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.PictureBox TypeWater 
         AutoRedraw      =   -1  'True
         Height          =   350
         Left            =   1200
         Picture         =   "Form1.frx":9021A
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   44
         Top             =   6000
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.PictureBox TypeGreen 
         AutoRedraw      =   -1  'True
         Height          =   350
         Left            =   480
         Picture         =   "Form1.frx":9214C
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   43
         Top             =   6360
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.PictureBox SampleComp 
         Height          =   375
         Left            =   120
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   42
         Top             =   6360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox TypeHole 
         AutoRedraw      =   -1  'True
         Height          =   350
         Left            =   1200
         Picture         =   "Form1.frx":93202
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   19
         TabIndex        =   41
         Top             =   6360
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.PictureBox Course4 
         Height          =   375
         Left            =   1320
         Picture         =   "Form1.frx":949C8
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   40
         Top             =   2280
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Course4but 
         Caption         =   "Hole #4"
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton Course3but 
         Caption         =   "Hole #3"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   1920
         Width           =   1095
      End
      Begin VB.PictureBox Course3 
         Height          =   375
         Left            =   1320
         Picture         =   "Form1.frx":121CCA
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   37
         Top             =   1920
         Visible         =   0   'False
         Width           =   1575
         Begin VB.Line Line1 
            X1              =   160
            X2              =   160
            Y1              =   112
            Y2              =   80
         End
      End
      Begin VB.PictureBox Course1 
         Height          =   375
         Left            =   1320
         Picture         =   "Form1.frx":1AEFCC
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   35
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox Course2 
         Height          =   375
         Left            =   1320
         Picture         =   "Form1.frx":23C2CE
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   34
         Top             =   1560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Course2but 
         Caption         =   "Hole #2"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox PinCheck 
         Caption         =   "Pin Flag Visible"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CommandButton Course1but 
         Caption         =   "Hole #1"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Line Separate1 
         X1              =   120
         X2              =   2880
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label12 
         Caption         =   "Select Course:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Wind"
      Height          =   1695
      Left            =   5400
      TabIndex        =   19
      Top             =   120
      Width           =   3255
      Begin VB.PictureBox WindPic 
         BackColor       =   &H00C0C0C0&
         Height          =   1005
         Left            =   1560
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   23
         Top             =   600
         Width           =   1005
      End
      Begin VB.TextBox WindSpeed 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   21
         Text            =   "0"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Direction:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Wind Speed:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.Image WindDown 
         Height          =   255
         Left            =   1080
         Picture         =   "Form1.frx":2C95D0
         Stretch         =   -1  'True
         Top             =   840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image WindUp 
         Height          =   240
         Left            =   840
         Picture         =   "Form1.frx":2CC612
         Stretch         =   -1  'True
         Top             =   840
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image WindLeft 
         Height          =   255
         Left            =   480
         Picture         =   "Form1.frx":2CF654
         Stretch         =   -1  'True
         Top             =   840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image WindRight 
         Height          =   240
         Left            =   240
         Picture         =   "Form1.frx":2D2696
         Stretch         =   -1  'True
         Top             =   840
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stats"
      Height          =   3255
      Left            =   5400
      TabIndex        =   5
      Top             =   1920
      Width           =   3255
      Begin VB.TextBox TotalDist 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   18
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox LeftDist 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox HoleNumber 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Text            =   "1"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox HolePar 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Text            =   "3"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox HoleScore 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Text            =   "0"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox BallLocate 
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Text            =   "Tee Off"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Distance:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Distance Left:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Hole:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Par:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Strokes:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Location:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.PictureBox CoursePic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   8400
      Left            =   120
      Picture         =   "Form1.frx":2D56D8
      ScaleHeight     =   556
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   346
      TabIndex        =   0
      Top             =   120
      Width           =   5250
      Begin VB.Line Flag3 
         BorderColor     =   &H000000FF&
         X1              =   168
         X2              =   150
         Y1              =   136
         Y2              =   142
      End
      Begin VB.Line Flag2 
         BorderColor     =   &H000000FF&
         X1              =   150
         X2              =   168
         Y1              =   130
         Y2              =   135
      End
      Begin VB.Line Flag1 
         BorderColor     =   &H000000FF&
         X1              =   150
         X2              =   150
         Y1              =   168
         Y2              =   128
      End
      Begin VB.Line MeaLine 
         Visible         =   0   'False
         X1              =   32
         X2              =   152
         Y1              =   80
         Y2              =   168
      End
      Begin VB.Line ShotLine 
         BorderColor     =   &H000080FF&
         X1              =   8
         X2              =   8
         Y1              =   72
         Y2              =   8
      End
      Begin VB.Shape Ball 
         BackStyle       =   1  'Opaque
         Height          =   105
         Left            =   2400
         Shape           =   3  'Circle
         Top             =   7560
         Width           =   105
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Ball Color"
      Height          =   1215
      Left            =   8760
      TabIndex        =   50
      Top             =   7320
      Width           =   3015
      Begin VB.PictureBox BallCol 
         AutoRedraw      =   -1  'True
         Height          =   375
         Left            =   2520
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   21
         TabIndex        =   51
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image BallColor 
         Height          =   270
         Index           =   9
         Left            =   840
         Picture         =   "Form1.frx":3629DA
         Top             =   720
         Width           =   270
      End
      Begin VB.Image BallColor 
         Height          =   270
         Index           =   8
         Left            =   1320
         Picture         =   "Form1.frx":362E0C
         Top             =   720
         Width           =   270
      End
      Begin VB.Image BallColor 
         Height          =   270
         Index           =   7
         Left            =   1320
         Picture         =   "Form1.frx":36323E
         Top             =   360
         Width           =   270
      End
      Begin VB.Image BallColor 
         Height          =   270
         Index           =   6
         Left            =   360
         Picture         =   "Form1.frx":363670
         Top             =   720
         Width           =   270
      End
      Begin VB.Image BallColor 
         Height          =   270
         Index           =   5
         Left            =   1800
         Picture         =   "Form1.frx":363AA2
         Top             =   720
         Width           =   270
      End
      Begin VB.Image BallColor 
         Height          =   270
         Index           =   4
         Left            =   360
         Picture         =   "Form1.frx":363ED4
         Top             =   360
         Width           =   270
      End
      Begin VB.Image BallColor 
         Height          =   270
         Index           =   3
         Left            =   840
         Picture         =   "Form1.frx":364306
         Top             =   360
         Width           =   270
      End
      Begin VB.Image BallColor 
         Height          =   270
         Index           =   2
         Left            =   2280
         Picture         =   "Form1.frx":364738
         Top             =   360
         Width           =   270
      End
      Begin VB.Image BallColor 
         Height          =   270
         Index           =   1
         Left            =   2280
         Picture         =   "Form1.frx":364B6A
         Top             =   720
         Width           =   270
      End
      Begin VB.Image BallColor 
         Height          =   270
         Index           =   0
         Left            =   1800
         Picture         =   "Form1.frx":364F9C
         Top             =   360
         Width           =   270
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Club Selection"
      Height          =   1575
      Left            =   5400
      TabIndex        =   14
      Top             =   5280
      Width           =   3255
      Begin VB.TextBox Club 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   30
         Top             =   1080
         Width           =   1695
      End
      Begin VB.ComboBox IronBox 
         Height          =   315
         Left            =   1320
         TabIndex        =   29
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox WoodBox 
         Height          =   315
         ItemData        =   "Form1.frx":3653CE
         Left            =   1320
         List            =   "Form1.frx":3653D0
         TabIndex        =   26
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Club Choice:"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Irons:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Woods:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Golf Shot"
      Height          =   1575
      Left            =   5400
      TabIndex        =   1
      Top             =   6960
      Width           =   3255
      Begin VB.CommandButton SwingBut2 
         Height          =   375
         Left            =   2040
         TabIndex        =   54
         Top             =   1080
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Timer BallMove 
         Enabled         =   0   'False
         Interval        =   5
         Left            =   1800
         Top             =   240
      End
      Begin VB.Timer PowerTimer 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1320
         Top             =   240
      End
      Begin VB.CommandButton PowerMeterBut 
         Caption         =   "Start"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   25
      End
      Begin VB.CommandButton SwingBut 
         Caption         =   "Swing"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Image Image3 
         Height          =   375
         Left            =   1560
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   2760
         Picture         =   "Form1.frx":3653D2
         Stretch         =   -1  'True
         Top             =   360
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   2760
         Picture         =   "Form1.frx":366014
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   240
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim way As String
Dim stuck As Integer
Dim TotalPower As Integer
Dim LineLen2 As Integer
Dim d, e As Integer
Dim a, b, c As Integer
Dim aa, bb As Integer
Dim xx, yy As Integer
Dim xxx, yyy As Integer
Dim j, jj As Integer
Dim k As Integer

Public Sub FetchClubs()
WoodBox.AddItem "1 Wood"
WoodBox.AddItem "3 Wood"
WoodBox.AddItem "5 Wood"
IronBox.AddItem "3 Iron"
IronBox.AddItem "4 Iron"
IronBox.AddItem "5 Iron"
IronBox.AddItem "6 Iron"
IronBox.AddItem "7 Iron"
IronBox.AddItem "8 Iron"
IronBox.AddItem "9 Iron"
IronBox.AddItem "Putter"
IronBox.ListIndex = 0
WoodBox.ListIndex = 0
End Sub

Private Sub BallColor_Click(Index As Integer)
BallCol.Picture = BallColor(Index).Picture
Ball.BackColor = GetPixel(BallCol.hdc, 9, 9)
End Sub

Private Sub BallMove_Timer()
d = ShotLine.X1 - ShotLine.X2
e = ShotLine.Y1 - ShotLine.Y2
If ShotLine.Y1 > ShotLine.Y2 Then
    a = ShotLine.Y1 - ShotLine.Y2
Else
    a = ShotLine.Y2 - ShotLine.Y1
End If
If ShotLine.X1 > ShotLine.X2 Then
    b = ShotLine.X1 - ShotLine.X2
Else
    b = ShotLine.X2 - ShotLine.X1
End If
c = Sqr(a * a + b * b)
    
If BallLocate.Text = "Green" Then
    GoTo NoSizeIncrease
    End If
    
If c <= k / 2 Then
    If Ball.Width = 7 Then
    Else
        Ball.Width = Ball.Width - 1
        Ball.Height = Ball.Height - 1
    End If
Else
    Ball.Width = Ball.Width + 1
    Ball.Height = Ball.Height + 1
End If
    
NoSizeIncrease:
ShotLine.X1 = ShotLine.X2 + d / 1.05
ShotLine.Y1 = ShotLine.Y2 + e / 1.05
    
Ball.Left = ShotLine.X1 - Ball.Width / 2
Ball.Top = ShotLine.Y1 - Ball.Height / 2

If a < Ball.Height * 1.5 Then
    Ball.Left = ShotLine.X2 - Ball.Width / 2
    Ball.Top = ShotLine.Y2 - Ball.Height / 2
    SwingBut2.Value = True
    Ball.Height = 7
    Ball.Width = 7
    BallMove.Enabled = False
    End If
End Sub

Private Sub Command1_Click()
Form2.ColorRough.BackColor = GetPixel(TypeRough.hdc, 0, 0)
Form2.ColorFairWay.BackColor = GetPixel(TypeFairWay.hdc, 0, 0)
Form2.ColorGreen.BackColor = GetPixel(TypeGreen.hdc, 0, 0)
Form2.ColorSand.BackColor = GetPixel(TypeSand.hdc, 0, 0)
Form2.ColorWater.BackColor = GetPixel(TypeWater.hdc, 0, 0)
Form2.ColorSel.BackColor = Form2.ColorRough.BackColor
Form2.ColorSel2.BackColor = Form2.ColorRough.BackColor
Form2.CmdDot.Value = True
Form2.Show
Form2.CourseUndo(0).Height = Form2.CourseEdit.Height
Form2.CourseUndo(0).Width = Form2.CourseEdit.Width
Call BitBlt(Form2.CourseUndo(0).hdc, 0, 0, Form2.CourseEdit.ScaleWidth, Form2.CourseEdit.ScaleHeight, Form2.CourseEdit.hdc, 0, 0, SRCAND)
Call BitBlt(Form2.CourseUndo(0).hdc, 0, 0, Form2.CourseEdit.ScaleWidth, Form2.CourseEdit.ScaleHeight, Form2.CourseEdit.hdc, 0, 0, SRCPAINT)
End Sub

Private Sub Course5but_Click()
Dim ran As Integer
Ball.Left = 173
Ball.Top = 525
CoursePic.Picture = Course5.Picture
WindSpeed.Text = Int(Rnd * 16)
ran = Int(Rnd * 4)
If ran = 0 Then
    WindPic.Picture = WindLeft.Picture
    End If
If ran = 1 Then
    WindPic.Picture = WindRight.Picture
    End If
If ran = 2 Then
    WindPic.Picture = WindUp.Picture
    End If
If ran = 3 Then
    WindPic.Picture = WindDown.Picture
    End If
ShotLine.X1 = Ball.Left + Ball.Width / 2
ShotLine.X2 = ShotLine.X1

ShotLine.Y1 = Ball.Top + Ball.Height / 2
ShotLine.Y2 = Ball.Top - 50

Dim aa, bb As Integer
MeaLine.X1 = Ball.Left
MeaLine.Y1 = Ball.Top
If MeaLine.Y1 > MeaLine.Y2 Then
    aa = MeaLine.Y1 - MeaLine.Y2
Else
    aa = ShotLine.Y2 - MeaLine.Y1
End If
If MeaLine.X1 > MeaLine.X2 Then
    bb = MeaLine.X1 - ShotLine.X2
Else
    bb = MeaLine.X2 - MeaLine.X1
End If
Dim i As Integer
TotalDist.Text = Sqr(aa * aa + bb * bb)
If InStr(1, TotalDist.Text, ".", vbTextCompare) <> 0 Then
    i = InStr(1, TotalDist.Text, ".", vbTextCompare)
    TotalDist.Text = Mid(TotalDist.Text, 1, i - 1)
    End If
LeftDist.Text = TotalDist.Text
HolePar.Text = "4"
HoleNumber.Text = "5"
HoleScore.Text = "0"
BallLocate.Text = "Tee Off"
'Set Flag Position
Flag1.X1 = 165
Flag1.X2 = 165
Flag1.Y1 = 55
Flag1.Y2 = 15

Flag2.X1 = 165
Flag2.X2 = 165 + 15
Flag2.Y1 = 15
Flag2.Y2 = 20

Flag3.X1 = 165 + 15
Flag3.X2 = 165
Flag3.Y1 = 20
Flag3.Y2 = 25

IronBox.ListIndex = 0
WoodBox.ListIndex = 0
End Sub

Private Sub Course2but_Click()
Dim ran As Integer
Ball.Left = 90
Ball.Top = 500
CoursePic.Picture = Course2.Picture
WindSpeed.Text = Int(Rnd * 16)
ran = Int(Rnd * 4)
If ran = 0 Then
    WindPic.Picture = WindLeft.Picture
    End If
If ran = 1 Then
    WindPic.Picture = WindRight.Picture
    End If
If ran = 2 Then
    WindPic.Picture = WindUp.Picture
    End If
If ran = 3 Then
    WindPic.Picture = WindDown.Picture
    End If
ShotLine.X1 = Ball.Left + Ball.Width / 2
ShotLine.X2 = ShotLine.X1

ShotLine.Y1 = Ball.Top + Ball.Height / 2
ShotLine.Y2 = Ball.Top - 50

Dim aa, bb As Integer
MeaLine.X1 = Ball.Left
MeaLine.Y1 = Ball.Top
If MeaLine.Y1 > MeaLine.Y2 Then
    aa = MeaLine.Y1 - MeaLine.Y2
Else
    aa = ShotLine.Y2 - MeaLine.Y1
End If
If MeaLine.X1 > MeaLine.X2 Then
    bb = MeaLine.X1 - ShotLine.X2
Else
    bb = MeaLine.X2 - MeaLine.X1
End If
Dim i As Integer
TotalDist.Text = Sqr(aa * aa + bb * bb)
If InStr(1, TotalDist.Text, ".", vbTextCompare) <> 0 Then
    i = InStr(1, TotalDist.Text, ".", vbTextCompare)
    TotalDist.Text = Mid(TotalDist.Text, 1, i - 1)
    End If
LeftDist.Text = TotalDist.Text
HolePar.Text = "5"
HoleNumber.Text = "2"
HoleScore.Text = "0"
BallLocate.Text = "Tee Off"
'Set Flag Position
Flag1.X1 = 53
Flag1.X2 = 53
Flag1.Y1 = 0
Flag1.Y2 = 40

Flag2.X1 = 53
Flag2.X2 = 53 + 18
Flag2.Y1 = 0
Flag2.Y2 = 0 + 5

Flag3.X1 = 53 + 18
Flag3.X2 = 53
Flag3.Y1 = 0 + 5
Flag3.Y2 = 10

IronBox.ListIndex = 0
WoodBox.ListIndex = 0
End Sub

Private Sub Course1but_Click()
Dim ran As Integer
Ball.Left = 160
Ball.Top = 504
CoursePic.Picture = Course1.Picture
WindSpeed.Text = Int(Rnd * 16)
ran = Int(Rnd * 4)
If ran = 0 Then
    WindPic.Picture = WindLeft.Picture
    End If
If ran = 1 Then
    WindPic.Picture = WindRight.Picture
    End If
If ran = 2 Then
    WindPic.Picture = WindUp.Picture
    End If
If ran = 3 Then
    WindPic.Picture = WindDown.Picture
    End If
ShotLine.X1 = Ball.Left + Ball.Width / 2
ShotLine.X2 = ShotLine.X1

ShotLine.Y1 = Ball.Top + Ball.Height / 2
ShotLine.Y2 = Ball.Top - 50

Dim aa, bb As Integer
MeaLine.X1 = Ball.Left
MeaLine.Y1 = Ball.Top
If MeaLine.Y1 > MeaLine.Y2 Then
    aa = MeaLine.Y1 - MeaLine.Y2
Else
    aa = ShotLine.Y2 - MeaLine.Y1
End If
If MeaLine.X1 > MeaLine.X2 Then
    bb = MeaLine.X1 - ShotLine.X2
Else
    bb = MeaLine.X2 - MeaLine.X1
End If
Dim i As Integer
TotalDist.Text = Sqr(aa * aa + bb * bb)
If InStr(1, TotalDist.Text, ".", vbTextCompare) <> 0 Then
    i = InStr(1, TotalDist.Text, ".", vbTextCompare)
    TotalDist.Text = Mid(TotalDist.Text, 1, i - 1)
    End If
LeftDist.Text = TotalDist.Text
HolePar.Text = "3"
HoleNumber.Text = "2"
HoleScore.Text = "0"
BallLocate.Text = "Tee Off"
'Set Flag Position
Flag1.X1 = 150
Flag1.X2 = 150
Flag1.Y1 = 168
Flag1.Y2 = 128

Flag2.X1 = 150
Flag2.X2 = 168
Flag2.Y1 = 130
Flag2.Y2 = 135

Flag3.X1 = 168
Flag3.X2 = 150
Flag3.Y1 = 136
Flag3.Y2 = 142

IronBox.ListIndex = 0
WoodBox.ListIndex = 0
End Sub

Private Sub Course3but_Click()
Dim ran As Integer
Ball.Left = 215
Ball.Top = 505
CoursePic.Picture = Course3.Picture
WindSpeed.Text = Int(Rnd * 16)
ran = Int(Rnd * 4)
If ran = 0 Then
    WindPic.Picture = WindLeft.Picture
    End If
If ran = 1 Then
    WindPic.Picture = WindRight.Picture
    End If
If ran = 2 Then
    WindPic.Picture = WindUp.Picture
    End If
If ran = 3 Then
    WindPic.Picture = WindDown.Picture
    End If
ShotLine.X1 = Ball.Left + Ball.Width / 2
ShotLine.X2 = ShotLine.X1

ShotLine.Y1 = Ball.Top + Ball.Height / 2
ShotLine.Y2 = Ball.Top - 50

Dim aa, bb As Integer
MeaLine.X1 = Ball.Left
MeaLine.Y1 = Ball.Top
If MeaLine.Y1 > MeaLine.Y2 Then
    aa = MeaLine.Y1 - MeaLine.Y2
Else
    aa = ShotLine.Y2 - MeaLine.Y1
End If
If MeaLine.X1 > MeaLine.X2 Then
    bb = MeaLine.X1 - ShotLine.X2
Else
    bb = MeaLine.X2 - MeaLine.X1
End If
Dim i As Integer
TotalDist.Text = Sqr(aa * aa + bb * bb)
If InStr(1, TotalDist.Text, ".", vbTextCompare) <> 0 Then
    i = InStr(1, TotalDist.Text, ".", vbTextCompare)
    TotalDist.Text = Mid(TotalDist.Text, 1, i - 1)
    End If
LeftDist.Text = TotalDist.Text
HolePar.Text = "4"
HoleNumber.Text = "3"
HoleScore.Text = "0"
BallLocate.Text = "Tee Off"
'Set Flag Position
Flag1.X1 = 163
Flag1.X2 = 163
Flag1.Y1 = 113
Flag1.Y2 = 73

Flag2.X1 = 163
Flag2.X2 = 163 + 15
Flag2.Y1 = 73
Flag2.Y2 = 78

Flag3.X1 = 163 + 15
Flag3.X2 = 163
Flag3.Y1 = 78
Flag3.Y2 = 83

IronBox.ListIndex = 0
WoodBox.ListIndex = 0
End Sub

Private Sub Course4but_Click()
Dim ran As Integer
Ball.Left = 155
Ball.Top = 540
CoursePic.Picture = Course4.Picture
WindSpeed.Text = Int(Rnd * 16)
ran = Int(Rnd * 4)
If ran = 0 Then
    WindPic.Picture = WindLeft.Picture
    End If
If ran = 1 Then
    WindPic.Picture = WindRight.Picture
    End If
If ran = 2 Then
    WindPic.Picture = WindUp.Picture
    End If
If ran = 3 Then
    WindPic.Picture = WindDown.Picture
    End If
ShotLine.X1 = Ball.Left + Ball.Width / 2
ShotLine.X2 = ShotLine.X1

ShotLine.Y1 = Ball.Top + Ball.Height / 2
ShotLine.Y2 = Ball.Top - 50

Dim aa, bb As Integer
MeaLine.X1 = Ball.Left
MeaLine.Y1 = Ball.Top
If MeaLine.Y1 > MeaLine.Y2 Then
    aa = MeaLine.Y1 - MeaLine.Y2
Else
    aa = ShotLine.Y2 - MeaLine.Y1
End If
If MeaLine.X1 > MeaLine.X2 Then
    bb = MeaLine.X1 - ShotLine.X2
Else
    bb = MeaLine.X2 - MeaLine.X1
End If
Dim i As Integer
TotalDist.Text = Sqr(aa * aa + bb * bb)
If InStr(1, TotalDist.Text, ".", vbTextCompare) <> 0 Then
    i = InStr(1, TotalDist.Text, ".", vbTextCompare)
    TotalDist.Text = Mid(TotalDist.Text, 1, i - 1)
    End If
LeftDist.Text = TotalDist.Text
HolePar.Text = "5"
HoleNumber.Text = "4"
HoleScore.Text = "0"
BallLocate.Text = "Tee Off"
'Set Flag Position
Flag1.X1 = 199
Flag1.X2 = 199
Flag1.Y1 = 103
Flag1.Y2 = 63

Flag2.X1 = 199
Flag2.X2 = 199 + 15
Flag2.Y1 = 63
Flag2.Y2 = 68

Flag3.X1 = 199 + 15
Flag3.X2 = 199
Flag3.Y1 = 68
Flag3.Y2 = 73

IronBox.ListIndex = 0
WoodBox.ListIndex = 0
End Sub

Private Sub Form_Resize()
If Form1.WindowState = vbMaximized Then
    Form1.WindowState = vbNormal
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form1
Unload Form2
End Sub

Private Sub Image3_Click()
PowerTimer.Interval = PowerTimer.Interval + 3
BallMove.Interval = BallMove.Interval + 3
Image3.Visible = False
End Sub

Private Sub IronBox_Change()
IronBox.ListIndex = 0
End Sub

Private Sub IronBox_Click()
Club.Text = IronBox.Text
End Sub

Private Sub PinCheck_Click()
If PinCheck.Value = 1 Then
    Flag1.Visible = True
    Flag2.Visible = True
    Flag3.Visible = True
Else
    Flag1.Visible = False
    Flag2.Visible = False
    Flag3.Visible = False
End If
End Sub

Private Sub CoursePic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim LineLen As Integer
Dim a, b, c As Integer
ShotLine.X2 = X
ShotLine.Y2 = Y
If ShotLine.Y1 > ShotLine.Y2 Then
    a = ShotLine.Y1 - ShotLine.Y2
Else
    a = ShotLine.Y2 - ShotLine.Y1
End If
If ShotLine.X1 > ShotLine.X2 Then
    b = ShotLine.X1 - ShotLine.X2
Else
    b = ShotLine.X2 - ShotLine.X1
End If
c = Sqr(a * a + b * b)
LineLen = c
End Sub

Public Sub GetType()
 
SampleComp.BackColor = GetPixel(TypeRough.hdc, 0, 0)
If Sample.BackColor = SampleComp.BackColor Then
    BallLocate.Text = "Rough"
    Exit Sub
    End If
SampleComp.BackColor = GetPixel(TypeFairWay.hdc, 0, 0)
If Sample.BackColor = SampleComp.BackColor Then
    BallLocate.Text = "FairWay"
    Exit Sub
    End If
SampleComp.BackColor = GetPixel(TypeGreen.hdc, 0, 0)
If Sample.BackColor = SampleComp.BackColor Then
    BallLocate.Text = "Green"
    Exit Sub
    End If
SampleComp.BackColor = GetPixel(TypeSand.hdc, 0, 0)
If Sample.BackColor = SampleComp.BackColor Then
    BallLocate.Text = "Sand"
    Exit Sub
    End If
SampleComp.BackColor = GetPixel(TypeWater.hdc, 0, 0)
If Sample.BackColor = SampleComp.BackColor Then
    BallLocate.Text = "Water"
    Exit Sub
    End If
SampleComp.BackColor = GetPixel(TypeHole.hdc, 0, 0)
If Sample.BackColor = SampleComp.BackColor Then
    BallLocate.Text = "In the Hole"
    Exit Sub
    End If
End Sub

Private Sub Form_Load()
Randomize
Dim ran As Integer
WindSpeed.Text = Int(Rnd * 16)
ran = Int(Rnd * 4)
If ran = 0 Then
    WindPic.Picture = WindLeft.Picture
    End If
If ran = 1 Then
    WindPic.Picture = WindRight.Picture
    End If
If ran = 2 Then
    WindPic.Picture = WindUp.Picture
    End If
If ran = 3 Then
    WindPic.Picture = WindDown.Picture
    End If
ShotLine.X1 = Ball.Left + Ball.Width / 2
ShotLine.X2 = ShotLine.X1

ShotLine.Y1 = Ball.Top + Ball.Height / 2
ShotLine.Y2 = Ball.Top - 50

Dim aa, bb As Integer
MeaLine.X1 = Ball.Left
MeaLine.Y1 = Ball.Top
If MeaLine.Y1 > MeaLine.Y2 Then
    aa = MeaLine.Y1 - MeaLine.Y2
Else
    aa = ShotLine.Y2 - MeaLine.Y1
End If
If MeaLine.X1 > MeaLine.X2 Then
    bb = MeaLine.X1 - ShotLine.X2
Else
    bb = MeaLine.X2 - MeaLine.X1
End If
Dim i As Integer
TotalDist.Text = Sqr(aa * aa + bb * bb)
If InStr(1, TotalDist.Text, ".", vbTextCompare) <> 0 Then
    i = InStr(1, TotalDist.Text, ".", vbTextCompare)
    TotalDist.Text = Mid(TotalDist.Text, 1, i - 1)
    End If
LeftDist.Text = TotalDist.Text
Call FetchClubs
End Sub

Private Sub PowerTimer_Timer()
If way = "down" Then
    If ProgressBar1.Value = 0 Then
        way = "up"
        PowerTimer.Enabled = False
        PowerMeterBut.Caption = "Start"
        Exit Sub
        End If
    ProgressBar1.Value = ProgressBar1.Value - 1
Else
    If ProgressBar1.Value = 24 Then
        way = "down"
        End If
    ProgressBar1.Value = ProgressBar1.Value + 1
End If
End Sub

Private Sub PowerMeterBut_Click()
If PowerMeterBut.Caption = "Start" Then
    PowerTimer.Enabled = True
    PowerMeterBut.Caption = "Stop"
Else
    SwingBut.Enabled = True
    PowerTimer.Enabled = False
    PowerMeterBut.Enabled = False
End If
End Sub

Private Sub SwingBut_Click()
On Error GoTo error
stuck = 0

xxx = Ball.Left
yyy = Ball.Top

If ProgressBar1.Value < 15 Then
    ProgressBar1.Value = 15
    End If

'Configure the Power of the Shot
If Club.Text = "" Then
    MsgBox "No Club Selected", vbCritical, "JeroGolf"
    Exit Sub
    End If
    
'Woods
If Club.Text = "1 Wood" Then
    If BallLocate.Text = "Tee Off" Then
        TotalPower = ProgressBar1.Value * 11
    Else
        TotalPower = ProgressBar1.Value * 10
    End If
    End If
If Club.Text = "3 Wood" Then
    If BallLocate.Text = "Tee Off" Then
        TotalPower = ProgressBar1.Value * 10
    Else
        TotalPower = ProgressBar1.Value * 9
    End If
    End If
If Club.Text = "5 Wood" Then
    If BallLocate.Text = "Tee Off" Then
        TotalPower = ProgressBar1.Value * 9
    Else
        TotalPower = ProgressBar1.Value * 8
    End If
    End If

'Irons
If Club.Text = "3 Iron" Then
    If BallLocate.Text = "Tee Off" Then
        TotalPower = ProgressBar1.Value * 8
    Else
        TotalPower = ProgressBar1.Value * 9
    End If
    End If
If Club.Text = "4 Iron" Then
    If BallLocate.Text = "Tee Off" Then
        TotalPower = ProgressBar1.Value * 7.75
    Else
        TotalPower = ProgressBar1.Value * 8.75
    End If
    End If
If Club.Text = "5 Iron" Then
    If BallLocate.Text = "Tee Off" Then
        TotalPower = ProgressBar1.Value * 7.5
    Else
        TotalPower = ProgressBar1.Value * 8.5
    End If
    End If
If Club.Text = "6 Iron" Then
    If BallLocate.Text = "Tee Off" Then
        TotalPower = ProgressBar1.Value * 7.25
    Else
        TotalPower = ProgressBar1.Value * 8.25
    End If
    End If
If Club.Text = "7 Iron" Then
    If BallLocate.Text = "Tee Off" Then
        TotalPower = ProgressBar1.Value * 7
    Else
        TotalPower = ProgressBar1.Value * 8
    End If
    End If
If Club.Text = "8 Iron" Then
    If BallLocate.Text = "Tee Off" Then
        TotalPower = ProgressBar1.Value * 6.5
    Else
        TotalPower = ProgressBar1.Value * 7.5
    End If
    End If
If Club.Text = "9 Iron" Then
    If BallLocate.Text = "Tee Off" Then
        TotalPower = ProgressBar1.Value * 6.25
    Else
        TotalPower = ProgressBar1.Value * 6.25
    End If
    End If
If Club.Text = "Putter" Then
    If BallLocate.Text = "Tee Off" Then
        TotalPower = ProgressBar1.Value * 9
        MsgBox "Putter is Broken", vbOKOnly, "JeroGolf"
        IronBox.RemoveItem IronBox.ListCount - 1
        IronBox.ListIndex = IronBox.ListCount - 1
    Else
        TotalPower = ProgressBar1.Value * 3
    End If
    End If

ShotLine.Visible = False

xx = Ball.Left + Ball.Width / 2
yy = Ball.Top + Ball.Height / 2
Sample.BackColor = GetPixel(CoursePic.hdc, xx, yy)
Call GetType

If BallLocate.Text = "Sand" Then
    TotalPower = TotalPower / 1.5
    End If
    
If BallLocate.Text = "Rough" Then
    TotalPower = TotalPower / 1.3
    End If
    
If BallLocate.Text = "Green" Then
    'Wind doesn't effect Putt
    GoTo Swing:
    End If
    
If WindPic.Picture = WindLeft.Picture Then
    ShotLine.X2 = ShotLine.X2 - WindSpeed.Text * 2
    End If
If WindPic.Picture = WindRight.Picture Then
    ShotLine.X2 = ShotLine.X2 + WindSpeed.Text * 2
    End If
If WindPic.Picture = WindUp.Picture Then
    ShotLine.Y2 = ShotLine.Y2 - WindSpeed.Text * 2
    End If
If WindPic.Picture = WindDown.Picture Then
    ShotLine.Y2 = ShotLine.Y2 + WindSpeed.Text * 2
    End If
Swing:
Do
stuck = stuck + 1

d = ShotLine.X1 - ShotLine.X2
e = ShotLine.Y1 - ShotLine.Y2

a = ShotLine.Y2 - ShotLine.Y1
b = ShotLine.X2 - ShotLine.X1

c = Sqr(a * a + b * b)
If c > TotalPower Then
    ShotLine.X2 = ShotLine.X1 - d / 1.005
    ShotLine.Y2 = ShotLine.Y1 - e / 1.005
Else
    GoTo DetectArea
    Exit Sub
End If
If stuck > 500 Then
    GoTo DetectArea
    End If
Loop
DetectArea:

jj = ShotLine.X2 - ShotLine.X1
jj = jj / 2 + ShotLine.X1
j = ShotLine.Y2 - ShotLine.Y1
j = j / 2 + ShotLine.Y1

d = ShotLine.X1 - ShotLine.X2
e = ShotLine.Y1 - ShotLine.Y2

a = ShotLine.Y2 - ShotLine.Y1
b = ShotLine.X2 - ShotLine.X1

k = Sqr(a * a + b * b)

BallMove.Enabled = True
Exit Sub
error:
MsgBox "Error: " & Err.Description, vbCritical, "JeroGolf"
End Sub

Private Sub SwingBut2_Click()
On Error GoTo error
'Get Area
xx = Ball.Left + Ball.Width / 2
yy = Ball.Top + Ball.Height / 2
Sample.BackColor = GetPixel(CoursePic.hdc, xx, yy)
Call GetType
If BallLocate.Text = "Water" Then
WaterShot:
    Ball.Visible = False
    MsgBox "In the Water", vbCritical, "JeroGolf"
    Ball.Left = xxx
    Ball.Top = yyy
    Ball.Visible = True
    HoleScore.Text = HoleScore.Text - -1
    xx = Ball.Left + Ball.Width / 2
    yy = Ball.Top + Ball.Height / 2
    Sample.BackColor = GetPixel(CoursePic.hdc, xx, yy)
    Call GetType
    End If
'ShotLine Stuff
ShotLine.X1 = Ball.Left + Ball.Width / 2
ShotLine.X2 = ShotLine.X1
ShotLine.Y1 = Ball.Top + Ball.Height / 2
ShotLine.Y2 = Ball.Top - 50
'Other Stuff
SwingBut.Enabled = False
way = "up"
PowerTimer.Enabled = False
ProgressBar1.Value = 0
PowerMeterBut.Enabled = True
PowerMeterBut.Caption = "Start"
'Determine Strokes
HoleScore.Text = HoleScore.Text - -1
'Create Distance Left To Hole
Dim aa, bb As Integer
MeaLine.X1 = Ball.Left
MeaLine.Y1 = Ball.Top
If MeaLine.Y1 > MeaLine.Y2 Then
    aa = MeaLine.Y1 - MeaLine.Y2
Else
    aa = ShotLine.Y2 - MeaLine.Y1
End If
If MeaLine.X1 > MeaLine.X2 Then
    bb = MeaLine.X1 - ShotLine.X2
Else
    bb = MeaLine.X2 - MeaLine.X1
End If
Dim i As Integer
LeftDist.Text = Sqr(aa * aa + bb * bb)
If InStr(1, LeftDist.Text, ".", vbTextCompare) <> 0 Then
    i = InStr(1, LeftDist.Text, ".", vbTextCompare)
    LeftDist.Text = Mid(LeftDist.Text, 1, i - 1)
    End If
If BallLocate.Text = "In the Hole" Then
    Ball.Visible = False
    If HoleScore.Text = HolePar.Text + 3 Then
        MsgBox "You got a Triple Bogey", "JeroGolf"
        GoTo DoneMsgBox
        End If
    If HoleScore.Text = HolePar.Text + 2 Then
        MsgBox "You got a Double Bogey", "JeroGolf"
        GoTo DoneMsgBox
        End If
    If HoleScore.Text = HolePar.Text + 1 Then
        MsgBox "You got a Bogey", vbOKOnly, "JeroGolf"
        GoTo DoneMsgBox
        End If
    If HoleScore.Text = HolePar.Text Then
        MsgBox "You got Par", vbOKOnly, "JeroGolf"
        GoTo DoneMsgBox
        End If
    If HoleScore.Text = HolePar.Text - 1 Then
        MsgBox "You got a Birdie", vbOKOnly, "JeroGolf"
        GoTo DoneMsgBox
        End If
    If HoleScore.Text = HolePar.Text - 2 Then
        MsgBox "You got an Eagle", vbOKOnly, "JeroGolf"
        GoTo DoneMsgBox
        End If
    If HoleScore.Text = "1" Then
        MsgBox "You got a Hole in One", vbOKOnly, "JeroGolf"
        GoTo DoneMsgBox
        End If
    If HoleScore.Text = HolePar.Text - 3 Then
        MsgBox "You got an Double Eagle", vbOKOnly, "JeroGolf"
        GoTo DoneMsgBox
        End If
    'If you got over triple bogey
    MsgBox "In the Hole in " & HoleScore.Text & " on a Par " & HolePar.Text, vbOKOnly, "JeroGolf"
DoneMsgBox:
    LeftDist.Text = "0"
    Ball.Visible = True
    ShotLine.Visible = True
    Call NextCourse
    Exit Sub
    End If
'Make a Default Club
If BallLocate.Text = "Sand" Then
    IronBox.ListIndex = IronBox.ListCount - 2
    End If
If BallLocate.Text = "Rough" Then
    IronBox.ListIndex = IronBox.ListCount - 4
    End If
If BallLocate.Text = "FairWay" Then
    If LeftDist.Text <= 50 Then
        IronBox.ListIndex = IronBox.ListCount - 2
        End If
    If LeftDist.Text <= 80 And LeftDist.Text > 50 Then
        IronBox.ListIndex = IronBox.ListCount - 4
        End If
    If LeftDist.Text <= 120 And LeftDist.Text > 80 Then
        IronBox.ListIndex = IronBox.ListCount - 6
        End If
    If LeftDist.Text > 120 Then
        WoodBox.ListIndex = 1
        End If
    End If
If BallLocate.Text = "Green" Then
    IronBox.ListIndex = IronBox.ListCount - 1
    End If
'View Shotline
ShotLine.Visible = True
Exit Sub
error:
MsgBox "Error: " & Err.Description, vbCritical, "JeroGolf"
End Sub

Private Sub WoodBox_Change()
WoodBox.ListIndex = 0
End Sub

Private Sub WoodBox_Click()
Club.Text = WoodBox.Text
End Sub

Public Sub NextCourse()
If HoleNumber.Text = "1" Then
    Course2but.Value = True
    Exit Sub
    End If
If HoleNumber.Text = "2" Then
    Course3but.Value = True
    Exit Sub
    End If
If HoleNumber.Text = "3" Then
    Course4but.Value = True
    Exit Sub
    End If
If HoleNumber.Text = "4" Then
    Course5but.Value = True
    Exit Sub
    End If
If HoleNumber.Text = "5" Then
    MsgBox "Finished Course", vbOKOnly, "JeroGolf"
    Course1but.Value = True
    Exit Sub
    End If
End Sub
