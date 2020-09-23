VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Particle Benchmark by Michael Pote"
   ClientHeight    =   5700
   ClientLeft      =   1650
   ClientTop       =   1560
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4050
      Left            =   315
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   270
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   280
      TabIndex        =   3
      Top             =   45
      Visible         =   0   'False
      Width           =   4200
      Begin VB.CheckBox Check5 
         Caption         =   "Shrinking"
         Height          =   315
         Left            =   2550
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Shrinking particles using StretchBlt instead of bitBlt"
         Top             =   1050
         Width           =   945
      End
      Begin VB.CommandButton Command4 
         Caption         =   "CLS"
         Height          =   270
         Left            =   3435
         TabIndex        =   26
         Top             =   2295
         Width           =   645
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H80000016&
         Caption         =   "Circular"
         Height          =   270
         Left            =   435
         TabIndex        =   25
         ToolTipText     =   "Particles rotate in a circular pattern"
         Top             =   1140
         Width           =   1440
      End
      Begin VB.PictureBox Picture4 
         Height          =   330
         Left            =   2700
         ScaleHeight     =   270
         ScaleWidth      =   255
         TabIndex        =   23
         Top             =   1950
         Width           =   315
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            Height          =   150
            Left            =   60
            Picture         =   "Form1.frx":37632
            ScaleHeight     =   10
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   10
            TabIndex        =   24
            Top             =   60
            Width           =   150
         End
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Fill Screen"
         Height          =   255
         Left            =   2250
         TabIndex        =   22
         ToolTipText     =   "Choose a big picture and check this!"
         Top             =   2310
         Width           =   1110
      End
      Begin VB.CommandButton Command3 
         Cancel          =   -1  'True
         Caption         =   "Revert"
         Height          =   285
         Left            =   2055
         TabIndex        =   21
         Top             =   1965
         Width           =   645
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Icicles"
         Height          =   240
         Left            =   450
         TabIndex        =   20
         Top             =   2010
         Width           =   1410
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Set Picture"
         Height          =   285
         Left            =   3030
         TabIndex        =   18
         Top             =   1965
         Width           =   1080
      End
      Begin VB.VScrollBar YSpeed 
         Height          =   1215
         Left            =   705
         Max             =   2
         Min             =   100
         TabIndex        =   15
         Top             =   2535
         Value           =   10
         Width           =   300
      End
      Begin VB.VScrollBar Xspeed 
         Height          =   1215
         Left            =   165
         Max             =   2
         Min             =   100
         TabIndex        =   14
         Top             =   2535
         Value           =   10
         Width           =   300
      End
      Begin VB.HScrollBar LifeLine 
         Height          =   240
         LargeChange     =   100
         Left            =   2115
         Max             =   1000
         Min             =   1
         SmallChange     =   10
         TabIndex        =   12
         Top             =   1425
         Value           =   75
         Width           =   1770
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Friction"
         Height          =   240
         Left            =   450
         TabIndex        =   10
         Top             =   1740
         Value           =   1  'Checked
         Width           =   1410
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Gravity"
         Height          =   240
         Left            =   450
         TabIndex        =   9
         ToolTipText     =   "Best viewed without friction"
         Top             =   1470
         Width           =   1410
      End
      Begin VB.HScrollBar NumParts 
         Height          =   255
         LargeChange     =   100
         Left            =   2100
         Max             =   10000
         Min             =   1
         SmallChange     =   10
         TabIndex        =   7
         Top             =   540
         Value           =   650
         Width           =   1770
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000016&
         Caption         =   "Organic"
         Height          =   270
         Left            =   435
         TabIndex        =   6
         ToolTipText     =   "Particles swim around like amoeba"
         Top             =   855
         Width           =   1440
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000016&
         Caption         =   "Mouse"
         Height          =   270
         Left            =   435
         TabIndex        =   5
         ToolTipText     =   "Particles follow mouse"
         Top             =   570
         Value           =   -1  'True
         Width           =   1440
      End
      Begin VB.Label lblBench 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   1140
         TabIndex        =   29
         ToolTipText     =   "CHeck this out"
         Top             =   3315
         Width           =   2010
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "X and Y Speed"
         Height          =   225
         Left            =   75
         TabIndex        =   28
         Top             =   2310
         Width           =   1140
      End
      Begin VB.Label lblAbout 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   1980
         TabIndex        =   19
         ToolTipText     =   "CHeck this out"
         Top             =   2685
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   1110
         TabIndex        =   17
         Top             =   3015
         Width           =   90
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   540
         TabIndex        =   16
         Top             =   3000
         Width           =   90
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Lifetime: 100"
         Height          =   195
         Left            =   2175
         TabIndex        =   13
         Top             =   1680
         Width           =   1755
      End
      Begin VB.Label lblOK 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   3225
         TabIndex        =   11
         Top             =   2685
         Width           =   780
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "600 Particles"
         Height          =   195
         Left            =   2100
         TabIndex        =   8
         Top             =   825
         Width           =   1755
      End
      Begin VB.Label lblExit 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   3255
         TabIndex        =   4
         Top             =   3390
         Width           =   780
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Height          =   270
      Left            =   30
      MaskColor       =   &H00000000&
      Picture         =   "Form1.frx":377B4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   30
      UseMaskColor    =   -1  'True
      Width           =   300
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   5715
      Picture         =   "Form1.frx":37936
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   1
      Top             =   3990
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   5550
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   0
      Top             =   2910
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Par
 X As Variant
 Y As Variant
 V As Variant
 Sv As Variant
 Life As Integer
End Type
Private Bench As Boolean, Tim As Variant, Fps As Variant
Private P(0 To 10000) As Par, I As Integer, Size As Integer, T As Integer, Ang As Integer

Private Sub Command1_Click()
picOptions.Visible = True
End Sub

Private Sub Command2_Click()
Dim file As String
file = OpenDialog(Form1, "Bitmaps|*.bmp;*.jpg;*.ico", "Open Particle Picture", "")
If file = "" Then Exit Sub
Picture1.Picture = LoadPicture(file)
Picture2.Move 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
End Sub

Private Sub Command3_Click()
Let Picture1.Picture = Command1.Picture
End Sub

Private Sub Command4_Click()
Form1.Cls
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' THE PARTICLE ENGINE!!!!
If Bench Then Tim = Timer
For I = 0 To NumParts.Value ' Loop through however many particles set by numparts slider
With P(I)  'With the current particle
If Check3.Value = 0 And Check4.Value = 0 Then BitBlt hdc, .X, .Y, Picture1.Width, Picture1.Height, Picture2.hdc, 0, 0, SRCCOPY  'Clear the space underneath the particle
If .Life <= 0 Then 'Life has expired
If Option1.Value Then  'Place new particle at mouse cursor
.X = X
.Y = Y
ElseIf Option3.Value = True Then
Ang = Ang + 1
If Ang >= 360 Then Let Ang = 0
.X = (ScaleWidth / 2) + Sin((3.145 / 180) * Ang) * 250
.Y = (ScaleHeight / 2) + Cos((3.145 / 180) * Ang) * 250
End If
' Set velocitys and lifetime
.V = (Rnd * YSpeed.Value) - (YSpeed.Value / 2)
.Sv = (Rnd * Xspeed.Value) - (Xspeed.Value / 2)
.Life = Rnd * LifeLine.Value
End If
' Main Loop
.Life = .Life - 1 'Count down life
.X = .X + .Sv
.Y = .Y + .V
' apply physics to x & y coords
If Check1.Value = 1 Then
.V = .V + 0.1 'Appy Gravity
If .Y >= ScaleHeight Then Let .V = -(.V * 0.6) 'Bounce Particle
End If
If Check2.Value = 1 Then 'Appy Friction
If .V < 0 Then Let .V = .V + 0.1
If .V > 0 Then Let .V = .V - 0.1
If .Sv < 0 Then Let .Sv = .Sv + 0.1
If .Sv > 0 Then Let .Sv = .Sv - 0.1
End If
' Blit the particle onto the form
If Check4.Value = 1 Then
BitBlt hdc, .X, .Y, 10, 10, Picture1.hdc, .X, .Y, SRCCOPY
Else
If Check5.Value = 1 Then
Size = 10 / (.Life + 2)
Trans Picture1, Picture3, .X, .Y, Form1.hdc, Size
Else
BitBlt hdc, .X, .Y, 10, 10, Picture1.hdc, 0, 0, SRCPAINT
End If
End If
End With
Next
If Bench Then
Bench = False
Fps = Timer - Tim
MsgBox "Your computer scores " & Format(Fps, "0.00") & " seconds to draw 10 000 dynamic objects.", vbInformation
End If
' END OF PARTICLE ENGINE
End Sub

Private Sub lblAbout_Click()
Form2.Show 1, Form1
End Sub

Private Sub lblBench_Click()
Dim Res As Integer
Res = NumParts.Value
NumParts.Value = 10000
Bench = True
Form_MouseMove 1, 1, 1, 1
NumParts.Value = Res
End Sub

Private Sub lblExit_Click()
End
Unload Me
End Sub

Private Sub LifeLine_Change()
LifeLine_Scroll
End Sub

Private Sub NumParts_Change()
Form1.Cls
NumParts_Scroll
End Sub

Private Sub NumParts_Scroll()
Let Label1.Caption = NumParts.Value & " Particles"
End Sub

Private Sub lblOK_Click()
picOptions.Visible = False
End Sub

Private Sub LifeLine_Scroll()
Let Label2.Caption = "Lifetime: " & LifeLine.Value
End Sub

Private Sub Xspeed_Change()
Label4.Caption = Xspeed.Value

End Sub

Private Sub Xspeed_Scroll()
Label4.Caption = Xspeed.Value
End Sub

Private Sub YSpeed_Change()
Label5.Caption = YSpeed.Value

End Sub

Private Sub YSpeed_Scroll()
Label5.Caption = YSpeed.Value
End Sub
