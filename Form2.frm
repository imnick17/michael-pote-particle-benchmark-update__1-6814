VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   6345
   ClientLeft      =   735
   ClientTop       =   930
   ClientWidth     =   9540
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   20.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   423
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   636
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private StopNow As Boolean, C As String, Tick As Integer

Private Sub Form_DblClick()
StopNow = True
Form2.Hide
End Sub

Private Sub Form_Load()
Show
StopNow = False
Do
DoEvents
Select Case Tick
Case 0
C = "PARTICLE BENCHMARK"
Case 1
C = "Programed by Michael Pote the Particle Master"
Case 2
C = "Check out my other submitions"
Case 3
C = "Mikes Sparks Program"
Case 4
C = "Have a look at it, it rules!"
Case 5
C = "All at planet source code."
Case 6
C = "Even this credit box is vauable code!"
Case 7
C = "Email me at Mikepote@mailcity.com"
Case 8
C = "Double Click to return"
Tick = -1
End Select
Form2.ForeColor = RGB(Rnd * 50 + 200, Rnd * 100 + 150, Rnd * 100 + 150)
CurrentX = ScaleWidth / 2 - (TextWidth(C) / 2)
CurrentY = ScaleHeight / 2 - (TextHeight(C) / 2)
Print C
Sleep 700
For I = 0 To 100
DoEvents
On Error Resume Next
ForeColor = ForeColor - RGB(5, 5, 5)
CurrentX = ScaleWidth / 2 - (TextWidth(C) / 2)
CurrentY = ScaleHeight / 2 - (TextHeight(C) / 2)
Print C
StretchBlt Form2.hdc, 0, 0, ScaleWidth, ScaleHeight, Form2.hdc, 20, 20, ScaleWidth - 40, ScaleHeight - 40, SRCCOPY
If StopNow Then Exit Sub
Next
Cls
Tick = Tick + 1
DoEvents
Loop Until StopNow = True
On Error Resume Next
Form2.Hide
End Sub
