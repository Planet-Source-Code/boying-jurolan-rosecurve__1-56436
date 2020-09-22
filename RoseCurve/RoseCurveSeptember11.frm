VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rose Curve"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Help 
      BackColor       =   &H00FFFFFF&
      Caption         =   "?"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "About"
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox Leaf 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "Enter Number of Leaves"
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Clear 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clear"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Clear Picture"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4575
      ScaleWidth      =   4575
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   0
         Top             =   0
      End
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stop"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "Clear Picture"
      ToolTipText     =   "Start/Stop Drawing"
      Top             =   4920
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim counter As Single
Dim x As Single
Dim y As Single
Dim R As Single
Dim a As Single
Dim T As Single
Dim N As Single

Private Sub Form_Load()
Leaf.Text = 10
Leaf.Enabled = False
End Sub

Private Sub Clear_Click()
Picture1.Cls
End Sub

Private Sub cmdStop_Click()
If cmdStop.Caption = "Stop" Then
Timer1.Enabled = False
cmdStop.Caption = "Resume"
Leaf.Enabled = True
Else
Timer1.Enabled = True
cmdStop.Caption = "Stop"
Leaf.Enabled = False
End If
End Sub

Private Sub Help_Click()
MsgBox "This program draws rose curves. The Text box indicates how many leaves are drawn. When it is even the number of curves are equal to twice that number. When it is odd the number of curves is equal to that number.", vbOKOnly + vbInformation, "About"
End Sub

Private Sub Timer1_Timer()
N = Val(Leaf.Text)
counter = counter + 1
Randomize
a = Rnd * 10
Picture1.ForeColor = QBColor(Rnd * 15)
Call Draw

End Sub


Private Sub Draw()

Picture1.Scale (-10, 10)-(10, -10)

For T = 0 To (2 * 4 * Atn(1)) Step 0.0001
R = a * Sin(N * T)
'In this part when T is replaced with a
'spikes are drwan instead of curves
x = R * Cos(T)
y = R * Sin(T)
Picture1.PSet (x, y)
Next T

End Sub
