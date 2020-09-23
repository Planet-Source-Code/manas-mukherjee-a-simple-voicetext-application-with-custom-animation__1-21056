VERSION 5.00
Begin VB.Form frmVTxt 
   BackColor       =   &H80000009&
   Caption         =   "vText"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Stop Speaking"
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4080
      Top             =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "End "
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdResume 
      BackColor       =   &H00FFFFFF&
      Caption         =   "resume"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdPause 
      BackColor       =   &H80000009&
      Caption         =   "Pause"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmVTxt.frx":0000
      Top             =   480
      Width           =   3495
   End
   Begin VB.CommandButton cmdSpeak 
      BackColor       =   &H80000009&
      Caption         =   "Speak"
      Height          =   495
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000C0&
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   0
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   135
   End
End
Attribute VB_Name = "frmVTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myvoice As VTxtAuto.VTxtAuto
Dim bPause As Boolean
Dim strSpeak As String
Dim bresume As Boolean
Dim Y As Integer
Private Sub cmdPause_Click()
On Error Resume Next
If myvoice.IsSpeaking Then myvoice.AudioPause 'Text1
bPause = True
Timer1.Enabled = False
frmFace.Picture1(0).Visible = True
frmFace.Picture1(1).Visible = False
End Sub
Private Sub cmdResume_Click()
On Error Resume Next
Dim strSpeak As String
strSpeak = Text1.Text
If bPause = True Then
myvoice.AudioResume '
bPause = False
bresume = True
Timer1.Enabled = True
End If
If myvoice.IsSpeaking = False Then Timer1.Enabled = False
End Sub
Private Sub cmdSpeak_Click()
If myvoice.IsSpeaking Or bPause = True Then Exit Sub
myvoice.Speak Text1, vtxtst_READING
If Timer1.Enabled = False Then Timer1.Enabled = True
Y = 0


End Sub
Private Sub Command1_Click()
Form_Terminate
End
End Sub

Private Sub Command2_Click()
On Error Resume Next
 myvoice.StopSpeaking
End Sub

Private Sub Form_Load()
Timer1.Interval = 100
Timer1.Enabled = False
Set myvoice = New VTxtAuto.VTxtAuto
myvoice.Register App.Title, App.EXEName
frmFace.Show
frmFace.Picture1(0).Visible = True
frmFace.Picture1(1).Visible = False
frmVTxt.Move 1600, 1600
frmFace.Move 300, 300
End Sub
Private Sub Form_Terminate()
Set frmFace = Nothing
If myvoice.IsSpeaking Then myvoice.StopSpeaking
Set myvoice = Nothing
Set frmVTxt = Nothing
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
If myvoice.IsSpeaking Then myvoice.StopSpeaking
Unload frmFace
End Sub

Private Sub Timer1_Timer()
Static x As Integer
'x = Left
x = Y
x = x + 10
Y = x
If myvoice.IsSpeaking Then
If frmFace.Picture1(0).Visible = True Then
frmFace.Picture1(0).Visible = False
frmFace.Picture1(1).Visible = True
    If Y = ScaleWidth Then
    Shape1.Left = 0
    Y = 0
    End If
    Shape1.Move Y
Else
frmFace.Picture1(0).Visible = True
frmFace.Picture1(1).Visible = False

End If
Else
frmFace.Picture1(0).Visible = True
frmFace.Picture1(1).Visible = False
End If

End Sub

