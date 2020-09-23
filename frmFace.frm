VERSION 5.00
Begin VB.Form frmFace 
   Caption         =   "Form2"
   ClientHeight    =   960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2625
   LinkTopic       =   "Form2"
   ScaleHeight     =   960
   ScaleWidth      =   2625
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   855
      Index           =   1
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   0
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1200
      Top             =   2880
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   975
      Index           =   0
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmFace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Sub Form_Load()
Picture1(0).Picture = LoadPicture(App.Path & "\Closekiss.bmp")
Picture1(1).Picture = LoadPicture(App.Path & "\openmouth.bmp")
SetWindowRgn hWnd, CreateEllipticRgn(10, 28, 155, 70), True
End Sub
Private Sub Label1_Click()
Label1.Caption = "Hey U r U"
End Sub
