VERSION 5.00
Begin VB.Form frmTest 
   Appearance      =   0  'Flat
   Caption         =   "Test SideBar Logo - Vertical Text"
   ClientHeight    =   5700
   ClientLeft      =   2850
   ClientTop       =   2025
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   6720
   Begin VB.PictureBox picLogo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   9375
      Left            =   0
      ScaleHeight     =   9375
      ScaleWidth      =   465
      TabIndex        =   0
      Top             =   0
      Width           =   465
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cL As New cLogo
Private Sub Form_Load()
    cL.DrawingObject = picLogo
    cL.Caption = "Vincenzo De Cristofaro"
    cL.StartColor = vbBlue
    cL.EndColor = vbBlack

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picLogo.Height = Me.ScaleHeight
    On Error GoTo 0
    cL.Draw
End Sub

