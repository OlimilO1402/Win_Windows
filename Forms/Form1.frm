VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnMoveWindow 
      Caption         =   "Move Window"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton BtnCreateWindow 
      Caption         =   "Create Window"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mForm2 As Window
Attribute mForm2.VB_VarHelpID = -1

Private Sub BtnCreateWindow_Click()
    Set mForm2 = MNew.Window("Form2")
    mForm2.Load
    'mForm2.Show
End Sub

Private Sub BtnMoveWindow_Click()
    mForm2.Move 100, 100, 800, 600
End Sub

Private Sub mForm2_Activate()
    Label1.Caption = "Form2_Activate"
    DoEvents
End Sub

Private Sub mForm2_Deactivate()
    Label1.Caption = "Form2_Deactivate"
    DoEvents
End Sub

Private Sub mForm2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Label1.Caption = "Button: " & Button & " Shift: " & Shift & " X: " & X & " Y: " & Y
    'DoEvents
End Sub
