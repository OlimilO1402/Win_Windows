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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mForm2 As Window

Private Sub BtnCreateWindow_Click()
    Set mForm2 = MNew.Window("ThunderVB64uWindow")
    mForm2.Load
    'mForm2.Show
End Sub

Private Sub BtnMoveWindow_Click()
    mForm2.Move 100, 100, 800, 600
    
End Sub

Private Sub Form_Activate()
    '
End Sub

Private Sub Form_Click()
    '
End Sub

Private Sub Form_DblClick()
    '
End Sub

Private Sub Form_Deactivate()
    '
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    '
End Sub

Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    '
End Sub

Private Sub Form_GotFocus()
    '
End Sub

Private Sub Form_Initialize()
    '
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    '
End Sub

Private Sub Form_LinkClose()
    '
End Sub

Private Sub Form_LinkError(LinkErr As Integer)
    '
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    '
End Sub

Private Sub Form_LinkOpen(Cancel As Integer)
    '
End Sub

Private Sub Form_Load()
    '
End Sub

Private Sub Form_LostFocus()
    '
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
End Sub

Private Sub Form_OLECompleteDrag(Effect As Long)
    '
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    '
End Sub

Private Sub Form_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    '
End Sub

Private Sub Form_OLESetData(Data As DataObject, DataFormat As Integer)
    '
End Sub

Private Sub Form_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    '
End Sub

Private Sub Form_Paint()
    '
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
End Sub

Private Sub Form_Resize()
    '
End Sub

Private Sub Form_Terminate()
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
End Sub
