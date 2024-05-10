VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Click()
Event DblClick()
Event Deactivate()
Event DragDrop(Source As Control, X As Single, Y As Single)
Event DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Event GotFocus()
Event Initialize()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event LinkClose()
Event LinkError(LinkErr As Integer)
Event LinkExecute(CmdStr As String, Cancel As Integer)
Event LinkOpen(Cancel As Integer)
Event Load()
Event LostFocus()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event OLECompleteDrag(Effect As Long)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event OLESetData(Data As DataObject, DataFormat As Integer)
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Event Paint()
Event QueryUnload(Cancel As Integer, UnloadMode As Integer)
Event Resize()
Event Terminate()
Event Unload(Cancel As Integer)

Private Sub Form_Activate()
    RaiseEvent Activate
End Sub

Private Sub Form_Click()
    RaiseEvent Click
End Sub

Private Sub Form_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Form_Deactivate()
    RaiseEvent Deactivate
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    RaiseEvent DragDrop(Source, X, Y)
End Sub

Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    RaiseEvent DragOver(Source, X, Y, State)
End Sub

Private Sub Form_GotFocus()
    RaiseEvent GotFocus
End Sub

Private Sub Form_Initialize()
    RaiseEvent Initialize
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Form_LinkClose()
    RaiseEvent LinkClose
End Sub

Private Sub Form_LinkError(LinkErr As Integer)
    RaiseEvent LinkError(LinkErr)
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    RaiseEvent LinkExecute(CmdStr, Cancel)
End Sub

Private Sub Form_LinkOpen(Cancel As Integer)
    RaiseEvent LinkOpen(Cancel)
End Sub

Private Sub Form_Load()
    RaiseEvent Load
End Sub

Private Sub Form_LostFocus()
    RaiseEvent LostFocus
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Form_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub Form_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub Form_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub Form_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub Form_Paint()
    RaiseEvent Paint
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    RaiseEvent QueryUnload(Cancel, UnloadMode)
End Sub

Private Sub Form_Resize()
    RaiseEvent Resize
End Sub

Private Sub Form_Terminate()
    RaiseEvent Terminate
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RaiseEvent Unload(Cancel)
End Sub
