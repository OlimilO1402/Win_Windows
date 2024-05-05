VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FMain"
   ClientHeight    =   12165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18585
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "FMain"
   ScaleHeight     =   12165
   ScaleWidth      =   18585
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnCreateVBForm 
      Caption         =   "Create VB.Form Form1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton BtnCreateWindow 
      Caption         =   "Create Window Form2"
      Height          =   375
      Left            =   7560
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   11295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Top             =   480
      Width           =   7575
   End
   Begin VB.TextBox Text2 
      Height          =   11295
      Left            =   7560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   480
      Width           =   7575
   End
   Begin VB.CommandButton BtnMoveWindow 
      Caption         =   "Move Window"
      Height          =   375
      Left            =   10080
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private WithEvents VBForm As VB.Form
Private WithEvents Form1 As Form1
Attribute Form1.VB_VarHelpID = -1
Private WithEvents Form2 As Window
Attribute Form2.VB_VarHelpID = -1

Private Sub BtnCreateVBForm_Click()
    Set Form1 = New Form1 'Got it's name/classname in the Property-Editor
    Load Form1
    Form1.Show
End Sub

Private Sub BtnCreateWindow_Click()
    Set Form2 = MNew.Window("Form2") 'Got it's name/classname by the constructor function
    Form2.Load
    Form2.Show
End Sub

Private Sub BtnMoveWindow_Click()
    Form2.Move 100, 100, 800, 600
End Sub

Private Sub Debug_Print1(s As String)
    Text1.Text = Text1.Text & s & vbCrLf
End Sub

Private Sub Debug_Print2(s As String)
    Text2.Text = Text2.Text & s & vbCrLf
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form2 = Nothing
End Sub

' v ############################## v '    Events Form1    ' v ############################## v '
Private Sub Form1_Activate()
    Debug_Print1 "Activate()"
End Sub

Private Sub Form1_Click()
    Debug_Print1 "Click()"
End Sub

Private Sub Form1_DblClick()
    Debug_Print1 "DblClick()"
End Sub

Private Sub Form1_Deactivate()
    Debug_Print1 "Deactivate()"
End Sub

Private Sub Form1_DragDrop(Source As Control, X As Single, Y As Single)
    Debug_Print1 "DragDrop(Source = " & Source.Name & ", X = " & X & ", Y = " & Y & ")"
End Sub

Private Sub Form1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Debug_Print1 "DragOver(Source = " & Source.Name & ", X = " & X & ", Y = " & Y & ", State = " & State & ")"
End Sub

Private Sub Form1_GotFocus()
    Debug_Print1 "GotFocus()"
End Sub

Private Sub Form1_Initialize()
    Debug_Print1 "Initialize()"
End Sub

Private Sub Form1_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug_Print1 "KeyDown(KeyCode = " & KeyCode & ", Shift = " & Shift & ")"
End Sub

Private Sub Form1_KeyPress(KeyAscii As Integer)
    Debug_Print1 "KeyPress(KeyAscii = " & KeyAscii & ")"
End Sub

Private Sub Form1_KeyUp(KeyCode As Integer, Shift As Integer)
    Debug_Print1 "KeyUp(KeyCode = " & KeyCode & ", Shift = " & Shift & ")"
End Sub

Private Sub Form1_LinkClose()
    Debug_Print1 "LinkClose()"
End Sub

Private Sub Form1_LinkError(LinkErr As Integer)
    Debug_Print1 "LinkError(LinkErr = " & LinkErr & ")"
End Sub

Private Sub Form1_LinkExecute(CmdStr As String, Cancel As Integer)
    Debug_Print1 "LinkExecute(CmdStr = " & CmdStr & ", Cancel = " & Cancel & ")"
End Sub

Private Sub Form1_LinkOpen(Cancel As Integer)
    Debug_Print1 "LinkOpen(Cancel = " & Cancel & ")"
End Sub

Private Sub Form1_Load()
    Debug_Print1 "Load()"
End Sub

Private Sub Form1_LostFocus()
    Debug_Print1 "LostFocus()"
End Sub

Private Sub Form1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug_Print1 "MouseDown(Button = " & Button & ", Shift = " & Shift & ", X = " & X & ", Y = " & Y & ")"
End Sub

Private Sub Form1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug_Print1 "MouseMove(Button = " & Button & ", Shift = " & Shift & ", X = " & X & ", Y = " & Y & ")"
End Sub

Private Sub Form1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug_Print1 "MouseUp(Button = " & Button & ", Shift = " & Shift & ", X = " & X & ", Y = " & Y & ")"
End Sub

Private Sub Form1_OLECompleteDrag(Effect As Long)
    Debug_Print1 "OLECompleteDrag(Effect = " & Effect & ")"
End Sub

Private Sub Form1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug_Print1 "OLEDragDrop(Data = " & Data.Files.Count & ", Effect = " & Effect & ", Shift = " & Shift & ", X = " & X & ", Y = " & Y & ")"
End Sub

Private Sub Form1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Debug_Print1 "OLEDragOver(Data = " & Data.Files.Count & ", Effect = " & Effect & ", Button = " & Button & ", Shift = " & Shift & ", X = " & X & ", Y = " & Y & ", State = " & State & ")"
End Sub

Private Sub Form1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    Debug_Print1 "OLEGiveFeedback(Effect = " & Effect & ", DefaultCursors = " & DefaultCursors & ")"
End Sub

Private Sub Form1_OLESetData(Data As DataObject, DataFormat As Integer)
    Debug_Print1 "OLESetData(Data = " & Data.Files.Count & ", DataFormat = " & DataFormat & ")"
End Sub

Private Sub Form1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    Debug_Print1 "OLEStartDrag(Data = " & Data.Files.Count & ", AllowedEffects = " & AllowedEffects & ")"
End Sub

Private Sub Form1_Paint()
    Debug_Print1 "Paint()"
End Sub

Private Sub Form1_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Debug_Print1 "QueryUnload(Cancel = " & Cancel & ", UnloadMode = " & UnloadMode & ")"
End Sub

Private Sub Form1_Resize()
    Debug_Print1 "Resize()"
End Sub

Private Sub Form1_Terminate()
    Debug_Print1 "Terminate()"
End Sub

Private Sub Form1_Unload(Cancel As Integer)
    Debug_Print1 "Unload(Cancel = " & Cancel & ")"
End Sub
' ^ ############################## ^ '    Events Form1    ' ^ ############################## ^ '

' v ############################## v '    Events Form2    ' v ############################## v '
Private Sub Form2_Activate()
    Debug_Print2 "Activate()"
End Sub

Private Sub Form2_Click()
    Debug_Print2 "Click()"
End Sub

Private Sub Form2_DblClick()
    Debug_Print2 "DblClick()"
End Sub

Private Sub Form2_Deactivate()
    Debug_Print2 "Deactivate()"
End Sub

Private Sub Form2_DragDrop(Source As Control, X As Single, Y As Single)
    Debug_Print2 "DragDrop(Source = " & Source.Name & ", X = " & X & ", Y = " & Y & ")"
End Sub

Private Sub Form2_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Debug_Print2 "DragOver(Source = " & Source.Name & ", X = " & X & ", Y = " & Y & ", State = " & State & ")"
End Sub

Private Sub Form2_GotFocus()
    Debug_Print2 "GotFocus()"
End Sub

Private Sub Form2_Initialize()
    Debug_Print2 "Initialize()"
End Sub

Private Sub Form2_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug_Print2 "KeyDown(KeyCode = " & KeyCode & ", Shift = " & Shift & ")"
End Sub

Private Sub Form2_KeyPress(KeyAscii As Integer)
    Debug_Print2 "KeyPress(KeyAscii = " & KeyAscii & ")"
End Sub

Private Sub Form2_KeyUp(KeyCode As Integer, Shift As Integer)
    Debug_Print2 "KeyUp(KeyCode = " & KeyCode & ", Shift = " & Shift & ")"
End Sub

Private Sub Form2_LinkClose()
    Debug_Print2 "LinkClose()"
End Sub

Private Sub Form2_LinkError(LinkErr As Integer)
    Debug_Print2 "LinkError(LinkErr = " & LinkErr & ")"
End Sub

Private Sub Form2_LinkExecute(CmdStr As String, Cancel As Integer)
    Debug_Print2 "LinkExecute(CmdStr = " & CmdStr & ", Cancel = " & Cancel & ")"
End Sub

Private Sub Form2_LinkOpen(Cancel As Integer)
    Debug_Print2 "LinkOpen(Cancel = " & Cancel & ")"
End Sub

Private Sub Form2_Load()
    Debug_Print2 "Load()"
End Sub

Private Sub Form2_LostFocus()
    Debug_Print2 "LostFocus()"
End Sub

Private Sub Form2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug_Print2 "MouseDown(Button = " & Button & ", Shift = " & Shift & ", X = " & X & ", Y = " & Y & ")"
End Sub

Private Sub Form2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug_Print2 "MouseMove(Button = " & Button & ", Shift = " & Shift & ", X = " & X & ", Y = " & Y & ")"
End Sub

Private Sub Form2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug_Print2 "MouseUp(Button = " & Button & ", Shift = " & Shift & ", X = " & X & ", Y = " & Y & ")"
End Sub

Private Sub Form2_OLECompleteDrag(Effect As Long)
    Debug_Print2 "OLECompleteDrag(Effect = " & Effect & ")"
End Sub

Private Sub Form2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug_Print2 "OLEDragDrop(Data = " & Data.Files.Count & ", Effect = " & Effect & ", Shift = " & Shift & ", X = " & X & ", Y = " & Y & ")"
End Sub

Private Sub Form2_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Debug_Print2 "OLEDragOver(Data = " & Data.Files.Count & ", Effect = " & Effect & ", Button = " & Button & ", Shift = " & Shift & ", X = " & X & ", Y = " & Y & ", State = " & State & ")"
End Sub

Private Sub Form2_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    Debug_Print2 "OLEGiveFeedback(Effect = " & Effect & ", DefaultCursors = " & DefaultCursors & ")"
End Sub

Private Sub Form2_OLESetData(Data As DataObject, DataFormat As Integer)
    Debug_Print2 "OLESetData(Data = " & Data.Files.Count & ", DataFormat = " & DataFormat & ")"
End Sub

Private Sub Form2_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    Debug_Print2 "OLEStartDrag(Data = " & Data.Files.Count & ", AllowedEffects = " & AllowedEffects & ")"
End Sub

Private Sub Form2_Paint()
    Debug_Print2 "Paint()"
End Sub

Private Sub Form2_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Debug_Print2 "QueryUnload(Cancel = " & Cancel & ", UnloadMode = " & UnloadMode & ")"
End Sub

Private Sub Form2_Resize()
    Debug_Print2 "Resize()"
End Sub

Private Sub Form2_Terminate()
    Debug_Print2 "Terminate()"
End Sub

Private Sub Form2_Unload(Cancel As Integer)
    Debug_Print2 "Unload(Cancel = " & Cancel & ")"
End Sub

' ^ ############################## ^ '    Events Form2    ' ^ ############################## ^ '
