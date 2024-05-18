VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5340
   FillColor       =   &H00C0E0FF&
   ForeColor       =   &H00FF80FF&
   LinkTopic       =   "Form2"
   MousePointer    =   1  'Pfeil
   ScaleHeight     =   214
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   356
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
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
Event DragDrop(Source As Control, x As Single, y As Single)
Event DragOver(Source As Control, x As Single, y As Single, State As Integer)
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
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event OLECompleteDrag(Effect As Long)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event OLESetData(Data As DataObject, DataFormat As Integer)
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Event Paint()
Event QueryUnload(Cancel As Integer, UnloadMode As Integer)
Event Resize()
Event Terminate()
Event Unload(Cancel As Integer)

Private WithEvents Btn2 As CommandButton
Attribute Btn2.VB_VarHelpID = -1


'
Public Property Get Style() As EWndStyle
    Style = MWindow.WindowStyle(Me.hWnd)
End Property
Public Property Let Style(ByVal Value As EWndStyle)
    MWindow.WindowStyle(Me.hWnd) = Value
End Property
Public Function Style_ToStr() As String
    Style_ToStr = MWindow.EWndStyle_ToStr(Me.Style)
End Function

Public Property Get StyleEx() As EWndStyleEx
    StyleEx = MWindow.WindowStyleEx(Me.hWnd)
End Property
Public Property Let StyleEx(ByVal Value As EWndStyleEx)
    MWindow.WindowStyleEx(Me.hWnd) = Value
End Property
Public Function StyleEx_ToStr() As String
    StyleEx_ToStr = MWindow.EWndStyleEx_ToStr(Me.StyleEx)
End Function

'BorderStyle: vbBSNone
'    Style:   WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
'    StyleEx:
'
'BorderStyle: vbFixedSingle
'    Style:   WS_SYSMENU Or WS_DLGFRAME Or WS_BORDER Or WS_CAPTION Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
'    StyleEx: WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW
'
'BorderStyle: vbSizable
'    Style:   WS_MAXIMIZEBOX Or WS_TABSTOP Or WS_GROUP Or WS_MINIMIZEBOX Or WS_SIZEBOX Or WS_THICKFRAME Or WS_SYSMENU Or WS_DLGFRAME Or WS_BORDER Or WS_CAPTION Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
'    StyleEx: WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW
'
'BorderStyle: vbFixedDialog
'    Style:   WS_SYSMENU Or WS_DLGFRAME Or WS_BORDER Or WS_CAPTION Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
'    StyleEx: WS_EX_DLGMODALFRAME Or WS_EX_WINDOWEDGE
'
'BorderStyle: vbFixedToolWindow
'    Style:   WS_SYSMENU Or WS_DLGFRAME Or WS_BORDER Or WS_CAPTION Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
'    StyleEx: WS_EX_TOOLWINDOW Or WS_EX_WINDOWEDGE
'
'BorderStyle: vbSizableToolWindow
'    Style:   WS_SIZEBOX Or WS_THICKFRAME Or WS_SYSMENU Or WS_DLGFRAME Or WS_BORDER Or WS_CAPTION Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
'    StyleEx: WS_EX_TOOLWINDOW Or WS_EX_WINDOWEDGE

Public Function BorderStyle_ToStr() As String
    Dim s As String
    Dim e As FormBorderStyleConstants: e = Me.BorderStyle
    Select Case e
    Case vbBSNone:            s = "vbBSNone"            ' 0
    Case vbFixedSingle:       s = "vbFixedSingle"       ' 1
    Case vbSizable:           s = "vbSizable"           ' 2
    Case vbFixedDialog:       s = "vbFixedDialog"       ' 3
    Case vbFixedToolWindow:   s = "vbFixedToolWindow"   ' 4
    Case vbSizableToolWindow: s = "vbSizableToolWindow" ' 5
    End Select
    BorderStyle_ToStr = s
End Function


Private Sub Command1_Click()
    Debug.Print "Scale- L: " & Me.ScaleLeft & "; T: " & Me.ScaleTop & "; W: " & Me.ScaleWidth & "; H: " & Me.ScaleHeight & ";"
    Debug.Print "Window- L: " & Me.Left & "; T: " & Me.Top & "; W: " & Me.Width & "; H: " & Me.Height & ";"
    
    'the Line-command has a special syntax
    'Line (StartX, StartY)-(EndX, EndY), Color, [B][F]
    
    'Draws a line, using the default foreground color
    Line (20, 50)-(200, 100)
    
    'Debug.Print Me.CurrentX & ", " & Me.CurrentY
    Me.CurrentX = Me.CurrentX + 20
    Me.CurrentY = Me.CurrentY + 20
    
    'Draws a line from the last end point to the new point
    Line -(300, 110)
    
    'Draws a line, using the specified color
    Line (50, 70)-(230, 120), vbRed
    
    'Draws an empty *B*ox, using the default foreground color
    Line (80, 90)-(260, 140), , B
    
    'Draws an empty box, using the specified color for the outline
    Line (110, 110)-(290, 160), vbBlue, B
    
    'Draws a *F*illed box, using the default foreground color for the outline and the filling
    Line (140, 130)-(320, 180), , BF
    
    'Draws a _F_illed box, using the speciefied color for the outline and the filling
    'specified color for the outline and the
    Line (140, 130)-(320, 180), vbGreen, BF
    
    'Debug.Print Me.Count
    'Debug.Print Me.Controls.Count
    
    'Set Btn2 = CreateObject("VB.CommandButton", "Thunder")
    'Btn2.Visible = True
    'Btn2.Width = 150
    'Btn2.Height = 150
    
End Sub

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

Private Sub Form_DragDrop(Source As Control, x As Single, y As Single)
    RaiseEvent DragDrop(Source, x, y)
End Sub

Private Sub Form_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    RaiseEvent DragOver(Source, x, y, State)
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub Form_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
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
