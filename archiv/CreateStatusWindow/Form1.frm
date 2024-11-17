VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      FillColor       =   &H8000000F&
      ForeColor       =   &H8000000F&
      Height          =   270
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   8835
      TabIndex        =   16
      Top             =   5040
      Width           =   8835
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   720
      TabIndex        =   14
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   315
      Index           =   7
      Left            =   2880
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   315
      Index           =   6
      Left            =   2880
      TabIndex        =   8
      Top             =   960
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   315
      Index           =   5
      Left            =   2880
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   315
      Index           =   4
      Left            =   2880
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   315
      Index           =   3
      Left            =   1800
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   315
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   315
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Value           =   1  'Aktiviert
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   315
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Value           =   1  'Aktiviert
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'mode flag used in case fIsNewComctl = False
'and status bar handle
Dim fIsNewComctl As Boolean
Dim hStatBar As Long



Private Sub Form_Load()
  'initialize the listview classes
  'of the comctl32 dll by calling
  'InitCommonControlsEx. If the dll
  'does not InitCommonControlsEx,
  'the InitComctl32 routine's error is
  'fired, and InitCommonControls is called.
  Call InitComctl32(ICC_LISTVIEW_CLASSES)
   
  'Rtns true & sets the flag if we have the new version
  'of Comctl32.dll.
   fIsNewComctl = InitComctl32(ICC_BAR_CLASSES)
  
   Me.Move (Screen.Width - Width) * 0.5, (Screen.Height - Height) * 0.5
   
  'We need pixels to for some of the msgs.
   ScaleMode = vbPixels
   
  'Enable controls accordingly.
   EnableCtrls False
   
End Sub

'Create a new form, and add the following controls:
'    text box          Text1
'    text box          Text2
'    8 check boxes     Check1(0) - Check1(7)
'    check box         Check2
'    command button    Command1
'    command button    Command2
'    command button    Command3
'    command button    Command4
'    command button    Command4
'Add the following to the general declarations section of the form:
'

Private Sub Command1_Click()

  'Brings a brand new status bar into the world...
   Dim adwParts(1) As Long
   
  'Creates a status bar. The specified text is placed in
  'the one and only part (aka Comctl32.ocx "Panel").
  'Is a bit simpler to call than CreateWindowEx()...
   hStatBar = CreateStatusWindow(GetStyles(), "A status bar...", Me.hWnd, 0)
  
   If hStatBar Then
   
     'When the status bar is created, it will automatically set its
     'own size & position, *unless* either the CCS_NORESIZE
     'or CCS_NOPARENTALIGN styles are specified. We won't
     'bother checking the styles...
      MoveWindow hStatBar, 0, ScaleHeight - 20, ScaleWidth, 20, True
   
     'We'll initially create a status bar with 2 "parts". The 1st is 100
     'pixels less than the width of the status bar, the 2nd is 100
     'pixels wide & extends to the right edge of the status bar.
     '(the SetParts() proc way below doesn't provide for setting
     'individual part widths)
      adwParts(0) = ScaleWidth - 100
      adwParts(1) = -1
   
     'wParam = number of parts
     'lParam = part position array, 0 based
      If SendMessage(hStatBar, SB_SETPARTS, ByVal 2, adwParts(0)) Then
   
        'We'll set the status bar's 2nd panel text now.
        'Each part stores its own text, independent of other parts' text.
        'The text is shown when the part is displayed.
         SetText hStatBar, 1, SBT_SUNKEN, "panel 2"
      End If
   
     'Enables all controls accordingly
      EnableCtrls True
    
  Else
     MsgBox "Uh oh..."
  End If

End Sub


Private Sub EnableCtrls(fEnable As Boolean)

  'Enables/Disables all controls, with the exception of the
  '"Text drawing operation" ctrls, per the fEnable flag.

   Dim cnt As Integer
     
  'Style checkboxes
   For cnt = 2 To 7
      Check1(cnt).Enabled = Not fEnable
   Next
   
   Command1.Enabled = Not fEnable
   Command2.Enabled = fEnable
   Command3.Enabled = fEnable
   Command4.Enabled = fEnable
   Command4.Enabled = True
      
End Sub


Private Function GetStyles() As Long

  'Returns the styles from the selected "Styles" checkboxes.
  '
  'Certain styles act differently when OR'd w/ other styles,
  'producing interesting status bar behavior.

  Dim dwRtn As Long
  
  If Check1(0) Then dwRtn = dwRtn Or WS_VISIBLE
  If Check1(1) Then dwRtn = dwRtn Or WS_CHILD
  If Check1(2) Then dwRtn = dwRtn Or SBARS_SIZEGRIP
  If Check1(3) Then dwRtn = dwRtn Or CCS_TOP
  If Check1(4) Then dwRtn = dwRtn Or CCS_NOMOVEY
  If Check1(5) Then dwRtn = dwRtn Or CCS_BOTTOM
  If Check1(6) Then dwRtn = dwRtn Or CCS_NORESIZE
  If Check1(7) Then dwRtn = dwRtn Or CCS_NOPARENTALIGN

  GetStyles = dwRtn

End Function


Private Sub Command2_Click()
  
  'Frees all resources associated with the progress bar &
  'enables all controls accordingly.
  '
  'If it is not destroyed here, the progress bar will automatically
  'be destroyed when its parent window (the window specified in
  'the hWndParent param of CreateStatusWindow()) is destroyed.

   If IsWindow(hStatBar) Then
      DestroyWindow hStatBar
      hStatBar = 0
      EnableCtrls False
   End If

End Sub

Private Sub Command3_Click()

   SetParts Me, hStatBar, Val(Text1.Text)

End Sub

Private Sub Command4_Click()

   SetText hStatBar, 0, SBT_SUNKEN, (Text2.Text)
        
End Sub




Private Sub Command5_Click()

   If IsWindow(hStatBar) Then DestroyWindow hStatBar
   Unload Me
   
End Sub

Private Sub Form_Resize()
Dim L As Single, T As Single, W As Single, H As Single
  L = 0
  T = Me.ScaleHeight - Picture1.Height
  W = Me.ScaleWidth
  H = Picture1.Height
  Picture1.Move L, T, W, H
End Sub
