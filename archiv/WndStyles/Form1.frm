VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "What is the ""Appearance"" property?"
   ClientHeight    =   9390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19095
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   19095
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   "toggle Appearance 0-1"
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get WndStyle"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1695
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   855
      Left            =   1680
      TabIndex        =   11
      Top             =   6480
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   6120
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   120
      TabIndex        =   9
      Top             =   5400
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   4800
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   600
      Width           =   17055
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   7800
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'ControlSpy: "C:\Program Files (x86)\Microsoft\ControlSpy\ControlSpyV6.exe"
'Spy++:      "C:\Program Files (x86)\Microsoft Visual Studio\Common\Tools\SPYXX.EXE"
'ThunderFormDC
'ThunderPictureBoxDC
'ThunderTextBox

'ThunderFrame
'ThunderCommandButton
'ThunderCheckBox
'ThunderOptionButton
'ThunderComboBox Edit
'ThunderListBox
'ThunderHScrollBar
'ThunderVScrollBar


Private Sub Command2_Click()
    
    Text2.Text = Styles_ToStr(Picture1) & vbCrLf & _
                 Styles_ToStr(Text1) & vbCrLf & _
                 Styles_ToStr(Frame1) & vbCrLf & _
                 Styles_ToStr(Command1) & vbCrLf & _
                 Styles_ToStr(Check1) & vbCrLf & _
                 Styles_ToStr(Option1) & vbCrLf & _
                 Styles_ToStr(Combo1) & vbCrLf & _
                 Styles_ToStr(List1) & vbCrLf & _
                 Styles_ToStr(HScroll1) & vbCrLf & _
                 Styles_ToStr(VScroll1)
        
End Sub

Function Styles_ToStr(Ctrl As Object) As String
    Dim Ctrl_Type As String:  Ctrl_Type = TypeName(Ctrl)
    Dim Ctrl_hWnd As LongPtr: Ctrl_hWnd = Ctrl.hWnd
    Styles_ToStr = Ctrl_Type & ".  Style = " & MWin.EWndStyle_ToStr(MWin.WindowStyle(Ctrl_hWnd)) & vbCrLf & _
                   Ctrl_Type & ".ExStyle = " & MWin.EWndStyleEx_ToStr(MWin.WindowStyleEx(Ctrl_hWnd)) & vbCrLf
End Function

Private Sub Command3_Click()
    ToggleAppearance Picture1
    ToggleAppearance Text1
    ToggleAppearance Frame1
    ToggleAppearance Command1
    ToggleAppearance Check1
    ToggleAppearance Option1
    ToggleAppearance Combo1
    ToggleAppearance List1
    ToggleAppearance HScroll1
    ToggleAppearance VScroll1
    Command2_Click
End Sub

Sub ToggleAppearance(Ctrl As Object)
On Error Resume Next
    Dim a As Long: a = IIf(Ctrl.Appearance, 0, 1)
    Ctrl.Appearance = a
End Sub

'Appearance = 1 - 3D:
'====================
'PictureBox.Style = &H56010000 = WS_TABSTOP Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD
'PictureBox.ExStyle = &H204 = WS_EX_NOPARENTNOTIFY Or WS_EX_CLIENTEDGE
'
'TextBox.Style = &H540100C0 = WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD Or &HC0
'TextBox.ExStyle = &H204 = WS_EX_NOPARENTNOTIFY Or WS_EX_CLIENTEDGE
'
'Frame.Style = &H56010007 = WS_TABSTOP Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD Or &H7
'Frame.ExStyle = &H4 = WS_EX_NOPARENTNOTIFY
'
'CommandButton.Style = &H54012000 = WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD Or &H2000
'CommandButton.ExStyle = &H4 = WS_EX_NOPARENTNOTIFY
'
'CheckBox.Style = &H54012006 = WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD Or &H2006
'CheckBox.ExStyle = &H4 = WS_EX_NOPARENTNOTIFY
'
'OptionButton.Style = &H54012004 = WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD Or &H2004
'OptionButton.ExStyle = &H0 =
'
'ComboBox.Style = &H54010242 = WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD Or &H242
'ComboBox.ExStyle = &H0 =
'
'ListBox.Style = &H54010081 = WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD Or &H81
'ListBox.ExStyle = &H204 = WS_EX_NOPARENTNOTIFY Or WS_EX_CLIENTEDGE
'
'HScrollBar.Style = &H54010000 = WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD
'HScrollBar.ExStyle = &H4 = WS_EX_NOPARENTNOTIFY
'
'VScrollBar.Style = &H54010001 = WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD Or &H1
'VScrollBar.ExStyle = &H4 = WS_EX_NOPARENTNOTIFY
'
'
'Appearance = 0 - 2D:
'====================
'PictureBox.Style = &H56810000 = WS_TABSTOP Or WS_BORDER Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD
'PictureBox.ExStyle = &H4 = WS_EX_NOPARENTNOTIFY
'
'TextBox.Style = &H548100C0 = WS_TABSTOP Or WS_BORDER Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD Or &HC0
'TextBox.ExStyle = &H4 = WS_EX_NOPARENTNOTIFY
'
'Frame.Style = &H56018007 = WS_TABSTOP Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD Or &H8007
'Frame.ExStyle = &H4 = WS_EX_NOPARENTNOTIFY
'
'CommandButton.Style = &H54012000 = WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD Or &H2000
'CommandButton.ExStyle = &H4 = WS_EX_NOPARENTNOTIFY
'
'CheckBox.Style = &H5401A006 = WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD Or &HA006
'CheckBox.ExStyle = &H4 = WS_EX_NOPARENTNOTIFY
'
'OptionButton.Style = &H5401A004 = WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD Or &HA004
'OptionButton.ExStyle = &H0 =
'
'ComboBox.Style = &H54010242 = WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD Or &H242
'ComboBox.ExStyle = &H0 =
'
'ListBox.Style = &H54810081 = WS_TABSTOP Or WS_BORDER Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD Or &H81
'ListBox.ExStyle = &H4 = WS_EX_NOPARENTNOTIFY
'
'HScrollBar.Style = &H54010000 = WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD
'HScrollBar.ExStyle = &H4 = WS_EX_NOPARENTNOTIFY
'
'VScrollBar.Style = &H54010001 = WS_TABSTOP Or WS_CLIPSIBLINGS Or WS_VISIBLE Or WS_CHILD Or &H1
'VScrollBar.ExStyle = &H4 = WS_EX_NOPARENTNOTIFY

