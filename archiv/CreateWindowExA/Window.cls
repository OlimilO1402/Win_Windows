VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function CreateWindowExA Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
'Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

'für dwExStyle
Public Enum WindowExStyle
  WS_EX_LEFT = &H0&
  WS_EX_LTRREADING = &H0&
  WS_EX_RIGHTSCROLLBAR = &H0&
  WS_EX_DLGMODALFRAME = &H1&
  WS_EX_NOPARENTNOTIFY = &H4&
  WS_EX_TOPMOST = &H8&
  WS_EX_ACCEPTFILES = &H10&
  WS_EX_TRANSPARENT = &H20&
  WS_EX_MDICHILD = &H40&
  WS_EX_TOOLWINDOW = &H80&
  WS_EX_WINDOWEDGE = &H100&
  WS_EX_CLIENTEDGE = &H200&
  WS_EX_CONTEXTHELP = &H400&
  WS_EX_RIGHT = &H1000&
  WS_EX_RTLREADING = &H2000&
  WS_EX_LEFTSCROLLBAR = &H4000&
  WS_EX_CONTROLPARENT = &H10000
  WS_EX_STATICEDGE = &H20000
  WS_EX_APPWINDOW = &H40000
  WS_EX_LAYERED = &H80000
  WS_EX_NOINHERITLAYOUT = &H100000
  WS_EX_LAYOUTRTL = &H400000
  WS_EX_NOACTIVATE = &H8000000
  WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
  WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
End Enum
'für dwStyle
Public Enum WindowStyle
  WS_OVERLAPPED = &H0&
  WS_ACTIVECAPTION = &H1&
  WS_MAXIMIZEBOX = &H10000
  WS_TABSTOP = &H10000
  WS_GROUP = &H20000
  WS_MINIMIZEBOX = &H20000
  WS_THICKFRAME = &H40000
  WS_SYSMENU = &H80000
  WS_HSCROLL = &H100000
  WS_VSCROLL = &H200000
  WS_DLGFRAME = &H400000
  WS_BORDER = &H800000
  WS_CAPTION = &HC00000
  WS_MAXIMIZE = &H1000000
  WS_CLIPCHILDREN = &H2000000
  WS_CLIPSIBLINGS = &H4000000
  WS_DISABLED = &H8000000
  WS_VISIBLE = &H10000000
  WS_MINIMIZE = &H20000000
  WS_CHILD = &H40000000
  WS_POPUP = &H80000000
  WS_CHILDWINDOW = (WS_CHILD)
  WS_SIZEBOX = WS_THICKFRAME
  WS_ICONIC = WS_MINIMIZE
  WS_GT = (WS_GROUP Or WS_TABSTOP)
  WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
  WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
  WS_TILED = WS_OVERLAPPED
  WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
End Enum
Private Const StrMdiClient = "MDICLIENT"
Private Const StrButton = "BUTTON"
Private Const StrEdit = "EDIT"
Private Const StrComboBox = "COMBOBOX"
Private Const StrListBox = "LISTBOX"
Private Const StrScrollBar = "SCROLLBAR"
Private Const StrStatic = "STATIC"
Private Const StrSysIP32 = "SysIPAddress32"

Private Const MDIS_ALLCHILDSTYLES As Long = &H1

Private Const nil As Long = &H0&
Private mHandle As Long
Private mLeft   As Long
Private mTop    As Long
Private mWidth  As Long
Private mHeight As Long
Private mDesktopHWnd As Long
Private mTitle  As String
Private mWStyle As WindowStyle
Private mWExStyle As Long
Private mStrClass As String

Private Sub Class_Initialize()
  mWExStyle = GetWExStyle
  mStrClass = GetStrClass
  mTitle = "Form"
  mWStyle = GetWStyle
  'Visible = False
  'ein moderner Öffnendialog ist ungefähr 563*419 groß
  'd.h. der Öffnendialog wird in der Mitte angezeigt,
  mWidth = 563&
  mHeight = 419&
  mLeft = CLng((Screen.Width / Screen.TwipsPerPixelX) / 2 - mWidth / 2) '1024:=> 230.5
  mTop = CLng((Screen.Height / Screen.TwipsPerPixelY) / 2 - mHeight / 2) '768:=> 174.5
  mDesktopHWnd = GetDesktopWindow
End Sub
Private Sub Class_Terminate()
  DestroyWindow (mHandle)
  Debug.Print "DestroyWindow: " & CStr(mHandle)
End Sub

Private Function GetWStyle() As WindowStyle
  'GetWStyle = GetWStyle Or WS_CAPTION
  'GetWStyle = GetWStyle Or WS_ACTIVECAPTION
  'GetWStyle = GetWStyle Or WS_SYSMENU
  'GetWStyle = GetWStyle Or WS_MAXIMIZEBOX
  'GetWStyle = GetWStyle Or WS_MINIMIZEBOX
  
  'GetWStyle = GetWStyle Or WS_HSCROLL
  'GetWStyle = GetWStyle Or WS_VSCROLL
  
  'GetWStyle = GetWStyle Or WS_TABSTOP
  'GetWStyle = GetWStyle Or WS_GROUP
  
  'GetWStyle = GetWStyle Or WS_DLGFRAME
  'GetWStyle = GetWStyle Or WS_THICKFRAME
  
  'GetWStyle = GetWStyle Or WS_OVERLAPPED
  'GetWStyle = GetWStyle Or WS_BORDER
  'GetWStyle = GetWStyle Or WS_MAXIMIZE
  'GetWStyle = GetWStyle Or WS_MINIMIZE
  'GetWStyle = GetWStyle Or WS_CLIPCHILDREN
  'GetWStyle = GetWStyle Or WS_CLIPSIBLINGS
  'GetWStyle = GetWStyle Or WS_DISABLED
  'GetWStyle = GetWStyle Or WS_VISIBLE
  'GetWStyle = GetWStyle Or WS_CHILD
  'GetWStyle = GetWStyle Or WS_POPUP
End Function
Private Function GetWExStyle() As WindowExStyle
  'GetWExStyle = GetWExStyle Or WS_EX_LEFTSCROLLBAR
  'GetWExStyle = GetWExStyle Or WS_EX_RIGHTSCROLLBAR
  'GetWExStyle = GetWExStyle Or WS_EX_DLGMODALFRAME
  'GetWExStyle = GetWExStyle Or WS_EX_NOPARENTNOTIFY
  'GetWExStyle = GetWExStyle Or WS_EX_TOPMOST
  'GetWExStyle = GetWExStyle Or WS_EX_ACCEPTFILES
  'GetWExStyle = GetWExStyle Or WS_EX_TRANSPARENT
  'GetWExStyle = GetWExStyle Or WS_EX_MDICHILD
  'GetWExStyle = GetWExStyle Or WS_EX_WINDOWEDGE
  'GetWExStyle = GetWExStyle Or WS_EX_CLIENTEDGE
  'GetWExStyle = GetWExStyle Or WS_EX_CONTEXTHELP
  'GetWExStyle = GetWExStyle Or WS_EX_RIGHT
  'GetWExStyle = GetWExStyle Or WS_EX_CONTROLPARENT
  'GetWExStyle = GetWExStyle Or WS_EX_STATICEDGE
  'GetWExStyle = GetWExStyle Or WS_EX_APPWINDOW
  'GetWExStyle = GetWExStyle Or WS_EX_TOOLWINDOW
  'GetWExStyle = GetWExStyle Or WS_EX_LAYERED
  'GetWExStyle = GetWExStyle Or WS_EX_NOINHERITLAYOUT
  'GetWExStyle = GetWExStyle Or WS_EX_NOACTIVATE
  'GetWExStyle = GetWExStyle Or WS_EX_OVERLAPPEDWINDOW
  'GetWExStyle = GetWExStyle Or WS_EX_PALETTEWINDOW
  'GetWExStyle = GetWExStyle Or WS_EX_LEFT
  'GetWExStyle = GetWExStyle Or WS_EX_LTRREADING
  'GetWExStyle = GetWExStyle Or WS_EX_RTLREADING
  'GetWExStyle = GetWExStyle Or WS_EX_LAYOUTRTL
End Function
Private Function GetStrClass() As String
  'GetStrClass = StrMdiClient
  'GetStrClass = StrButton
  'GetStrClass = StrEdit
  'GetStrClass = StrComboBox
  'GetStrClass = StrListBox
  'GetStrClass = StrScrollBar
  GetStrClass = StrStatic
  'GetStrClass = StrSYSIP32
End Function

Public Sub Show()
  mHandle = CreateWindowExA(mWExStyle, mStrClass, mTitle, mWStyle, mLeft, mTop, mWidth, mHeight, mDesktopHWnd, 0&, App.hInstance, 0&)
  Debug.Print mHandle
End Sub

Public Property Get hwnd() As Long
  hwnd = mHandle
End Property

Public Property Get Visible() As Boolean
  Visible = (mWStyle And WS_VISIBLE)
End Property
Public Property Let Visible(BolVal As Boolean)
  mWStyle = mWStyle Or WS_VISIBLE
  If Not BolVal Then mWStyle = mWStyle Xor WS_VISIBLE
End Property

