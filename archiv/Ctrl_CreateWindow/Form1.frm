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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Windows styles
Private Const WS_TILED            As Long = &H0            ' The window is an overlapped window. An overlapped window has a title bar and a border. Same as the WS_OVERLAPPED style.
Private Const WS_OVERLAPPED       As Long = &H0            ' The window is an overlapped window. An overlapped window has a title bar and a border. Same as the WS_TILED style.
Private Const WS_ACTIVECAPTION    As Long = &H1&

Private Const WS_TABSTOP          As Long = &H10000        ' The window is a control that can receive the keyboard focus when the user presses the TAB key. Pressing the TAB key changes the keyboard focus to the next control with the WS_TABSTOP style.
Private Const WS_GROUP            As Long = &H20000        ' The window is the first control of a group of controls. The group consists of this first control and all controls defined after it, up to the next control with the WS_GROUP style. The first control in each group usually has the WS_TABSTOP style so that the user can move from group to group. The user can subsequently change the keyboard focus from one control in the group to the next control in the group by using the direction keys.
Private Const WS_MAXIMIZEBOX      As Long = &H10000        ' The window has a maximize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also be specified.
Private Const WS_MINIMIZEBOX      As Long = &H20000        ' The window has a minimize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also be specified.
Private Const WS_SIZEBOX          As Long = &H40000        ' The window has a sizing border. Same as the WS_THICKFRAME style.
Private Const WS_THICKFRAME       As Long = &H40000        ' The window has a sizing border. Same as the WS_SIZEBOX style.
Private Const WS_SYSMENU          As Long = &H80000        ' The window has a window menu on its title bar. The WS_CAPTION style must also be specified.

Private Const WS_HSCROLL          As Long = &H100000       ' The window has a horizontal scroll bar.
Private Const WS_VSCROLL          As Long = &H200000       ' The window has a vertical scroll bar.
Private Const WS_DLGFRAME         As Long = &H400000       ' The window has a border of a style typically used with dialog boxes. A window with this style cannot have a title bar.
Private Const WS_BORDER           As Long = &H800000       ' The window has a thin-line border
Private Const WS_CAPTION          As Long = &HC00000       ' The window has a title bar (includes the WS_BORDER style).

Private Const WS_MAXIMIZE         As Long = &H1000000      ' The window is initially maximized.
Private Const WS_CLIPCHILDREN     As Long = &H2000000      ' Excludes the area occupied by child windows when drawing occurs within the parent window. This style is used when creating the parent window.
Private Const WS_CLIPSIBLINGS     As Long = &H4000000      ' Clips child windows relative to each other; that is, when a particular child window receives a WM_PAINT message, the WS_CLIPSIBLINGS style clips all other overlapping child windows out of the region of the child window to be updated. If WS_CLIPSIBLINGS is not specified and child windows overlap, it is possible, when drawing within the client area of a child window, to draw within the client area of a neighboring child window.
Private Const WS_DISABLED         As Long = &H8000000      ' The window is initially disabled. A disabled window cannot receive input from the user. To change this after a window has been created, use the EnableWindow function.

Private Const WS_VISIBLE          As Long = &H10000000     ' The window is initially visible. This style can be turned on and off by using the ShowWindow or SetWindowPos function.
Private Const WS_ICONIC           As Long = &H20000000     ' The window is initially minimized. Same as the WS_MINIMIZE style.
Private Const WS_MINIMIZE         As Long = &H20000000     ' The window is initially minimized. Same as the WS_ICONIC style.
Private Const WS_CHILD            As Long = &H40000000     ' The window is a child window. A window with this style cannot have a menu bar. This style cannot be used with the WS_POPUP style.
Private Const WS_CHILDWINDOW      As Long = &H40000000     ' Same as the WS_CHILD style.
Private Const WS_POPUP            As Long = &H80000000     ' The window is a pop-up window. This style cannot be used with the WS_CHILD style.
                                                           ' You can turn this style on and off to change dialog box navigation. To change this style after a window has been created, use the SetWindowLong function.
Private Const WS_OVERLAPPEDWINDOW As Long = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX) 'The window is an overlapped window. Same as the WS_TILEDWINDOW style.
Private Const WS_TILEDWINDOW      As Long = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX) 'The window is an overlapped window. Same as the WS_OVERLAPPEDWINDOW style.
Private Const WS_POPUPWINDOW      As Long = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)       'The window is a pop-up window. The WS_CAPTION and WS_POPUPWINDOW styles must be combined to make the window menu visible.
                                                           ' You can turn this style on and off to change dialog box navigation. To change this style after a window has been created, use the SetWindowLong function. For user-created windows and modeless dialogs to work with tab stops, alter the message loop to call the IsDialogMessage function.



Private Const WS_EX_LEFT                As Long = &H0         '    The window has generic left-aligned properties. This is the default.
Private Const WS_EX_LTRREADING          As Long = &H0         '    The window text is displayed using left-to-right reading-order properties. This is the default.
Private Const WS_EX_RIGHTSCROLLBAR      As Long = &H0         '    The vertical scroll bar (if present) is to the right of the client area. This is the default.
Private Const WS_EX_DLGMODALFRAME       As Long = &H1         '    The window has a double border; the window can, optionally, be created with a title bar by specifying the WS_CAPTION style in the dwStyle parameter.
'                                        As Long = &H2&
Private Const WS_EX_NOPARENTNOTIFY      As Long = &H4         '    The child window created with this style does not send the WM_PARENTNOTIFY message to its parent window when it is created or destroyed.
Private Const WS_EX_TOPMOST             As Long = &H8         '    The window should be placed above all non-topmost windows and should stay above them, even when the window is deactivated. To add or remove this style, use the SetWindowPos function.

Private Const WS_EX_ACCEPTFILES         As Long = &H10        '    The window accepts drag-drop files.
Private Const WS_EX_TRANSPARENT         As Long = &H20        '    The window should not be painted until siblings beneath the window (that were created by the same thread) have been painted. The window appears transparent because the bits of underlying sibling windows have already been painted.To achieve transparency without these restrictions, use the SetWindowRgn function.
Private Const WS_EX_MDICHILD            As Long = &H40        '    The window is a MDI child window.
Private Const WS_EX_TOOLWINDOW          As Long = &H80        '    The window is intended to be used as a floating toolbar. A tool window has a title bar that is shorter than a normal title bar, and the window title is drawn using a smaller font. A tool window does not appear in the taskbar or in the dialog that appears when the user presses ALT+TAB. If a tool window has a system menu, its icon is not displayed on the title bar. However, you can display the system menu by right-clicking or by typing ALT+SPACE.

Private Const WS_EX_WINDOWEDGE          As Long = &H100       '    The window has a border with a raised edge.
Private Const WS_EX_CLIENTEDGE          As Long = &H200       '    The window has a border with a sunken edge.
Private Const WS_EX_CONTEXTHELP         As Long = &H400       '    The title bar of the window includes a question mark. When the user clicks the question mark, the cursor changes to a question mark with a pointer. If the user then clicks a child window, the child receives a WM_HELP message. The child window should pass the message to the parent window procedure, which should call the WinHelp function using the HELP_WM_HELP command. The Help application displays a pop-up window that typically contains help for the child window. WS_EX_CONTEXTHELP cannot be used with the WS_MAXIMIZEBOX or WS_MINIMIZEBOX styles.
Private Const WS_EX_OVERLAPPEDWINDOW    As Long = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)                  '    The window is an overlapped window.
Private Const WS_EX_PALETTEWINDOW       As Long = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST) '    The window is palette window, which is a modeless dialog box that presents an array of commands.

Private Const WS_EX_RIGHT               As Long = &H1000      '    The window has generic "right-aligned" properties. This depends on the window class. This style has an effect only if the shell language is Hebrew, Arabic, or another language that supports reading-order alignment; otherwise, the style is ignored.Using the WS_EX_RIGHT style for static or edit controls has the same effect as using the SS_RIGHT or ES_RIGHT style, respectively. Using this style with button controls has the same effect as using BS_RIGHT and BS_RIGHTBUTTON styles.
Private Const WS_EX_RTLREADING          As Long = &H2000      '    If the shell language is Hebrew, Arabic, or another language that supports reading-order alignment, the window text is displayed using right-to-left reading-order properties. For other languages, the style is ignored.
Private Const WS_EX_LEFTSCROLLBAR       As Long = &H4000      '    If the shell language is Hebrew, Arabic, or another language that supports reading order alignment, the vertical scroll bar (if present) is to the left of the client area. For other languages, the style is ignored.

Private Const WS_EX_CONTROLPARENT       As Long = &H10000     '    The window itself contains child windows that should take part in dialog box navigation. If this style is specified, the dialog manager recurses into children of this window when performing navigation operations such as handling the TAB key, an arrow key, or a keyboard mnemonic.
Private Const WS_EX_STATICEDGE          As Long = &H20000     '    The window has a three-dimensional border style intended to be used for items that do not accept user input.
Private Const WS_EX_APPWINDOW           As Long = &H40000     '    Forces a top-level window onto the taskbar when the window is visible.
Private Const WS_EX_LAYERED             As Long = &H80000     '    The window is a layered window. This style cannot be used if the window has a class style of either CS_OWNDC or CS_CLASSDC. Windows 8: The WS_EX_LAYERED style is supported for top-level windows and child windows. Previous Windows versions support WS_EX_LAYERED only for top-level windows.

Private Const WS_EX_NOINHERITLAYOUT     As Long = &H100000    '    The window does not pass its window layout to its child windows.
Private Const WS_EX_NOREDIRECTIONBITMAP As Long = &H200000    '    The window does not render to a redirection surface. This is for windows that do not have visible content or that use mechanisms other than surfaces to provide their visual.
Private Const WS_EX_LAYOUTRTL           As Long = &H400000    '    If the shell language is Hebrew, Arabic, or another language that supports reading order alignment, the horizontal origin of the window is on the right edge. Increasing horizontal values advance to the left.

Private Const WS_EX_COMPOSITED          As Long = &H2000000   '    Paints all descendants of a window in bottom-to-top painting order using double-buffering. Bottom-to-top painting order allows a descendent window to have translucency (alpha) and transparency (color-key) effects, but only if the descendent window also has the WS_EX_TRANSPARENT bit set. Double-buffering allows the window and its descendents to be painted without flicker. This cannot be used if the window has a class style of either CS_OWNDC or CS_CLASSDC.Windows 2000: This style is not supported.
Private Const WS_EX_NOACTIVATE          As Long = &H8000000   '    A top-level window created with this style does not become the foreground window when the user clicks it. The system does not bring this window to the foreground when the user minimizes or closes the foreground window.
                                                              '    The window should not be activated through programmatic access or via keyboard navigation by accessible technology, such as Narrator.
                                                              '    To activate the window, use the SetActiveWindow or SetForegroundWindow function.
                                                              '    The window does not appear on the taskbar by default. To force the window to appear on the taskbar, use the WS_EX_APPWINDOW style.


Private Sub Form_Load()

End Sub
