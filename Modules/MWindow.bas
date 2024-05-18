Attribute VB_Name = "MWindow"
Option Explicit

Public Enum EWinMsg
    WM_NULL = &H0&                             '   0
    WM_CREATE = &H1&                           '   1
    WM_DESTROY = &H2&                          '   2
    WM_MOVE = &H3&                             '   3
    WM_ERASEBACKGROUND = &H4&                  '   4
    WM_SIZE = &H5&                             '   5
    WM_ACTIVATE = &H6&                         '   6
    WM_SETFOCUS = &H7&                         '   7
    WM_KILLFOCUS = &H8&                        '   8
    WM_ENABLE = &HA&                           '  10
    WM_SETREDRAW = &HB&                        '  11
    WM_SETTEXT = &HC&                          '  12
    WM_GETTEXT = &HD&                          '  13
    WM_GETTEXTLENGTH = &HE&                    '  14
    WM_PAINT = &HF&                            '  15
    WM_CLOSE = &H10&                           '  16
    WM_QUERYENDSESSION = &H11&                 '  17
    WM_QUIT = &H12&                            '  18
    WM_QUERYOPEN = &H13&                       '  19
    WM_ERASEBKGND = &H14&                      '  20
    WM_SYSCOLORCHANGE = &H15&                  '  21
    WM_ENDSESSION = &H16&                      '  22
    WM_SHOWWINDOW = &H18&                      '  24
    WM_WININICHANGE = &H1A&                    '  26
    WM_DEVMODECHANGE = &H1B&                   '  27
    WM_ACTIVATEAPP = &H1C&                     '  28
    WM_FONTCHANGE = &H1D&                      '  29
    WM_TIMECHANGE = &H1E&                      '  30
    WM_CANCELMODE = &H1F&                      '  31
    WM_SETCURSOR = &H20&                       '  32
    WM_MOUSEACTIVATE = &H21&                   '  33
    WM_CHILDACTIVATE = &H22&                   '  34
    WM_QUEUESYNC = &H23&                       '  35
    WM_GETMINMAXINFO = &H24&                   '  36
    WM_PAINTICON = &H26&                       '  38
    WM_ICONERASEBKGND = &H27&                  '  39
    WM_NEXTDLGCTL = &H28&                      '  40
    WM_SPOOLERSTATUS = &H2A&                   '  42
    WM_DRAWITEM = &H2B&                        '  43
    WM_MEASUREITEM = &H2C&                     '  44
    WM_DELETEITEM = &H2D&                      '  45
    WM_VKEYTOITEM = &H2E&                      '  46
    WM_CHARTOITEM = &H2F&                      '  47
    WM_SETFONT = &H30&                         '  48
    WM_GETFONT = &H31&                         '  49
    WM_SETHOTKEY = &H32&                       '  50
    WM_GETHOTKEY = &H33&                       '  51
    WM_QUERYDRAGICON = &H37&                   '  55
    WM_COMPAREITEM = &H39&                     '  57
    WM_GETOBJECT = &H3D&                       '  61
    WM_COMPACTING = &H41&                      '  65
    WM_COMMNOTIFY = &H44&                      '  68
    'WM_WINDOWPOSCHANGING = &H46&               '  70
    WM_WINDOWPOSCHANGED = &H47&                '  71
    WM_POWER = &H48&                           '  72
    WM_COPYDATA = &H4A&                        '  74
    WM_CANCELJOURNAL = &H4B&                   '  75
    WM_NOTIFY = &H4E&                          '  78
    WM_INPUTLANGCHANGEREQUEST = &H50&          '  80
    WM_INPUTLANGCHANGE = &H51&                 '  81
    WM_TCARD = &H52&                           '  82
    WM_HELP = &H53&                            '  83
    WM_USERCHANGED = &H54&                     '  84
    WM_NOTIFYFORMAT = &H55&                    '  85
    WM_CONTEXTMENU = &H7B&                     ' 123
    WM_STYLECHANGING = &H7C&                   ' 124
    WM_STYLECHANGED = &H7D&                    ' 125
    WM_DISPLAYCHANGE = &H7E&                   ' 126
    WM_GETICON = &H7F&                         ' 127
    WM_SETICON = &H80&                         ' 128
    WM_NCCREATE = &H81&                        ' 129
    WM_NCDESTROY = &H82&                       ' 130
    WM_NCCALCSIZE = &H83&                      ' 131
    WM_NCHITTEST = &H84&                       ' 132
    WM_NCPAINT = &H85&                         ' 133
    WM_NCACTIVATE = &H86&                      ' 134
    WM_GETDLGCODE = &H87&                      ' 135
    WM_SYNCPAINT = &H88&                       ' 136
    WM_NCMOUSEMOVE = &HA0&                     ' 160
    WM_NCLBUTTONDOWN = &HA1&                   ' 161
    WM_NCLBUTTONUP = &HA2&                     ' 162
    WM_NCLBUTTONDBLCLK = &HA3&                 ' 163
    WM_NCRBUTTONDOWN = &HA4&                   ' 164
    WM_NCRBUTTONUP = &HA5&                     ' 165
    WM_NCRBUTTONDBLCLK = &HA6&                 ' 166
    WM_NCMBUTTONDOWN = &HA7&                   ' 167
    WM_NCMBUTTONUP = &HA8&                     ' 168
    WM_NCMBUTTONDBLCLK = &HA9&                 ' 169
    WM_NCXBUTTONDOWN = &HAB&                   ' 171
    WM_NCXBUTTONUP = &HAC&                     ' 172
    WM_NCXBUTTONDBLCLK = &HAD&                 ' 173
    WM_INPUT_DEVICE_CHANGE = &HFE&             ' 254
    WM_INPUT = &HFF&                           ' 255
    WM_KEYDOWN = &H100&                        ' 256
    WM_KEYUP = &H101&                          ' 257
    WM_CHAR = &H102&                           ' 258
    WM_DEADCHAR = &H103&                       ' 259
    WM_SYSKEYDOWN = &H104&                     ' 260
    WM_SYSKEYUP = &H105&                       ' 261
    WM_SYSCHAR = &H106&                        ' 262
    WM_SYSDEADCHAR = &H107&                    ' 263
'    WM_KEYLAST = &H108&                        ' 264
    WM_KEYLAST = &H109&                        ' 265
    WM_IME_STARTCOMPOSITION = &H10D&           ' 269
    WM_IME_ENDCOMPOSITION = &H10E&             ' 270
    WM_IME_COMPOSITION = &H10F&                ' 271
    WM_INITDIALOG = &H110&                     ' 272
    WM_COMMAND = &H111&                        ' 273
    WM_SYSCOMMAND = &H112&                     ' 274
    WM_TIMER = &H113&                          ' 275
    WM_HSCROLL = &H114&                        ' 276
    WM_VSCROLL = &H115&                        ' 277
    WM_INITMENU = &H116&                       ' 278
    WM_INITMENUPOPUP = &H117&                  ' 279
    WM_GESTURE = &H119&                        ' 281
    WM_GESTURENOTIFY = &H11A&                  ' 282
    WM_MENUSELECT = &H11F&                     ' 287
    WM_MENUCHAR = &H120&                       ' 288
    WM_ENTERIDLE = &H121&                      ' 289
    WM_MENURBUTTONUP = &H122&                  ' 290
    WM_MENUDRAG = &H123&                       ' 291
    WM_MENUGETOBJECT = &H124&                  ' 292
    WM_UNINITMENUPOPUP = &H125&                ' 293
    WM_MENUCOMMAND = &H126&                    ' 294
    WM_CHANGEUISTATE = &H127&                  ' 295
    WM_UPDATEUISTATE = &H128&                  ' 296
    WM_QUERYUISTATE = &H129&                   ' 297
    WM_CTLCOLORMSGBOX = &H132&                 ' 306
    WM_CTLCOLOREDIT = &H133&                   ' 307
    WM_CTLCOLORLISTBOX = &H134&                ' 308
    WM_CTLCOLORBTN = &H135&                    ' 309
    WM_CTLCOLORDLG = &H136&                    ' 310
    WM_CTLCOLORSCROLLBAR = &H137&              ' 311
    WM_CTLCOLORSTATIC = &H138&                 ' 312
    WM_MOUSEMOVE = &H200&                      ' 512
    WM_LBUTTONDOWN = &H201&                    ' 513
    WM_LBUTTONUP = &H202&                      ' 514
    WM_LBUTTONDBLCLK = &H203&                  ' 515
    WM_RBUTTONDOWN = &H204&                    ' 516
    WM_RBUTTONUP = &H205&                      ' 517
    WM_RBUTTONDBLCLK = &H206&                  ' 518
    WM_MBUTTONDOWN = &H207&                    ' 519
    WM_MBUTTONUP = &H208&                      ' 520
    WM_MBUTTONDBLCLK = &H209&                  ' 521
    WM_MOUSEWHEEL = &H20A&                     ' 522
    WM_XBUTTONDOWN = &H20B&                    ' 523
    WM_XBUTTONUP = &H20C&                      ' 524
    WM_XBUTTONDBLCLK = &H20D&                  ' 525
    WM_MOUSEHWHEEL = &H20E&                    ' 526
    WM_PARENTNOTIFY = &H210&                   ' 528
    WM_ENTERMENULOOP = &H211&                  ' 529
    WM_EXITMENULOOP = &H212&                   ' 530
    WM_NEXTMENU = &H213&                       ' 531
    WM_SIZING = &H214&                         ' 532
    WM_CAPTURECHANGED = &H215&                 ' 533
    WM_MOVING = &H216&                         ' 534
    WM_POWERBROADCAST = &H218&                 ' 536
    WM_DEVICECHANGE = &H219&                   ' 537
    WM_MDICREATE = &H220&                      ' 544
    WM_MDIDESTROY = &H221&                     ' 545
    WM_MDIACTIVATE = &H222&                    ' 546
    WM_MDIRESTORE = &H223&                     ' 547
    WM_MDINEXT = &H224&                        ' 548
    WM_MDIMAXIMIZE = &H225&                    ' 549
    WM_MDITILE = &H226&                        ' 550
    WM_MDICASCADE = &H227&                     ' 551
    WM_MDIICONARRANGE = &H228&                 ' 552
    WM_MDIGETACTIVE = &H229&                   ' 553
    WM_MDISETMENU = &H230&                     ' 560
    WM_ENTERSIZEMOVE = &H231&                  ' 561
    WM_EXITSIZEMOVE = &H232&                   ' 562
    WM_DROPFILES = &H233&                      ' 563
    WM_MDIREFRESHMENU = &H234&                 ' 564
    WM_TOUCH = &H240&                          ' 576
    WM_IME_SETCONTEXT = &H281&                 ' 641
    WM_IME_NOTIFY = &H282&                     ' 642
    WM_IME_CONTROL = &H283&                    ' 643
    WM_IME_COMPOSITIONFULL = &H284&            ' 644
    WM_IME_SELECT = &H285&                     ' 645
    WM_IME_CHAR = &H286&                       ' 646
    WM_IME_REQUEST = &H288&                    ' 648
    WM_IME_KEYDOWN = &H290&                    ' 656
    WM_IME_KEYUP = &H291&                      ' 657
    WM_NCMOUSEHOVER = &H2A0&                   ' 672
    WM_MOUSEHOVER = &H2A1&                     ' 673
    WM_NCMOUSELEAVE = &H2A2&                   ' 674
    WM_MOUSELEAVE = &H2A3&                     ' 675
    WM_WTSSESSION_CHANGE = &H2B1&              ' 689
    WM_TABLET_FIRST = &H2C0&                   ' 704
    WM_TABLET_LAST = &H2DF&                    ' 735
    WM_CUT = &H300&                            ' 768
    WM_COPY = &H301&                           ' 769
    WM_PASTE = &H302&                          ' 770
    WM_CLEAR = &H303&                          ' 771
    WM_UNDO = &H304&                           ' 772
    WM_RENDERFORMAT = &H305&                   ' 773
    WM_RENDERALLFORMATS = &H306&               ' 774
    WM_DESTROYCLIPBOARD = &H307&               ' 775
    WM_DRAWCLIPBOARD = &H308&                  ' 776
    WM_PAINTCLIPBOARD = &H309&                 ' 777
    WM_VSCROLLCLIPBOARD = &H30A&               ' 778
    WM_SIZECLIPBOARD = &H30B&                  ' 779
    WM_ASKCBFORMATNAME = &H30C&                ' 780
    WM_CHANGECBCHAIN = &H30D&                  ' 781
    WM_HSCROLLCLIPBOARD = &H30E&               ' 782
    WM_QUERYNEWPALETTE = &H30F&                ' 783
    WM_PALETTEISCHANGING = &H310&              ' 784
    WM_PALETTECHANGED = &H311&                 ' 785
    WM_HOTKEY = &H312&                         ' 786
    WM_PRINT = &H317&                          ' 791
    WM_PRINTCLIENT = &H318&                    ' 792
    WM_APPCOMMAND = &H319&                     ' 793
    WM_THEMECHANGED = &H31A&                   ' 794
    WM_CLIPBOARDUPDATE = &H31D&                ' 797
    WM_DWMCOMPOSITIONCHANGED = &H31E&          ' 798
    WM_DWMNCRENDERINGCHANGED = &H31F&          ' 799
    WM_DWMCOLORIZATIONCOLORCHANGED = &H320&    ' 800
    WM_DWMWINDOWMAXIMIZEDCHANGE = &H321&       ' 801
    WM_DWMSENDICONICTHUMBNAIL = &H323&         ' 803
    WM_DWMSENDICONICLIVEPREVIEWBITMAP = &H326& ' 806
    WM_GETTITLEBARINFOEX = &H33F&              ' 831
    WM_HANDHELDFIRST = &H358&                  ' 856
    WM_HANDHELDLAST = &H35F&                   ' 863
    WM_AFXFIRST = &H360&                       ' 864
    WM_AFXLAST = &H37F&                        ' 895
    WM_PENWINFIRST = &H380&                    ' 896
    WM_PENWINLAST = &H38F&                     ' 911
    WM_WINDOWPOSCHANGING = &H400&              ' 1024
    'WM_APP = &HCCC&                            ' 3276
    WM_APP = &H8000&                           ' 32768
End Enum


Private Const WA_INACTIVE    As Long = 0&    ' Deaktiviert.
Private Const WA_ACTIVE      As Long = 1&    ' Aktiviert durch eine andere Methode als einen Mausklick (z. B. durch einen Aufruf der SetActiveWindow-Funktion oder durch Die Verwendung der Tastaturschnittstelle zum Auswählen des Fensters).
Private Const WA_CLICKACTIVE As Long = 2&    ' Durch Einen Mausklick aktiviert.

Private Const GWL_WNDPROC    As Long = -4&   ' Ruft die Adresse der Fensterprozedur oder ein Handle ab, das die Adresse der Fensterprozedur darstellt. Sie müssen die CallWindowProc-Funktion verwenden, um die Fensterprozedur aufzurufen.
Private Const GWL_HINSTANCE  As Long = -6&   ' Ruft ein Handle für die anwendung instance ab.
Private Const GWL_HWNDPARENT As Long = -8&   ' Ruft ggf. ein Handle für das übergeordnete Fenster ab.
Private Const GWL_ID         As Long = -12&  ' Ruft den Bezeichner des Fensters ab.
Private Const GWL_STYLE      As Long = -16&  ' Ruft die Fensterstile ab.
Private Const GWL_EXSTYLE    As Long = -20&  ' Ruft die erweiterten Fensterstile ab.
Private Const GWL_USERDATA   As Long = -21&  ' Ruft die dem Fenster zugeordneten Benutzerdaten ab. Diese Daten sind für die Verwendung durch die Anwendung vorgesehen, die das Fenster erstellt hat. Sein Wert ist anfänglich 0 (null).

#If VBA7 Then
    
    'Private Declare PtrSafe Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
    
    Private Declare PtrSafe Function DefWindowProcW Lib "user32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetWindowLongW Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLongW Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
#Else
    'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-translatemessage
    'BOOL TranslateMessage( [in] const MSG *lpMsg )
    'Private Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-defwindowprocw
    Private Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    
    'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-getwindowlongw
    'LONG GetWindowLongW([in] HWND hWnd, [in] int  nIndex);
    Private Declare Function GetWindowLongW Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

#End If

Private m_Atoms   As Collection
Private m_Windows As Collection
'Public WindowAdded As Boolean

Public Function Atoms_Add(ByVal ClassName As String, ByVal Atom As LongPtr) As Boolean
    If m_Atoms Is Nothing Then Set m_Atoms = New Collection
    Atoms_Add = Not MPtr.Col_Contains(m_Atoms, ClassName)
    If Atoms_Add Then m_Atoms.Add Atom, ClassName
End Function

Public Function Atoms_Contains(ByVal ClassName As String) As Boolean
    If m_Atoms Is Nothing Then Set m_Atoms = New Collection
    On Error Resume Next
'  '"Extras->Optionen->Allgemein->Unterbrechen bei Fehlern->Bei nicht verarbeiteten Fehlern"
    If IsEmpty(m_Atoms(ClassName)) Then: 'DoNothing
    Atoms_Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function Atoms_Item(ByVal ClassName As String) As LongPtr
    If m_Atoms Is Nothing Then Set m_Atoms = New Collection
    If MPtr.Col_Contains(m_Atoms, ClassName) Then
        Atoms_Item = m_Atoms.Item(ClassName)
    End If
End Function

Public Function Atoms_Delete(ByVal ClassName As String) As Boolean
    If Atoms_Contains(ClassName) Then m_Atoms.Remove ClassName
End Function

Public Function Windows_Add(hWnd As LongPtr, Window As Window) As Boolean
    If m_Windows Is Nothing Then Set m_Windows = New Collection
    Windows_Add = Not MPtr.Col_Contains(m_Windows, CStr(hWnd))
    If Windows_Add Then m_Windows.Add Window, CStr(hWnd)
End Function

Public Function Windows_Contains(hWnd As LongPtr) As Boolean
    If m_Windows Is Nothing Then Set m_Windows = New Collection
    On Error Resume Next
'  '"Extras->Optionen->Allgemein->Unterbrechen bei Fehlern->Bei nicht verarbeiteten Fehlern"
    If IsEmpty(m_Windows(CStr(hWnd))) Then: 'DoNothing
    Windows_Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function Windows_Delete(hWnd As LongPtr) As Boolean
    If Windows_Contains(CStr(hWnd)) Then m_Windows.Remove CStr(hWnd)
End Function


Public Function LoWord(ByVal lngValue As Long) As Integer
    LoWord = (lngValue And &H7FFF)
End Function

Public Function HiWord(ByVal lngValue As Long) As Integer 'Long
    If (lngValue And &H80000000) = &H80000000 Then
        HiWord = ((lngValue And &H7FFF0000) \ &H10000) Or &H8000
    Else
        HiWord = (lngValue And &HFFFF0000) \ &H10000
    End If
End Function

'LRESULT Wndproc( HWND unnamedParam1, UINT unnamedParam2, WPARAM unnamedParam3, LPARAM unnamedParam4 )
Public Function WndProc(ByVal hWnd_Param1 As LongPtr, ByVal uiMsg_Param2 As EWinMsg, ByVal wParam3 As LongPtr, ByVal lParam4 As LongPtr) As LongPtr
    If Not Windows_Contains(hWnd_Param1) Then
        WndProc = DefWindowProcW(hWnd_Param1, uiMsg_Param2, wParam3, lParam4)
        Exit Function
    End If
    'If m_Windows Is Nothing Then
    '    Set m_Windows = New Collection
    '    WndProc = DefWindowProcW(hWnd_Param1, uiMsg_Param2, wParam3, lParam4)
    '    Exit Function
    'End If
'    If Not Windows_Contains(hWnd_Param1) Then
'        WndProc = DefWindowProcW(hWnd_Param1, uiMsg_Param2, wParam3, lParam4)
'        Exit Function
'    End If
    Dim Window As Window: Set Window = m_Windows.Item(CStr(hWnd_Param1))
    'Debug.Print WindowMessage_ToStr(uiMsg_Param2)
    Dim Cancel As Integer
    Select Case uiMsg_Param2
    Case WM_ACTIVATE:
        'Debug.Print "wparam3: " & wParam3 & " " & (wParam3 And &HFFFF)
                           If (wParam3 And &HFFFF) = WA_INACTIVE Then Window.OnDeactivate Else Window.OnActivate
    'Case WM_MOUSEACTIVATE: 'If (wParam3 And &HFFFF) = WA_INACTIVE Then Window.OnDeactivate Else
    '                        Window.OnActivate
    Case WM_CREATE:         Window.OnLoad
                            'Window.OnInitialize
    
    'Closing-Events
    Case WM_CLOSE
        Window.OnUnload Cancel
        If Cancel Then
            'Nein nicht nope
        End If
    
    Case WM_DESTROY:        Window.OnUnload Cancel
    
    Case WM_QUIT:           Window.OnUnload Cancel
    
    Case WM_DROPFILES:      'Window.OnDragDrop
    Case WM_GESTURE
    Case WM_GESTURENOTIFY:  'Window.onGesture
    
    'Keyboard-Events
    Case WM_KEYDOWN:        Window.OnKeyDown LoWord(wParam3), LoWord(lParam4)
    Case WM_CHAR:           Window.OnKeyPress LoWord(wParam3)
    'Case WM_KEYLAST:        Window.OnKeyPress
    Case WM_KEYUP:          Window.OnKeyUp LoWord(wParam3), LoWord(lParam4)
    'Mouse-Events
    Case WM_MOUSEMOVE:      Dim tmp As Integer: tmp = LoWord(wParam3)
                            Window.OnMouseMove tmp, tmp, CSng(LoWord(lParam4)), CSng(HiWord(lParam4))
    Case WM_LBUTTONDBLCLK:  Window.OnDblClick
    
    Case WM_LBUTTONDOWN:    Window.OnMouseDown MouseButtonConstants.vbLeftButton, CInt(LoWord(wParam3)), CSng(LoWord(lParam4)), CSng(HiWord(lParam4))
    Case WM_LBUTTONUP:      Window.OnMouseUp MouseButtonConstants.vbLeftButton, CInt(LoWord(wParam3)), CSng(LoWord(lParam4)), CSng(HiWord(lParam4))
                            Window.OnClick
                            Window.OnMouseMove MouseButtonConstants.vbLeftButton, CInt(LoWord(wParam3)), CSng(LoWord(lParam4)), CSng(HiWord(lParam4))
    'Case WM_MBUTTONDBLCLK:
    Case WM_MBUTTONDOWN:    Window.OnMouseDown MouseButtonConstants.vbMiddleButton, CInt(LoWord(wParam3)), CSng(LoWord(lParam4)), CSng(HiWord(lParam4))
    Case WM_MBUTTONUP:      Window.OnMouseUp MouseButtonConstants.vbMiddleButton, CInt(LoWord(wParam3)), CSng(LoWord(lParam4)), CSng(HiWord(lParam4))
    'Case WM_RBUTTONDBLCLK
    Case WM_RBUTTONDOWN:    Window.OnMouseDown MouseButtonConstants.vbRightButton, CInt(LoWord(wParam3)), CSng(LoWord(lParam4)), CSng(HiWord(lParam4))
    Case WM_RBUTTONUP:      Window.OnMouseUp MouseButtonConstants.vbRightButton, CInt(LoWord(wParam3)), CSng(LoWord(lParam4)), CSng(HiWord(lParam4))
    
    Case WM_SETFOCUS:       Window.OnGotFocus
    Case WM_KILLFOCUS:      Window.OnLostFocus
    Case WM_HSCROLL:        Window.OnScrollH
    Case WM_VSCROLL:        Window.OnScrollV

    'Case WM_MOUSEHOVER:
    'Case WM_MOUSEHWHEEL:
    'Case WM_MOUSEWHEEL:
    'Case WM_MOUSELEAVE:
    'Case WM_SIZE:           Window.OnResize
    'Case WM_MOVE:           Window.OnResize
    'Case WM_MOVING:         Window.OnResize
    Case WM_PAINT:          Window.OnPaint
    Case WM_WINDOWPOSCHANGED: Window.OnResize
                            Exit Function
    
    
    
    'Case 0: Debug.Print "WM_NULL"
    'Case 2: PostQuitMessage 0 ': Exit Sub
    'Case Else: WndProc = DefWindowProcW(hWnd_Param1, uiMsg_Param2, wParam3, lParam4)
    End Select
    
    WndProc = DefWindowProcW(hWnd_Param1, uiMsg_Param2, wParam3, lParam4)
End Function

Public Function EWinMsg_ToStr(ByVal WM_ As Long) As String
    Dim s As String
    Select Case WM_
    Case &H0&: s = "WM_NULL"
    Case &H1&: s = "WM_CREATE"
    Case &H2&: s = "WM_DESTROY"
    Case &H3&: s = "WM_MOVE"
    Case &H4&: s = "WM_ERASEBACKGROUND"
    Case &H5&: s = "WM_SIZE"
    Case &H6&: s = "WM_ACTIVATE"
    Case &H7&: s = "WM_SETFOCUS"
    Case &H8&: s = "WM_KILLFOCUS"
    Case &HA&: s = "WM_ENABLE"
    Case &HB&: s = "WM_SETREDRAW"
    Case &HC&: s = "WM_SETTEXT"
    Case &HD&: s = "WM_GETTEXT"
    Case &HE&: s = "WM_GETTEXTLENGTH"
    Case &HF&: s = "WM_PAINT"
    Case &H10&: s = "WM_CLOSE"
    Case &H11&: s = "WM_QUERYENDSESSION"
    Case &H12&: s = "WM_QUIT"
    Case &H13&: s = "WM_QUERYOPEN"
    Case &H14&: s = "WM_ERASEBKGND"
    Case &H15&: s = "WM_SYSCOLORCHANGE"
    Case &H16&: s = "WM_ENDSESSION"
    Case &H18&: s = "WM_SHOWWINDOW"
    Case &H1A&: s = "WM_WININICHANGE"
    Case &H1A&: s = "WM_SETTINGCHANGE"
    Case &H1B&: s = "WM_DEVMODECHANGE"
    Case &H1C&: s = "WM_ACTIVATEAPP"
    Case &H1D&: s = "WM_FONTCHANGE"
    Case &H1E&: s = "WM_TIMECHANGE"
    Case &H1F&: s = "WM_CANCELMODE"
    'Case &H20&: s = "WM_NCCALCSIZE"
    Case &H20&: s = "WM_SETCURSOR"
    Case &H21&: s = "WM_MOUSEACTIVATE"
    Case &H22&: s = "WM_CHILDACTIVATE"
    Case &H23&: s = "WM_QUEUESYNC"
    Case &H24&: s = "WM_GETMINMAXINFO"
    '
    Case &H26&: s = "WM_PAINTICON"
    Case &H27&: s = "WM_ICONERASEBKGND"
    Case &H28&: s = "WM_NEXTDLGCTL"
    '
    Case &H2A&: s = "WM_SPOOLERSTATUS"
    Case &H2B&: s = "WM_DRAWITEM"
    Case &H2C&: s = "WM_MEASUREITEM"
    Case &H2D&: s = "WM_DELETEITEM"
    Case &H2E&: s = "WM_VKEYTOITEM"
    Case &H2F&: s = "WM_CHARTOITEM"
    Case &H30&: s = "WM_SETFONT"
    Case &H31&: s = "WM_GETFONT"
    Case &H32&: s = "WM_SETHOTKEY"
    Case &H33&: s = "WM_GETHOTKEY"
    '
    Case &H37&: s = "WM_QUERYDRAGICON"
    '
    Case &H39&: s = "WM_COMPAREITEM"
    '
    Case &H3D&: s = "WM_GETOBJECT"
    '
    Case &H41&: s = "WM_COMPACTING"
    '
    Case &H44&: s = "WM_COMMNOTIFY"
    '
    Case &H46&: s = "WM_WINDOWPOSCHANGING"
    Case &H47&: s = "WM_WINDOWPOSCHANGED"
    Case &H48&: s = "WM_POWER"
    '
    Case &H4A&: s = "WM_COPYDATA"
    Case &H4B&: s = "WM_CANCELJOURNAL"
    '
    Case &H4E&: s = "WM_NOTIFY"
    '
    Case &H50&: s = "WM_INPUTLANGCHANGEREQUEST"
    Case &H51&: s = "WM_INPUTLANGCHANGE"
    Case &H52&: s = "WM_TCARD"
    Case &H53&: s = "WM_HELP"
    Case &H54&: s = "WM_USERCHANGED"
    Case &H55&: s = "WM_NOTIFYFORMAT"
    '
    Case &H7B&: s = "WM_CONTEXTMENU"
    Case &H7C&: s = "WM_STYLECHANGING"
    Case &H7D&: s = "WM_STYLECHANGED"
    Case &H7E&: s = "WM_DISPLAYCHANGE"
    Case &H7F&: s = "WM_GETICON"
    'Case &H80&: s = "WM_CHAR"
    Case &H80&: s = "WM_SETICON"
    Case &H81&: s = "WM_NCCREATE"
    Case &H82&: s = "WM_NCDESTROY"
    Case &H83&: s = "WM_NCCALCSIZE"
    Case &H84&: s = "WM_NCHITTEST"
    Case &H85&: s = "WM_NCPAINT"
    Case &H86&: s = "WM_NCACTIVATE"
    Case &H87&: s = "WM_GETDLGCODE"
    Case &H88&: s = "WM_SYNCPAINT"
    '
    Case &HA0&: s = "WM_NCMOUSEMOVE"
    Case &HA1&: s = "WM_NCLBUTTONDOWN"
    Case &HA2&: s = "WM_NCLBUTTONUP"
    Case &HA3&: s = "WM_NCLBUTTONDBLCLK"
    Case &HA4&: s = "WM_NCRBUTTONDOWN"
    Case &HA5&: s = "WM_NCRBUTTONUP"
    Case &HA6&: s = "WM_NCRBUTTONDBLCLK"
    Case &HA7&: s = "WM_NCMBUTTONDOWN"
    Case &HA8&: s = "WM_NCMBUTTONUP"
    Case &HA9&: s = "WM_NCMBUTTONDBLCLK"
    '
    Case &HAB&: s = "WM_NCXBUTTONDOWN"
    Case &HAC&: s = "WM_NCXBUTTONUP"
    Case &HAD&: s = "WM_NCXBUTTONDBLCLK"
    '
    Case &HFE&: s = "WM_INPUT_DEVICE_CHANGE"
    Case &HFF&: s = "WM_INPUT"
    'Case &H100&: s = "WM_KEYFIRST"
    'Case &H100&: s = "WM_ENTERIDLE"
    Case &H100&: s = "WM_KEYDOWN"
    Case &H101&: s = "WM_KEYUP"
    Case &H102&: s = "WM_CHAR"
    Case &H103&: s = "WM_DEADCHAR"
    Case &H104&: s = "WM_SYSKEYDOWN"
    Case &H105&: s = "WM_SYSKEYUP"
    Case &H106&: s = "WM_SYSCHAR"
    Case &H107&: s = "WM_SYSDEADCHAR"
    Case &H108&: s = "WM_KEYLAST"
    Case &H109&: s = "WM_KEYLAST"
    'Case &H109&: s = "WM_UNICHAR"
    '
    Case &H10D&: s = "WM_IME_STARTCOMPOSITION"
    Case &H10E&: s = "WM_IME_ENDCOMPOSITION"
    Case &H10F&: s = "WM_IME_COMPOSITION"
    Case &H10F&: s = "WM_IME_KEYLAST"
    Case &H110&: s = "WM_INITDIALOG"
    Case &H111&: s = "WM_COMMAND"
    Case &H112&: s = "WM_SYSCOMMAND"
    Case &H113&: s = "WM_TIMER"
    Case &H114&: s = "WM_HSCROLL"
    Case &H115&: s = "WM_VSCROLL"
    Case &H116&: s = "WM_INITMENU"
    Case &H117&: s = "WM_INITMENUPOPUP"
    '
    Case &H119&: s = "WM_GESTURE"
    Case &H11A&: s = "WM_GESTURENOTIFY"
    '
    Case &H11F&: s = "WM_MENUSELECT"
    Case &H120&: s = "WM_MENUCHAR"
    Case &H121&: s = "WM_ENTERIDLE"
    Case &H122&: s = "WM_MENURBUTTONUP"
    Case &H123&: s = "WM_MENUDRAG"
    Case &H124&: s = "WM_MENUGETOBJECT"
    Case &H125&: s = "WM_UNINITMENUPOPUP"
    Case &H126&: s = "WM_MENUCOMMAND"
    Case &H127&: s = "WM_CHANGEUISTATE"
    Case &H128&: s = "WM_UPDATEUISTATE"
    Case &H129&: s = "WM_QUERYUISTATE"
    '
    Case &H132&: s = "WM_CTLCOLORMSGBOX"
    Case &H133&: s = "WM_CTLCOLOREDIT"
    Case &H134&: s = "WM_CTLCOLORLISTBOX"
    Case &H135&: s = "WM_CTLCOLORBTN"
    Case &H136&: s = "WM_CTLCOLORDLG"
    Case &H137&: s = "WM_CTLCOLORSCROLLBAR"
    Case &H138&: s = "WM_CTLCOLORSTATIC"
    '
    'Case &H200&: s = "WM_MOUSEFIRST"
    Case &H200&: s = "WM_MOUSEMOVE"
    Case &H201&: s = "WM_LBUTTONDOWN"
    Case &H202&: s = "WM_LBUTTONUP"
    Case &H203&: s = "WM_LBUTTONDBLCLK"
    Case &H204&: s = "WM_RBUTTONDOWN"
    Case &H205&: s = "WM_RBUTTONUP"
    Case &H206&: s = "WM_RBUTTONDBLCLK"
    Case &H207&: s = "WM_MBUTTONDOWN"
    Case &H208&: s = "WM_MBUTTONUP"
    Case &H209&: s = "WM_MBUTTONDBLCLK"
    'Case &H209&: s = "WM_MOUSELAST(95)"
    Case &H209&: s = "WM_MOUSELAST"
    Case &H20A&: s = "WM_MOUSEWHEEL"
    'Case &H20A&: s = "WM_MOUSELAST(NT4,98)"
    Case &H20A&: s = "WM_MOUSELAST"
    Case &H20B&: s = "WM_XBUTTONDOWN"
    Case &H20C&: s = "WM_XBUTTONUP"
    Case &H20D&: s = "WM_XBUTTONDBLCLK"
    Case &H20E&: s = "WM_MOUSEHWHEEL"
    'Case &H20E&: s = "WM_MOUSELAST"
    Case &H20D&: s = "WM_MOUSELAST"
    'Case &H20D&: s = "WM_MOUSELAST(2K,XP,2k3)"
    Case &H210&: s = "WM_PARENTNOTIFY"
    Case &H211&: s = "WM_ENTERMENULOOP"
    Case &H212&: s = "WM_EXITMENULOOP"
    Case &H213&: s = "WM_NEXTMENU"
    Case &H214&: s = "WM_SIZING"
    Case &H215&: s = "WM_CAPTURECHANGED"
    Case &H216&: s = "WM_MOVING"
    '
    Case &H218&: s = "WM_POWERBROADCAST"
    Case &H219&: s = "WM_DEVICECHANGE"
    Case &H220&: s = "WM_MDICREATE"
    Case &H221&: s = "WM_MDIDESTROY"
    Case &H222&: s = "WM_MDIACTIVATE"
    Case &H223&: s = "WM_MDIRESTORE"
    Case &H224&: s = "WM_MDINEXT"
    Case &H225&: s = "WM_MDIMAXIMIZE"
    Case &H226&: s = "WM_MDITILE"
    Case &H227&: s = "WM_MDICASCADE"
    Case &H228&: s = "WM_MDIICONARRANGE"
    Case &H229&: s = "WM_MDIGETACTIVE"
    Case &H230&: s = "WM_MDISETMENU"
    Case &H231&: s = "WM_ENTERSIZEMOVE"
    Case &H232&: s = "WM_EXITSIZEMOVE"
    Case &H233&: s = "WM_DROPFILES"
    Case &H234&: s = "WM_MDIREFRESHMENU"
    '
    Case &H240&: s = "WM_TOUCH"
    '
    Case &H281&: s = "WM_IME_SETCONTEXT"
    Case &H282&: s = "WM_IME_NOTIFY"
    Case &H283&: s = "WM_IME_CONTROL"
    Case &H284&: s = "WM_IME_COMPOSITIONFULL"
    Case &H285&: s = "WM_IME_SELECT"
    Case &H286&: s = "WM_IME_CHAR"
    '
    Case &H288&: s = "WM_IME_REQUEST"
    '
    Case &H290&: s = "WM_IME_KEYDOWN"
    Case &H291&: s = "WM_IME_KEYUP"
    '
    Case &H2A0&: s = "WM_NCMOUSEHOVER"
    Case &H2A1&: s = "WM_MOUSEHOVER"
    Case &H2A2&: s = "WM_NCMOUSELEAVE"
    Case &H2A3&: s = "WM_MOUSELEAVE"
    
    Case &H2B1&: s = "WM_WTSSESSION_CHANGE"
    '
    Case &H2C0&: s = "WM_TABLET_FIRST"
    '
    Case &H2DF&: s = "WM_TABLET_LAST"
    '
    Case &H300&: s = "WM_CUT"
    Case &H301&: s = "WM_COPY"
    Case &H302&: s = "WM_PASTE"
    Case &H303&: s = "WM_CLEAR"
    Case &H304&: s = "WM_UNDO"
    Case &H305&: s = "WM_RENDERFORMAT"
    Case &H306&: s = "WM_RENDERALLFORMATS"
    Case &H307&: s = "WM_DESTROYCLIPBOARD"
    Case &H308&: s = "WM_DRAWCLIPBOARD"
    Case &H309&: s = "WM_PAINTCLIPBOARD"
    Case &H30A&: s = "WM_VSCROLLCLIPBOARD"
    Case &H30B&: s = "WM_SIZECLIPBOARD"
    Case &H30C&: s = "WM_ASKCBFORMATNAME"
    Case &H30D&: s = "WM_CHANGECBCHAIN"
    Case &H30E&: s = "WM_HSCROLLCLIPBOARD"
    Case &H30F&: s = "WM_QUERYNEWPALETTE"
    Case &H310&: s = "WM_PALETTEISCHANGING"
    Case &H311&: s = "WM_PALETTECHANGED"
    Case &H312&: s = "WM_HOTKEY"
    '
    Case &H317&: s = "WM_PRINT"
    Case &H318&: s = "WM_PRINTCLIENT"
    Case &H319&: s = "WM_APPCOMMAND"
    Case &H31A&: s = "WM_THEMECHANGED"
    '
    Case &H31D&: s = "WM_CLIPBOARDUPDATE"
    Case &H31E&: s = "WM_DWMCOMPOSITIONCHANGED"
    Case &H31F&: s = "WM_DWMNCRENDERINGCHANGED"
    Case &H320&: s = "WM_DWMCOLORIZATIONCOLORCHANGED"
    Case &H321&: s = "WM_DWMWINDOWMAXIMIZEDCHANGE"
    '
    Case &H323&: s = "WM_DWMSENDICONICTHUMBNAIL"
    '
    Case &H326&: s = "WM_DWMSENDICONICLIVEPREVIEWBITMAP"
    '
    Case &H33F&: s = "WM_GETTITLEBARINFOEX"
    '
    Case &H358&: s = "WM_HANDHELDFIRST"
    '
    Case &H35F&: s = "WM_HANDHELDLAST"
    Case &H360&: s = "WM_AFXFIRST"
    '
    Case &H37F&: s = "WM_AFXLAST"
    Case &H380&: s = "WM_PENWINFIRST"
    '
    Case &H38F&: s = "WM_PENWINLAST"
    'Case &H400&: s = "WM_USER"
    Case &H400&: s = "WM_WINDOWPOSCHANGING"
    Case &HCCC&: s = "WM_APP"
    Case &H8000&: s = "WM_APP"
    'Case Else: s = CStr(WM_)
    End Select
    EWinMsg_ToStr = s
End Function

Public Function GetEnumWM() As String
    Dim i As Long, u As Long: u = &H8000&
    ReDim sa(0 To u) As String
    Dim mx As Long
    For i = 0 To u
        sa(i) = EWinMsg_ToStr(i)
        mx = Max(mx, Len(sa(i)))
    Next
    Dim sEnum As String: sEnum = "Public Enum EWinMsg" & vbCrLf
    Dim s As String, sp As String
    For i = 0 To UBound(sa)
        s = sa(i): sp = Space(mx - Len(s) + 4)
        If Len(s) Then
            sEnum = sEnum & "    " & s & " = &H" & Hex(i) & "&" & sp & "' " & CStr(i) & vbCrLf
        End If
    Next
    sEnum = sEnum & "End Enum" & vbCrLf
    Clipboard.Clear
    Clipboard.SetText sEnum
End Function

'    Private Declare Function GetWindowLongW Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
'    Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Const GWL_WNDPROC    As Long = –4  '    Ruft die Adresse der Fensterprozedur oder ein Handle ab, das die Adresse der Fensterprozedur darstellt. Sie müssen die CallWindowProc-Funktion verwenden, um die Fensterprozedur aufzurufen.
'Private Const GWL_HINSTANCE  As Long = –6  '    Ruft ein Handle für die anwendung instance ab.
'Private Const GWL_HWNDPARENT As Long = -8  '    Ruft ggf. ein Handle für das übergeordnete Fenster ab.
'Private Const GWL_ID         As Long = -12 '    Ruft den Bezeichner des Fensters ab.
'Private Const GWL_STYLE      As Long = -16 '    Ruft die Fensterstile ab.
'Private Const GWL_EXSTYLE    As Long = -20 '    Ruft die erweiterten Fensterstile ab.
'Private Const GWL_USERDATA   As Long = -21 '    Ruft die dem Fenster zugeordneten Benutzerdaten ab. Diese Daten sind für die Verwendung durch die Anwendung vorgesehen, die das Fenster erstellt hat. Sein Wert ist anfänglich 0 (null).

Public Property Get WindowStyle(ByVal ahWnd As LongPtr) As Long
    WindowStyle = GetWindowLongW(ahWnd, GWL_STYLE)
End Property
Public Property Let WindowStyle(ByVal ahWnd As LongPtr, ByVal Value As Long)
    Dim hr As Long: hr = SetWindowLongW(ahWnd, GWL_STYLE, Value)
End Property

Public Property Get WindowStyleEx(ByVal ahWnd As LongPtr) As Long
    WindowStyleEx = GetWindowLongW(ahWnd, GWL_EXSTYLE)
End Property
Public Property Let WindowStyleEx(ByVal ahWnd As LongPtr, ByVal Value As Long)
    Dim hr As Long: hr = SetWindowLongW(ahWnd, GWL_EXSTYLE, Value)
End Property


Public Function EWndStyle_ToStr(e As EWndStyle) As String
    Dim sOr As String: sOr = " Or "
    Dim s As String
    If e And WS_TILED Then s = s & IIf(Len(s), sOr, "") & "WS_TILED"
    If e And WS_OVERLAPPED Then s = s & IIf(Len(s), sOr, "") & "WS_OVERLAPPED"

    If e And WS_MAXIMIZEBOX Then s = s & IIf(Len(s), sOr, "") & "WS_MAXIMIZEBOX"
    If e And WS_TABSTOP Then s = s & IIf(Len(s), sOr, "") & "WS_TABSTOP"
    If e And WS_GROUP Then s = s & IIf(Len(s), sOr, "") & "WS_GROUP"
    If e And WS_MINIMIZEBOX Then s = s & IIf(Len(s), sOr, "") & "WS_MINIMIZEBOX"
    If e And WS_SIZEBOX Then s = s & IIf(Len(s), sOr, "") & "WS_SIZEBOX"
    If e And WS_THICKFRAME Then s = s & IIf(Len(s), sOr, "") & "WS_THICKFRAME"
    If e And WS_SYSMENU Then s = s & IIf(Len(s), sOr, "") & "WS_SYSMENU"

    If e And WS_HSCROLL Then s = s & IIf(Len(s), sOr, "") & "WS_HSCROLL"
    If e And WS_VSCROLL Then s = s & IIf(Len(s), sOr, "") & "WS_VSCROLL"
    If e And WS_DLGFRAME Then s = s & IIf(Len(s), sOr, "") & "WS_DLGFRAME"
    If e And WS_BORDER Then s = s & IIf(Len(s), sOr, "") & "WS_BORDER"
    If e And WS_CAPTION Then s = s & IIf(Len(s), sOr, "") & "WS_CAPTION"

    If e And WS_MAXIMIZE Then s = s & IIf(Len(s), sOr, "") & "WS_MAXIMIZE"
    If e And WS_CLIPCHILDREN Then s = s & IIf(Len(s), sOr, "") & "WS_CLIPCHILDREN"
    If e And WS_CLIPSIBLINGS Then s = s & IIf(Len(s), sOr, "") & "WS_CLIPSIBLINGS"
    If e And WS_DISABLED Then s = s & IIf(Len(s), sOr, "") & "WS_DISABLED"

    If e And WS_VISIBLE Then s = s & IIf(Len(s), sOr, "") & "WS_VISIBLE"
    If e And WS_ICONIC Then s = s & IIf(Len(s), sOr, "") & "WS_ICONIC"
    If e And WS_MINIMIZE Then s = s & IIf(Len(s), sOr, "") & "WS_MINIMIZE"
    If e And WS_CHILD Then s = s & IIf(Len(s), sOr, "") & "WS_CHILD"
    If e And WS_CHILDWINDOW Then s = s & IIf(Len(s), sOr, "") & "WS_CHILDWINDOW"
    If e And WS_POPUP Then s = s & IIf(Len(s), sOr, "") & "WS_POPUP"
    EWndStyle_ToStr = s
End Function

Public Function EWndStyleEx_ToStr(e As EWndStyleEx) As String
    Dim sOr As String: sOr = " Or "
    Dim s As String
    If e And WS_EX_LEFT Then s = s & IIf(Len(s), sOr, "") & "WS_EX_LEFT"
    If e And WS_EX_LTRREADING Then s = s & IIf(Len(s), sOr, "") & "WS_EX_LTRREADING"
    If e And WS_EX_RIGHTSCROLLBAR Then s = s & IIf(Len(s), sOr, "") & "WS_EX_RIGHTSCROLLBAR"
    If e And WS_EX_DLGMODALFRAME Then s = s & IIf(Len(s), sOr, "") & "WS_EX_DLGMODALFRAME"

    If e And WS_EX_NOPARENTNOTIFY Then s = s & IIf(Len(s), sOr, "") & "WS_EX_NOPARENTNOTIFY"
    If e And WS_EX_TOPMOST Then s = s & IIf(Len(s), sOr, "") & "WS_EX_TOPMOST"

    If e And WS_EX_ACCEPTFILES Then s = s & IIf(Len(s), sOr, "") & "WS_EX_ACCEPTFILES"
    If e And WS_EX_TRANSPARENT Then s = s & IIf(Len(s), sOr, "") & "WS_EX_TRANSPARENT"
    If e And WS_EX_MDICHILD Then s = s & IIf(Len(s), sOr, "") & "WS_EX_MDICHILD"
    If e And WS_EX_TOOLWINDOW Then s = s & IIf(Len(s), sOr, "") & "WS_EX_TOOLWINDOW"

    If e And WS_EX_WINDOWEDGE Then s = s & IIf(Len(s), sOr, "") & "WS_EX_WINDOWEDGE"
    If e And WS_EX_CLIENTEDGE Then s = s & IIf(Len(s), sOr, "") & "WS_EX_CLIENTEDGE"
    If e And WS_EX_CONTEXTHELP Then s = s & IIf(Len(s), sOr, "") & "WS_EX_CONTEXTHELP"

    If e And WS_EX_RIGHT Then s = s & IIf(Len(s), sOr, "") & "WS_EX_RIGHT"
    If e And WS_EX_RTLREADING Then s = s & IIf(Len(s), sOr, "") & "WS_EX_RTLREADING"
    If e And WS_EX_LEFTSCROLLBAR Then s = s & IIf(Len(s), sOr, "") & "WS_EX_LEFTSCROLLBAR"

    If e And WS_EX_CONTROLPARENT Then s = s & IIf(Len(s), sOr, "") & "WS_EX_CONTROLPARENT"
    If e And WS_EX_STATICEDGE Then s = s & IIf(Len(s), sOr, "") & "WS_EX_STATICEDGE"
    If e And WS_EX_APPWINDOW Then s = s & IIf(Len(s), sOr, "") & "WS_EX_APPWINDOW"
    If e And WS_EX_LAYERED Then s = s & IIf(Len(s), sOr, "") & "WS_EX_LAYERED"

    If e And WS_EX_NOINHERITLAYOUT Then s = s & IIf(Len(s), sOr, "") & "WS_EX_NOINHERITLAYOUT"
    If e And WS_EX_NOREDIRECTIONBITMAP Then s = s & IIf(Len(s), sOr, "") & "WS_EX_NOREDIRECTIONBITMAP"
    If e And WS_EX_LAYOUTRTL Then s = s & IIf(Len(s), sOr, "") & "WS_EX_LAYOUTRTL"

    If e And WS_EX_COMPOSITED Then s = s & IIf(Len(s), sOr, "") & "WS_EX_COMPOSITED"
    If e And WS_EX_NOACTIVATE Then s = s & IIf(Len(s), sOr, "") & "WS_EX_NOACTIVATE"
    EWndStyleEx_ToStr = s
End Function

