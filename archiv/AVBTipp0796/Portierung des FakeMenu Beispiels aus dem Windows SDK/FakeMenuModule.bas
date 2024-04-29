Attribute VB_Name = "FakeMenuModule"
'//——————————————————————————————————————————————————————————————————————————————
'//
'// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
'// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
'// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
'// PARTICULAR PURPOSE.
'//
'// Copyright (c) Microsoft Corporation. All rights reserved.
'//
'//——————————————————————————————————————————————————————————————————————————————
'//
'//      Overview
'//
'//      Normally, pop-up windows receive activation, resulting in the
'//      owner window being de-activated.  To prevent the owner from
'//      being de-activated, the pop-up window should not receive
'//      activation.
'//
'//      Since the pop-up window is not active, input messages are not
'//      delivered to the pop-up.  Instead, the input messages must be
'//      explicitly inspected by the message loop.
'//
'//      Our sample program illustrates how you can create a pop-up
'//      window that contains a selection of colors.
'//
'//      Right-click in the window to change its background color
'//      via the fake menu popup.  Observe
'//
'//      -   The caption of the main application window remains
'//          highlighted even though the fake-menu is "active".
'//
'//      -   The current fake-menu item highlight follows the mouse.
'//
'//      -   The keyboard arrows can be used to move the highlight,
'//          ESC cancels the fake-menu, Enter accepts the fake-menu.
'//
'//      -   The fake-menu appears on the correct monitor (for
'//          multiple-monitor systems).
'//
'//——————————————————————————————————————————————————————————————————————————————
Option Explicit


Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left                    As Long
    Top                     As Long
    Right                   As Long
    Bottom                  As Long
End Type

Private Type CREATESTRUCT
    lpCreateParams          As Long
    hInstance               As Long
    hMenu                   As Long
    hWndParent              As Long
    cy                      As Long
    cx                      As Long
    y                       As Long
    x                       As Long
    style                   As Long
    lpszName                As Long 'Ptr to String
    lpszClass               As Long 'Ptr to String
    ExStyle                 As Long
End Type

Private Type PAINTSTRUCT
    hDC                     As Long
    fErase                  As Long
    rcPaint                 As RECT
    fRestore                As Long
    fIncUpdate              As Long
    rgbReserved(0 To 31)    As Byte
End Type

Private Type MONITORINFO
    cbSize                  As Long
    rcMonitor               As RECT
    rcWork                  As RECT
    dwFlags                 As Long
End Type

Private Type MSG
    hWnd                    As Long
    message                 As Long
    wParam                  As Long
    lParam                  As Long
    time                    As Long
    pt                      As POINTAPI
End Type

Private Type WNDCLASSEX
    cbSize                  As Long
    style                   As Long
    lpfnWndProc             As Long
    cbClsExtra              As Long
    cbWndExtra              As Long
    hInstance               As Long
    hIcon                   As Long
    hCursor                 As Long
    hbrBackground           As Long
    lpszMenuName            As String
    lpszClassName           As String
    hIconSm                 As Long
End Type


Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function BeginPaint Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function InflateRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function EndPaint Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long

Private Declare Function InvalidateRect Lib "user32.dll" _
 (ByVal hWnd As Long, _
  ByVal lpRect As Long, _
  ByVal bErase As Long) As Long   'lpRect as long, damit NULL übergeben werden kann

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" _
 (ByVal hWnd As Long, _
  ByVal nIndex As Long) As Long

Private Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function MonitorFromPoint Lib "user32.dll" (ByVal x As Long, ByVal y As Long, ByVal dwFlags As Long) As Long
Private Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long

Private Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function AdjustWindowRectEx Lib "user32.dll" (ByRef lpRect As RECT, ByVal dsStyle As Long, ByVal bMenu As Long, ByVal dwEsStyle As Long) As Long
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetCapture Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetMessage Lib "user32.dll" Alias "GetMessageA" (ByRef lpMsg As MSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
Private Declare Function IsChild Lib "user32.dll" (ByVal hWndParent As Long, ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32.dll" () As Long
Private Declare Function MapWindowPoints Lib "user32.dll" (ByVal hwndFrom As Long, ByVal hwndTo As Long, ByVal lppt As Long, ByVal cPoints As Long) As Long
Private Declare Function TranslateMessage Lib "user32.dll" (ByRef lpMsg As MSG) As Long
Private Declare Function DispatchMessage Lib "user32.dll" Alias "DispatchMessageA" (ByRef lpMsg As MSG) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
'Private Declare Sub PostQuitMessage Lib "user32.dll" (ByVal nExitCode As Long)

Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long

Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Private Declare Function RegisterClassEx Lib "user32.dll" Alias "RegisterClassExA" (ByRef pcWndClassEx As WNDCLASSEX) As Integer

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
 (ByVal Destination As Long, _
  ByVal Source As Long, _
  ByVal Length As Long)


Private Const GWLP_USERDATA             As Long = -21

Private Const SW_SHOWNOACTIVATE         As Long = 4
Private Const SW_SHOWNORMAL             As Long = 1

Private Const COLOR_HIGHLIGHT           As Long = 13
Private Const SM_CXEDGE                 As Long = 45
Private Const SM_CYEDGE                 As Long = 46

Private Const VK_ESCAPE                 As Long = &H1B
Private Const VK_RETURN                 As Long = &HD
Private Const VK_UP                     As Long = &H26
Private Const VK_DOWN                   As Long = &H28

Private Const WM_CREATE                 As Long = &H1
Private Const WM_MOUSEACTIVATE          As Long = &H21
Private Const WM_PAINT                  As Long = &HF&

Private Const WM_MOUSEMOVE              As Long = &H200
Private Const WM_LBUTTONDOWN            As Long = &H201
Private Const WM_LBUTTONUP              As Long = &H202
Private Const WM_LBUTTONDBLCLK          As Long = &H203
Private Const WM_MBUTTONDOWN            As Long = &H207
Private Const WM_MBUTTONUP              As Long = &H208
Private Const WM_MBUTTONDBLCLK          As Long = &H209
Private Const WM_RBUTTONDOWN            As Long = &H204
Private Const WM_RBUTTONUP              As Long = &H205
Private Const WM_RBUTTONDBLCLK          As Long = &H206
Private Const WM_NCMOUSEMOVE            As Long = &HA0
Private Const WM_NCLBUTTONDOWN          As Long = &HA1
Private Const WM_NCLBUTTONUP            As Long = &HA2
Private Const WM_NCLBUTTONDBLCLK        As Long = &HA3
Private Const WM_NCRBUTTONDOWN          As Long = &HA4
Private Const WM_NCRBUTTONUP            As Long = &HA5
Private Const WM_NCRBUTTONDBLCLK        As Long = &HA6
Private Const WM_NCMBUTTONDOWN          As Long = &HA7
Private Const WM_NCMBUTTONUP            As Long = &HA8
Private Const WM_NCMBUTTONDBLCLK        As Long = &HA9
Private Const WM_KEYDOWN                As Long = &H100
Private Const WM_KEYUP                  As Long = &H101
Private Const WM_CHAR                   As Long = &H102
Private Const WM_DEADCHAR               As Long = &H103
Private Const WM_SYSKEYDOWN             As Long = &H104
Private Const WM_SYSKEYUP               As Long = &H105
Private Const WM_SYSCHAR                As Long = &H106
Private Const WM_SYSDEADCHAR            As Long = &H107
Private Const WM_QUIT                   As Long = &H12

Private Const WM_ERASEBKGND             As Long = &H14
Private Const WM_CONTEXTMENU            As Long = &H7B
Private Const WM_DESTROY                As Long = &H2

Private Const MA_NOACTIVATE             As Long = 3

Private Const MONITOR_DEFAULTTONULL     As Long = &H0
Private Const MONITOR_DEFAULTTONEAREST  As Long = &H2

Private Const CW_USEDEFAULT             As Long = &H80000000

Private Const WS_POPUP                  As Long = &H80000000
Private Const WS_BORDER                 As Long = &H800000
Private Const WS_EX_TOOLWINDOW          As Long = &H80&
Private Const WS_EX_DLGMODALFRAME       As Long = &H1&
Private Const WS_EX_WINDOWEDGE          As Long = &H100&
Private Const WS_EX_TOPMOST             As Long = &H8&

Private Const WS_OVERLAPPED             As Long = &H0&
Private Const WS_CAPTION                As Long = &HC00000
Private Const WS_SYSMENU                As Long = &H80000
Private Const WS_THICKFRAME             As Long = &H40000
Private Const WS_MINIMIZEBOX            As Long = &H20000
Private Const WS_MAXIMIZEBOX            As Long = &H10000
Private Const WS_OVERLAPPEDWINDOW       As Long = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)


'————————————————————————————————————————————————————————————————————————————————

Private g_bQuit As Boolean
Private g_hInstance As Long
Private g_hbrColor  As Long     '// The selected color

'// This is the array of predefined colors we put into the color picker.
Private c_rgclrPredef(0 To 15) As Long


'//
'//  COLORPICKSTATE
'//
'//      Structure that records the state of a color-picker pop-up.
'//
'//      A pointer to this state information is kept in the GWLP_USERDATA
'//      window long.
'//
'//      The iSel field is the index of the selected color, or the
'//      special value -1 to mean that no item is highlighted.
'//
Private Type COLORPICKSTATE
    fDone       As Long         '// Set when we should get out (C-BOOL)
    iSel        As Long         '// Which color is selected?
    iResult     As Long         '// Which color should be returned?
    hwndOwner   As Long         '// Our owner window
End Type

Const CYCOLOR As Long = 16      '// Height of a single color pick
Const CXFAKEMENU As Long = 100  '// Width of our fake menu


'————————————————————————————————————————————————————————————————————————————————


'Private Function CastCOLORPICKSTATE(ByVal Ptr As Long) As COLORPICKSTATE
'  Dim Dummy As COLORPICKSTATE
'    Call CopyMemory(VarPtr(CastCOLORPICKSTATE), Ptr, LenB(Dummy))
'End Function
'
'Private Function CastCREATESTRUCT(ByVal Ptr As Long) As CREATESTRUCT
'  Dim Dummy As CREATESTRUCT
'    Call CopyMemory(VarPtr(CastCREATESTRUCT), Ptr, LenB(Dummy))
'End Function


'————————————————————————————————————————————————————————————————————————————————


Private Function LoWord(ByVal lngValue As Long) As Long
    LoWord = (lngValue And &H7FFF)
End Function

Private Function HiWord(ByVal lngValue As Long) As Long
    If (lngValue And &H80000000) = &H80000000 Then
        HiWord = ((lngValue And &H7FFF0000) \ &H10000) Or &H8000
    Else
        HiWord = (lngValue And &HFFFF0000) \ &H10000
    End If
End Function


'*----------------------------------------------------------*
'* Name       : MAKELONG                                    *
'*----------------------------------------------------------*
'* Purpose    : Combines two integers into a long integer.  *
'*----------------------------------------------------------*
'* Parameters : wLow   Required. Low WORD.                  *
'*            : wHigh  Required. High WORD.                 *
'*----------------------------------------------------------*
'* Description: This function is equivalent to the 'C'      *
'*            : language MAKELONG macro.                    *
'*----------------------------------------------------------*
Public Function MAKELONG(ByVal wLow As Long, ByVal wHigh As Long) As Long
    MAKELONG = LoWord(wLow) Or (&H10000 * LoWord(wHigh))
End Function


'*----------------------------------------------------------*
'* Name       : MAKELPARAM                                  *
'*----------------------------------------------------------*
'* Purpose    : Combines two integers into a long integer.  *
'*----------------------------------------------------------*
'* Parameters : wLow   Required. Low WORD.                  *
'*            : wHigh  Required. High WORD.                 *
'*----------------------------------------------------------*
'* Description: This function is equivalent to the 'C'      *
'*            : language MAKELPARAM macro.                  *
'*----------------------------------------------------------*
Public Function MAKELPARAM(ByVal wLow As Long, ByVal wHigh As Long) As Long
    MAKELPARAM = MAKELONG(wLow, wHigh)
End Function


'————————————————————————————————————————————————————————————————————————————————


'//
'//  ColorPick_GetColorRect
'//
'//      Returns the rectangle that encloses the specified color.
'//
Private Sub ColorPick_GetColorRect(ByRef prc As RECT, ByVal iColor As Long)
    '// Build the "menu" item rect.
    prc.Left = 0
    prc.Right = CXFAKEMENU
    prc.Top = iColor * CYCOLOR
    prc.Bottom = prc.Top + CYCOLOR
End Sub


'//
'//  ColorPick_OnCreate
'//
'//      Stash away our state.
'//
Private Function ColorPick_OnCreate(ByVal hWnd As Long, ByRef pcs As CREATESTRUCT)
    Call SetWindowLong(hWnd, GWLP_USERDATA, pcs.lpCreateParams)
    ColorPick_OnCreate = 0
End Function


'//
'//  ColorPick_OnPaint
'//
'//      Draw the color bars, and put a border around the selected color.
'//
Private Sub ColorPick_OnPaint(ByRef pcps As COLORPICKSTATE, ByVal hWnd As Long)
  Dim ps    As PAINTSTRUCT
  Dim hDC   As Long

    hDC = BeginPaint(hWnd, ps)

    If (hDC) Then
        Dim rcClient As RECT
        Call GetClientRect(hWnd, rcClient)

        '// For each of our predefined colors, draw it in a little
        '// rectangular region, leaving some border so the user can
        '// see if the item is highlighted or not.
        Dim iColor As Long

        For iColor = LBound(c_rgclrPredef) To UBound(c_rgclrPredef)
            '// Build the "menu" item rect.
            Dim rc As RECT
            Call ColorPick_GetColorRect(rc, iColor)

            '// If the item is highlighted, then draw a highlighted background.
            If iColor = pcps.iSel Then
                Call FillRect(hDC, rc, GetSysColorBrush(COLOR_HIGHLIGHT))
            End If

            '// Now shrink the rectangle by an edge and fill the rest with the
            '// color of the item itself.
            Call InflateRect(rc, -GetSystemMetrics(SM_CXEDGE), -GetSystemMetrics(SM_CYEDGE))

            Dim hBr As Long
            hBr = CreateSolidBrush(c_rgclrPredef(iColor))
            Call FillRect(hDC, rc, hBr)
            Call DeleteObject(hBr)
        Next iColor

        Call EndPaint(hWnd, ps)
    End If

End Sub


'//
'//  ColorPick_ChangeSel
'//
'//      Change the selection to the specified item.
'//
Private Sub ColorPick_ChangeSel(ByRef pcps As COLORPICKSTATE, ByVal hWnd As Long, ByVal iSel As Long)
  Dim rc As RECT
    
    '// If the selection changed, then repaint the items that need repainting.
    If pcps.iSel <> iSel Then
        If pcps.iSel >= 0 Then
            Call ColorPick_GetColorRect(rc, pcps.iSel)
            Call InvalidateRect(hWnd, VarPtr(rc), 1)
        End If

        pcps.iSel = iSel
        If pcps.iSel >= 0 Then
            Call ColorPick_GetColorRect(rc, pcps.iSel)
            Call InvalidateRect(hWnd, VarPtr(rc), 1)
        End If
    End If

End Sub


'//
'//  ColorPick_OnMouseMove
'//
'//      Track the mouse to see if it is over any of our colors.
'//
Private Sub ColorPick_OnMouseMove(ByRef pcps As COLORPICKSTATE, ByVal hWnd As Long, ByVal x As Long, ByVal y As Long)
  Dim iSel As Long

    If x >= 0 And x < CXFAKEMENU And y >= 0 And y <= (UBound(c_rgclrPredef) + 1) * CYCOLOR Then
        iSel = Int(y / CYCOLOR)
    Else
        iSel = -1
    End If
    
    Call ColorPick_ChangeSel(pcps, hWnd, iSel)

End Sub


'//
'//  ColorPick_OnLButtonUp
'//
'//      When the button comes up, we are done.
'//
Private Sub ColorPick_OnLButtonUp(ByRef pcps As COLORPICKSTATE, ByVal hWnd As Long, ByVal x As Long, ByVal y As Long)

    '// First track to the final location, in case the user moves the mouse
    '// REALLY FAST and immediately lets go.
    Call ColorPick_OnMouseMove(pcps, hWnd, x, y)

    '// Set the result to the current selection.
    pcps.iResult = pcps.iSel

    '// And tell the message loop that we're done.
    pcps.fDone = True

End Sub


'//
'//  ColorPick_OnRButtonDown
'//
'//      When the button comes up, we are done.
'//
Private Sub ColorPick_OnRButtonDown(ByRef pcps As COLORPICKSTATE, ByVal hWnd As Long, ByVal x As Long, ByVal y As Long)

    '// First track to the final location, in case the user moves the mouse
    '// REALLY FAST and immediately lets go.
    Call ColorPick_OnMouseMove(pcps, hWnd, x, y)

    '// Tell the message loop that we're done,
    '// if the user clicked right outside the context menu.
    If pcps.iSel = -1 Then
        pcps.fDone = True
    End If

End Sub


'//
'//  ColorPick_OnKeyDown
'//
'//      If the ESC key is pressed, then abandon the fake menu.
'//      If the Enter key is pressed, then accept the current selection.
'//      If an arrow key is pressed, the move the selection.
'//
Private Sub ColorPick_OnKeyDown(ByRef pcps As COLORPICKSTATE, ByVal hWnd As Long, ByVal vk As Long)

    Select Case vk
        
        Case VK_ESCAPE:
            pcps.fDone = True           '// Abandoned

        Case VK_RETURN:
            pcps.iResult = pcps.iSel    '// Accept current selection
            pcps.fDone = True

        Case VK_UP:
            If pcps.iSel > 0 Then       '// Decrement selection
                Call ColorPick_ChangeSel(pcps, hWnd, pcps.iSel - 1)
            Else
                Call ColorPick_ChangeSel(pcps, hWnd, UBound(c_rgclrPredef))
            End If

        Case VK_DOWN:                   '// Increment selection
            If pcps.iSel < UBound(c_rgclrPredef) Then
                Call ColorPick_ChangeSel(pcps, hWnd, pcps.iSel + 1)
            Else
                Call ColorPick_ChangeSel(pcps, hWnd, 0)
            End If
            
    End Select

End Sub


'//
'//  ColorPick_WndProc
'//
'//      Window procedure for the color picker popup.
'//
Public Function ColorPick_WndProc(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
  Dim pcps  As COLORPICKSTATE
  Dim pcs   As CREATESTRUCT
  Dim ret   As Long
    
    ret = GetWindowLong(hWnd, GWLP_USERDATA)
    If ret Then
        Call CopyMemory(VarPtr(pcps), ret, LenB(pcps))
    End If

    Select Case uiMsg

        Case WM_CREATE:
            Call CopyMemory(VarPtr(pcs), lParam, LenB(pcs))
            ColorPick_WndProc = ColorPick_OnCreate(hWnd, pcs)
            Exit Function

        Case WM_MOUSEMOVE:
            Call ColorPick_OnMouseMove(pcps, hWnd, LoWord(lParam), HiWord(lParam))

        Case WM_RBUTTONDOWN:
            Call ColorPick_OnRButtonDown(pcps, hWnd, LoWord(lParam), HiWord(lParam))
            
        Case WM_LBUTTONUP, WM_RBUTTONUP:
            Call ColorPick_OnLButtonUp(pcps, hWnd, LoWord(lParam), HiWord(lParam))
            
        Case WM_SYSKEYDOWN, WM_KEYDOWN:
            Call ColorPick_OnKeyDown(pcps, hWnd, wParam)

        '// Do not activate when somebody clicks the window.
        Case WM_MOUSEACTIVATE:
            ColorPick_WndProc = MA_NOACTIVATE
            Exit Function

        Case WM_PAINT:
            Call ColorPick_OnPaint(pcps, hWnd)
            ColorPick_WndProc = 0
            If ret Then
                Call CopyMemory(ret, VarPtr(pcps), LenB(pcps))
            End If
            Exit Function

    End Select

    If ret Then
        Call CopyMemory(ret, VarPtr(pcps), LenB(pcps))
    End If
    ColorPick_WndProc = DefWindowProc(hWnd, uiMsg, wParam, lParam)

End Function


'//
'//  ColorPick_ChooseLocation
'//
'//      Find a place to put the window so it won't go off the screen
'//      or straddle two monitors.
'//
'//      x, y = location of mouse click (preferred upper-left corner)
'//      cx, cy = size of window being created
'//
'//      We use the same logic that real menus use.
'//
'//      -   If (x, y) is too high or too far left, then slide onto screen.
'//      -   Use (x, y) if all fits on the monitor.
'//      -   If too low, then slide up.
'//      -   If too far right, then flip left.
'//
Private Sub ColorPick_ChooseLocation(ByVal hWnd As Long, _
                                     ByVal x As Long, _
                                     ByVal y As Long, _
                                     ByVal cx As Long, _
                                     ByVal cy As Long, _
                                     ByRef ppt As POINTAPI)

    '// First get the dimensions of the monitor that contains (x, y).
    ppt.x = x
    ppt.y = y
     
    Dim hMon As Long
    hMon = MonitorFromPoint(ppt.x, ppt.y, MONITOR_DEFAULTTONULL)

    '// If (x, y) is not on any monitor, then use the monitor that the owner
    '// window is on.
    If hMon = 0 Then
        hMon = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST)
    End If

    Dim minf As MONITORINFO
    minf.cbSize = Len(minf)
    Call GetMonitorInfo(hMon, minf)

    '// Now slide things around until they fit.

    '// If too high, then slide down.
    If ppt.y < minf.rcMonitor.Top Then
        ppt.y = minf.rcMonitor.Top
    End If

    '// If too far left, then slide right.
    If ppt.x < minf.rcMonitor.Left Then
        ppt.x = minf.rcMonitor.Left
    End If

    '// If too low, then slide up.
    If ppt.y > minf.rcMonitor.Bottom - cy Then
        ppt.y = minf.rcMonitor.Bottom - cy
    End If

    '// If too far right, then flip left.
    If ppt.x > minf.rcMonitor.Right - cx Then
        ppt.x = ppt.x - cx
    End If

End Sub


'//
'//  ColorPick_Popup
'//
'//      Display a fake menu to allow the user to select the color.
'//
'//      Return the color index the user selected, or -1 if no color
'//      was selected.
'//
Private Function ColorPick_Popup(ByVal hwndOwner As Long, ByVal x As Long, ByVal y As Long) As Long

    '// Early check:  We must be on same thread as the owner so we can see
    '// its mouse and keyboard messages when we set capture to it.
    If GetCurrentThreadId() <> GetWindowThreadProcessId(hwndOwner, 0&) Then
        '// Error: Menu must be on same thread as parent window.
        ColorPick_Popup = -1
        Debug.Print "GetCurrentThreadId() <> GetWindowThreadProcessId(hwndOwner, 0&)"
        Exit Function
    End If

    Dim cps As COLORPICKSTATE
    cps.fDone = False           '// Not done yet
    cps.iSel = -1               '// No initial selection
    cps.iResult = -1            '// No result
    cps.hwndOwner = hwndOwner   '// Owner window

    '// Set up the style and extended style we want to use.
    Const dwStyle As Long = WS_POPUP Or WS_BORDER
    ' WS_EX_TOOLWINDOW:     So it doesn't show up in taskbar
    ' WS_EX_DLGMODALFRAME:  Get the edges right
    ' WS_EX_WINDOWEDGE:
    ' WS_EX_TOPMOST:        So it isn't obscured
    Const dwExStyle As Long = WS_EX_TOOLWINDOW Or _
                              WS_EX_DLGMODALFRAME Or _
                              WS_EX_WINDOWEDGE Or _
                              WS_EX_TOPMOST

    '// We want a client area of size (CXFAKEMENU, ARRAYSIZE(c_rgclrPredef) * CYCOLOR),
    '// so use AdjustWindowRectEx to figure out what window rect will give us a
    '// client rect of that size.
    Dim rc As RECT
    rc.Left = 0
    rc.Top = 0
    rc.Right = CXFAKEMENU
    rc.Bottom = (UBound(c_rgclrPredef) + 1) * CYCOLOR
    Call AdjustWindowRectEx(rc, dwStyle, 0, dwExStyle)

    '// Now find a proper home for the window that won't go off the screen or
    '// straddle two monitors.
    Dim cx As Long: cx = rc.Right - rc.Left
    Dim cy As Long: cy = rc.Bottom - rc.Top
    Dim pt As POINTAPI
    Call ColorPick_ChooseLocation(hwndOwner, x, y, cx, cy, pt)

    Dim hwndPopup As Long
    hwndPopup = CreateWindowEx(dwExStyle, "ColorPick", "", dwStyle, _
                               pt.x, pt.y, cx, cy, _
                               hwndOwner, 0&, g_hInstance, VarPtr(cps))

    '// Show the window but don't activate it!
    Call ShowWindow(hwndPopup, SW_SHOWNOACTIVATE)

    '// We want to receive all mouse messages, but since only the active
    '// window can capture the mouse, we have to set the capture to our
    '// owner window, and then steal the mouse messages out from under it.
    Call SetCapture(hwndOwner)

    '// Go into a message loop that filters all the messages it receives
    '// and route the interesting ones to the color picker window.
    Dim tMsg As MSG
    Do While GetMessage(tMsg, 0&, 0, 0)

        '// Something may have happened that caused us to stop.
        If cps.fDone Then
            Exit Do
        End If

        '// If our owner stopped being the active window (e.g. the user
        '// Alt+Tab'd to another window in the meantime), then stop.
        Dim hwndActive As Long
        hwndActive = GetActiveWindow()
        If hwndActive <> hwndOwner And Not IsChild(hwndActive, hwndOwner) Or GetCapture() <> hwndOwner Then
            Exit Do
        End If

        '// At this point, we get to snoop at all input messages before
        '// they get dispatched.  This allows us to route all input to our
        '// popup window even if really belongs to somebody else.

        '// All mouse messages are remunged and directed at our popup
        '// menu. If the mouse message arrives as client coordinates, then
        '// we have to convert it from the client coordinates of the original
        '// target to the client coordinates of the new target.
        Select Case tMsg.message

            '// These mouse messages arrive in client coordinates, so in
            '// addition to stealing the message, we also need to convert
            '// the coordinates.
            Case WM_MOUSEMOVE, _
                 WM_LBUTTONDOWN, WM_LBUTTONUP, WM_LBUTTONDBLCLK, _
                 WM_RBUTTONDOWN, WM_RBUTTONUP, WM_RBUTTONDBLCLK, _
                 WM_MBUTTONDOWN, WM_MBUTTONUP, WM_MBUTTONDBLCLK
                 
                pt.x = LoWord(tMsg.lParam)
                pt.y = HiWord(tMsg.lParam)
                Call MapWindowPoints(tMsg.hWnd, hwndPopup, VarPtr(pt), 1)
                tMsg.lParam = MAKELPARAM(pt.x, pt.y)
                tMsg.hWnd = hwndPopup

            '// These mouse messages arrive in screen coordinates, so we
            '// just need to steal the message.
            Case WM_NCMOUSEMOVE, _
                 WM_NCLBUTTONDOWN, WM_NCLBUTTONUP, WM_NCLBUTTONDBLCLK, _
                 WM_NCRBUTTONDOWN, WM_NCRBUTTONUP, WM_NCRBUTTONDBLCLK, _
                 WM_NCMBUTTONDOWN, WM_NCMBUTTONUP, WM_NCMBUTTONDBLCLK
                tMsg.hWnd = hwndPopup

            '// We need to steal all keyboard messages, too.
            Case WM_KEYDOWN, WM_KEYUP, WM_CHAR, WM_DEADCHAR, _
                 WM_SYSKEYDOWN, WM_SYSKEYUP, WM_SYSCHAR, WM_SYSDEADCHAR
                tMsg.hWnd = hwndPopup

        End Select

        Call TranslateMessage(tMsg)
        Call DispatchMessage(tMsg)

        '// Something may have happened that caused us to stop.
        If cps.fDone Then
            Exit Do
        End If

        '// If our owner stopped being the active window (e.g. the user
        '// Alt+Tab'd to another window in the meantime), then stop.
        hwndActive = GetActiveWindow()
        If hwndActive <> hwndOwner And Not IsChild(hwndActive, hwndOwner) Or GetCapture() <> hwndOwner Then
            Exit Do
        End If

    Loop

    '// Clean up the capture we created.
    Call ReleaseCapture

    Call DestroyWindow(hwndPopup)

    '// If we got a WM_QUIT message, then re-post it so the caller's message
    '// loop will see it.
    If tMsg.message = WM_QUIT Then
        g_bQuit = True                                      'Call PostQuitMessage(tMsg.wParam)
    End If

    ColorPick_Popup = cps.iResult

End Function


'//
'//  FakeMenuDemo_OnEraseBkgnd
'//
'//      Erase the background in the selected color.
'//
Private Sub FakeMenuDemo_OnEraseBkgnd(ByVal hWnd As Long, ByVal hDC As Long)
  Dim rc As RECT
    Call GetClientRect(hWnd, rc)
    Call FillRect(hDC, rc, g_hbrColor)
End Sub


'//
'//  FakeMenuDemo_OnContextMenu
'//
'//      Display the color-picker pseudo-menu so the user can change
'//      the color.
'//
Private Sub FakeMenuDemo_OnContextMenu(ByVal hWnd As Long, ByVal x As Long, ByVal y As Long)
    
    '// If the coordinates are (-1, -1), then the user used the keyboard -
    '// we'll pretend the user clicked at client (0, 0).
    If x = -1 And y = -1 Then
        Dim pt As POINTAPI
        pt.x = 0
        pt.y = 0
        Call ClientToScreen(hWnd, pt)
        x = pt.x
        y = pt.y
    End If

    Dim iColor As Long
    iColor = ColorPick_Popup(hWnd, x, y)

    '// If the user picked a color, then change to that color.
    If iColor >= 0 Then
        Call DeleteObject(g_hbrColor)
        g_hbrColor = CreateSolidBrush(c_rgclrPredef(iColor))
        Call InvalidateRect(hWnd, 0&, 1)
    End If

End Sub


'//
'//  FakeMenuDemo_WndProc
'//
'//      Window procedure for the fake menu demo.
'//
Public Function FakeMenuDemo_WndProc(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Select Case uiMsg

        Case WM_ERASEBKGND:
            Call FakeMenuDemo_OnEraseBkgnd(hWnd, wParam)
            FakeMenuDemo_WndProc = 1                        'return TRUE;
            Exit Function

        Case WM_CONTEXTMENU:
            Call FakeMenuDemo_OnContextMenu(hWnd, LoWord(lParam), HiWord(lParam))
            FakeMenuDemo_WndProc = 0
            Exit Function

        Case WM_DESTROY:
            g_bQuit = True                                  'Call PostQuitMessage(0)

    End Select

    FakeMenuDemo_WndProc = DefWindowProc(hWnd, uiMsg, wParam, lParam)

End Function


'
Private Function GetAddress(ByVal Address As Long) As Long
    GetAddress = Address
End Function


'//
'//      Program entry point - demonstrate pseudo-menus.
'//
Public Function wWinMain(ByVal hInstance As Long, ByVal hInstPrev As Long, ByVal pszCmdLine As String, ByVal nCmdShow As Long) As Long
  Const IDC_ARROW       As Long = 32512&
  Const COLOR_BTNFACE   As Long = 15
  Const COLOR_3DFACE    As Long = COLOR_BTNFACE

    c_rgclrPredef(0) = RGB(&H0, &H0, &H0)                 '// 0 = black
    c_rgclrPredef(1) = RGB(&H80, &H0, &H0)                '// 1 = maroon
    c_rgclrPredef(2) = RGB(&H0, &H80, &H0)                '// 2 = green
    c_rgclrPredef(3) = RGB(&H80, &H80, &H0)               '// 3 = olive
    c_rgclrPredef(4) = RGB(&H0, &H0, &H80)                '// 4 = navy
    c_rgclrPredef(5) = RGB(&H80, &H0, &H80)               '// 5 = purple
    c_rgclrPredef(6) = RGB(&H0, &H80, &H80)               '// 6 = teal
    c_rgclrPredef(7) = RGB(&H80, &H80, &H80)              '// 7 = gray
    c_rgclrPredef(8) = RGB(&HC0, &HC0, &HC0)              '// 8 = silver
    c_rgclrPredef(9) = RGB(&HFF, &H0, &H0)                '// 9 = red
    c_rgclrPredef(10) = RGB(&H0, &HFF, &H0)               '// A = lime
    c_rgclrPredef(11) = RGB(&HFF, &HFF, &H0)              '// B = yellow
    c_rgclrPredef(12) = RGB(&H0, &H0, &HFF)               '// C = blue
    c_rgclrPredef(13) = RGB(&HFF, &H0, &HFF)              '// D = fuchsia
    c_rgclrPredef(14) = RGB(&H0, &HFF, &HFF)              '// E = cyan
    c_rgclrPredef(15) = RGB(&HFF, &HFF, &HFF)             '// F = white

    g_hInstance = hInstance

    Dim wc As WNDCLASSEX
    wc.cbSize = Len(wc)
    
    wc.lpfnWndProc = GetAddress(AddressOf ColorPick_WndProc)
    wc.hInstance = g_hInstance
    wc.hCursor = LoadCursor(0&, IDC_ARROW)
    wc.hbrBackground = (COLOR_3DFACE + 1)
    wc.lpszClassName = "ColorPick"
    Call RegisterClassEx(wc)

    wc.style = 0
    wc.lpfnWndProc = GetAddress(AddressOf FakeMenuDemo_WndProc)
    wc.hInstance = g_hInstance
    wc.hbrBackground = 0&               '// Background color is dynamic
    wc.lpszClassName = "FakeMenuDemo"
    Call RegisterClassEx(wc)

    g_hbrColor = CreateSolidBrush(RGB(&HFF, &HFF, &HFF))

    Dim hWnd As Long
    'hWnd = CreateWindowEx(0, "FakeMenuDemo", "Fake Menu Demo - Right-click in window to change color", WS_OVERLAPPEDWINDOW, _
                          CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, _
                          0&, 0&, g_hInstance, 0)
    hWnd = CreateWindowEx(0, "FakeMenuDemo", "Fake Menu Demo - Right-click in window to change color", WS_OVERLAPPEDWINDOW, _
                          CW_USEDEFAULT, CW_USEDEFAULT, 300, 200, _
                          0&, 0&, g_hInstance, 0)
                          
    If hWnd Then
      Call ShowWindow(hWnd, nCmdShow)
  
      Dim tMsg As MSG
      Do While (GetMessage(tMsg, 0&, 0, 0) <> 0) And Not g_bQuit
          Call TranslateMessage(tMsg)
          Call DispatchMessage(tMsg)
      Loop
    End If
    
    If g_hbrColor Then Call DeleteObject(g_hbrColor)

    wWinMain = 0

End Function


Public Sub Main()
    Call wWinMain(App.hInstance, 0&, "", SW_SHOWNORMAL)
End Sub

