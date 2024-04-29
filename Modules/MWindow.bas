Attribute VB_Name = "MWindow"
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function DefWindowProcW Lib "user32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
    'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-defwindowprocw
    Private Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#End If

'LRESULT Wndproc( HWND unnamedParam1, UINT unnamedParam2, WPARAM unnamedParam3, LPARAM unnamedParam4 )
Public Function WndProc(ByVal hWnd_Param1 As LongPtr, ByVal uiMsg_Param2 As Long, ByVal wParam3 As LongPtr, ByVal lParam4 As LongPtr) As LongPtr
    'Debug.Print "WndProc"
    
    WndProc = DefWindowProcW(hWnd_Param1, uiMsg_Param2, wParam3, lParam4)
    
End Function
