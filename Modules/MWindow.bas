Attribute VB_Name = "MWindow"
Option Explicit

'LRESULT Wndproc( HWND unnamedParam1, UINT unnamedParam2, WPARAM unnamedParam3, LPARAM unnamedParam4 )
Public Function WndProc(ByVal hWnd_Param1 As LongPtr, ByVal Param2 As Long, ByVal WParam3 As LongPtr, ByVal LParam4 As LongPtr) As LongPtr
    Debug.Print "WndProc"
End Function
