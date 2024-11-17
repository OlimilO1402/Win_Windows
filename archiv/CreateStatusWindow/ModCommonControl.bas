Attribute VB_Name = "ModCommonControl"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2006 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Set of bit flags that indicate which common control classes will be loaded
'from the DLL. The dwICC value of tagINITCOMMONCONTROLSEX can
'be a combination of the following:
Public Const ICC_LISTVIEW_CLASSES As Long = &H1 'listview, header
Public Const ICC_TREEVIEW_CLASSES As Long = &H2 'treeview, tooltips
Public Const ICC_BAR_CLASSES As Long = &H4      'toolbar, statusbar, trackbar, tooltips
Public Const ICC_TAB_CLASSES As Long = &H8      'tab, tooltips
Public Const ICC_UPDOWN_CLASS As Long = &H10    'updown
Public Const ICC_PROGRESS_CLASS As Long = &H20  'progress
Public Const ICC_HOTKEY_CLASS As Long = &H40    'hotkey
Public Const ICC_ANIMATE_CLASS As Long = &H80   'animate
Public Const ICC_WIN95_CLASSES As Long = &HFF   'everything else
Public Const ICC_DATE_CLASSES As Long = &H100   'month picker, date picker, time picker, updown
Public Const ICC_USEREX_CLASSES As Long = &H200 'comboex
Public Const ICC_COOL_CLASSES As Long = &H400   'rebar (coolbar) control

'WIN32_IE >= 0x0400
Public Const ICC_INTERNET_CLASSES As Long = &H800
Public Const ICC_PAGESCROLLER_CLASS As Long = 1000 'page scroller
Public Const ICC_NATIVEFNTCTL_CLASS As Long = 2000 'native font control

'WIN32_WINNT >= 0x501
Public Const ICC_STANDARD_CLASSES As Long = 4000
Public Const ICC_LINK_CLASS As Long = 8000

  
'Initializes the entire common control dynamic-link library.
'Exported by all versions of Comctl32.dll.
Public Declare Sub InitCommonControls Lib "comctl32" ()
  
'Initializes specific common controls classes from the common
'control dynamic-link library.
'Returns TRUE (non-zero) if successful, or FALSE otherwise.
'Began being exported with Comctl32.dll version 4.7 (IE3.0 & later).
Public Declare Function InitCommonControlsEx Lib "comctl32" _
      (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean

Public Type tagINITCOMMONCONTROLSEX   ' icc
   dwSize As Long   ' size of this structure
   dwICC As Long    ' flags indicating which classes to be initialized.
End Type

  
Public Function InitComctl32(dwFlags As Long) As Boolean

  'Common control Initializing routine.
  'Returns True if the current working version of
  'Comctl32.dll is available on the system and the
  'new IE3 styles and messages can be used. Returns
  'False either if the old version is present or
  'Comctl32.dll isn't available at all. Also ensures
  'that the library is loaded for use.

  'We can get away with this hack rather than checking
  'the file's version because VB resolves declared API
  'function names only when they're called, not when
  'it compiles the code...
   Dim icc As tagINITCOMMONCONTROLSEX
   
   On Error GoTo Err_OldVersion
  
   With icc
      .dwSize = Len(icc)
      .dwICC = dwFlags
   End With
     
  'VB will generate error 453 "Specified DLL
  'function not found" if InitCommonControlsEx
  'can't be located in the library. The error
  'is trapped and the original InitCommonControls
  'is called instead below.
   InitComctl32 = InitCommonControlsEx(icc)
   Exit Function

Err_OldVersion:
   InitCommonControls

End Function

