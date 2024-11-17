Attribute VB_Name = "ModCommonStatusRoutine"
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
Public Declare Function DestroyWindow Lib "user32" _
  (ByVal hWnd As Long) As Long

Public Declare Function IsWindow Lib "user32" _
  (ByVal hWnd As Long) As Long

Public Declare Function MoveWindow Lib "user32" _
  (ByVal hWnd As Long, _
   ByVal x As Long, ByVal y As Long, _
   ByVal nWidth As Long, ByVal nHeight As Long, _
   ByVal bRepaint As Long) As Long

Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   wParam As Any, _
   lParam As Any) As Long

'The data type for lpRect was changed from "As RECT" to
'"As Any" to allow a null pointer to be passed (i.e. ByRef 0)
'BTW the RECT structure is declared in MStatusDefs
Public Declare Function InvalidateRect Lib "user32" _
  (ByVal hWnd As Long, _
   lpRect As Any, _
   ByVal bErase As Long) As Long

Public Declare Function GetWindowRect Lib "user32" _
  (ByVal hWnd As Long, _
   lpRect As RECT) As Long


Public Function GetParts(hStatBar As Long) As Integer
   'Returns the current number of existing parts in the status bar.
   'The SB_GETPARTS message also retrieve individual part
   'information. See the comments in Status.bas for more info.

  GetParts = SendMessage(hStatBar, SB_GETPARTS, 0, ByVal 0)
  
End Function


Public Sub SetText(hStatBar As Long, _
                   bPart As Byte, _
                   wNewDrawOp As Integer, _
                   sText As String)
  
   'Sets the specified part's text.
  '
  'bPart:zero-based part to set, 255 = simple mode text.
  'wNewDrawOp:text drawing operation.
  'sText:text to set

   Dim wCurDrawOp As Integer

   'Get the part's current drawing operation
   'before it might be updated below.
   wCurDrawOp = GetCurDrawOp(hStatBar, bPart, False)

   'Set the text w/ the drawing operation
   SendMessage hStatBar, SB_SETTEXT, ByVal bPart Or wNewDrawOp, ByVal sText
  
   'Redraw the status bar only if the part's drawing
   'operation changed (reduces flicker).
   If wCurDrawOp <> wNewDrawOp Then InvalidateRect hStatBar, ByVal 0, True
  
End Sub


Public Function GetCurDrawOp(hStatBar As Long, _
                             bPart As Byte, _
                             fRtnString As Boolean) As Integer 'String 'wieso String????
  
   'Returns the current text drawing operation for the specified part.
   '
   'SB_GETTEXTLENGTH is used to determine the part's current
   'drawing operation. SB_GETTEXT will rtn the exact same value,
   'but requires a text buffer.
   '
   'When not in simple mode, SB_GETTEXTLENGTH  retrieves the
   'text length for the part specified by bPart (0-254, 255 parts max).
   'If in simple mode, SB_GETTEXTLENGTH will retrieve the simple
   'mode text length *only* if bPart specifies any *valid* part index.
   'The simple mode text length is NOT retrieved when bPart = 255
   '(as is used to set text w/ SB_SETTEXT). Also applies to
   'SB_GETTEXT.
   '
   'If fRtnString = True, returns the text drawing operation constant
   'string. If False, returns the text drawing operation constant value.

   Dim dwRtn As Long

   dwRtn = SendMessage(hStatBar, SB_GETTEXTLENGTH, ByVal bPart, 0)

   'The text drawing operation for the specified
   'part is contained in the high word of dwRtn.
   dwRtn = (dwRtn And &HFFFF0000) \ &HFFFF&

   If fRtnString Then

     'Returning the string
      Select Case dwRtn
         Case SBT_SUNKEN:    GetCurDrawOp = "SBT_SUNKEN"
         Case SBT_NOBORDERS: GetCurDrawOp = "SBT_NOBORDERS"
         Case SBT_POPOUT:    GetCurDrawOp = "SBT_POPOUT"
      End Select

   Else

     'Returning the value
      GetCurDrawOp = dwRtn
  
   End If
End Function


Public Sub SetParts(frm As Form, hStatBar As Long, bParts As Byte) '1-255 max!

   'Sets the specified number of status bar parts.
  'Any existing part with a greater index than the number of parts
  'specified by bParts is destroyed,
  'i.e 8 existing parts (0-7), 6 is specified for bParts,
  'the last 2 parts (6 & 7) are destroyed.

  'Array is zero based, will error back to
  'cmdDoMsgs_Click() if 0 is passed.

   ReDim adwParts(bParts - 1) As Long
   Dim bPart As Byte

  'Set all but the last part so they have an equal width.
   For bPart = 1 To bParts - 1
     adwParts(bPart - 1) = (frm.ScaleWidth \ bParts) * bPart
   Next

  'Last part uses remaining real estate & extends to right edge.
   adwParts(bParts - 1) = -1
  
   SendMessage hStatBar, SB_SETPARTS, ByVal bParts, adwParts(0)
  
End Sub


