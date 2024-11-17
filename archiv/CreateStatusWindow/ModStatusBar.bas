Attribute VB_Name = "ModStatusBar"
Option Explicit

'CreateWindowEx(
'    0,
'    'msctls_statusbar32',
'    nil,
'    WS_CHILD or WS_VISIBLE,
'    0,
'    0,
'    0,
'    0,
'    Handle,
'    101,
'    hInstance,
'    nil
'    );
'
'    SysTabControl32,
'    SysTreeView32,
'    msctls_hotkey32,
'    msctls_progress32,
'    msctls_statusbar32,
'    msctls_trackbar32,
'    msctls_updown32,
'    ComboBoxEx32,
'    ReBarWindow32

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
'Creates a status window along the bottom of the
'parent window. Its width is the same as that of
'the parent window's client area. The height is based
'on the metrics of the font that is currently selected
'into the status window's device context and on the
'width of the window's borders.
Public Declare Function CreateStatusWindow Lib "comctl32.dll" Alias "CreateStatusWindowA" (ByVal style As Long, ByVal lpszText As String, ByVal hWndParent As Long, ByVal wID As Long) As Long
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long

Private Const STATUSCLASSNAME As String = "msctls_statusbar32"
Private Const STATUSCLASSNAMEA As String = "msctls_statusbar32"
Private Const STATUSCLASSNAMEW As String = "msctls_statusbar32"
 
    
'Status Bar Styles

'Include a sizing grip at the right end of the status window.
Public Const SBARS_SIZEGRIP As Long = &H100

'Creates a window that is initially visible.
Public Const WS_VISIBLE As Long = &H10000000

'Creates a child window.
Public Const WS_CHILD As Long = &H40000000
       
Public Declare Sub DrawStatusText Lib "comctl32" Alias "DrawStatusTextA" (ByVal hDC As Long, lprc As RECT, ByVal pszText As Long, uFlags As Long)

'Also used w/ SB_GETRECT below
Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
 End Type

'uFlags text drawing operation values:
'also used by SB_SETTEXT, SB_SETTEXT, SB_GETTEXTLENGTH below
'The text is drawn with a border to appear lower
'than the plane of the window.
Public Const SBT_SUNKEN As Long = &H0  'Default

'The text is drawn without borders.
Public Const SBT_NOBORDERS As Long = &H100

'The text is drawn with a border to appear higher
'than the plane of the window.
Public Const SBT_POPOUT As Long = &H200

'SB_SETTEXT, SB_SETTEXT, SB_GETTEXTLENGTH flags only:
'Displays text using right-to-left reading order on
'Hebrew or Arabic systems.
Public Const SBT_RTLREADING As Long = &H400

'The text is drawn by the parent window.
Public Const SBT_OWNERDRAW As Long = &H1000
 

'Status Bar Control Messages
 
 Public Const WM_USER As Long = &H400

'The SB_SETTEXT message sets the text in the specified
'part of a status window.
'wParam = iPart Or uType
'iPart = Zero-based index of the part to set. If this
'        value is 255, the status window is assumed to be a
'        simple window having only one part.
'uType = Type of drawing operation. This parameter can
'         be one of the following values:
'        0, SBT_NOBORDERS, SBT_OWNERDRAW, SBT_POPOUT, SBT_RTLREADING
'lParam = Pointer to a null-terminated string that specifies
'        the text to set. If uType is SBT_OWNERDRAW, this parameter
'        represents 32 bits of data. The parent window must interpret
'        the data and draw the text when it receives the
'        WM_DRAWITEM message.
'Rtns True on success, False if failure
'The message invalidates the portion of the window that
'has changed, causing it to display the new text when the
'window next receives the WM_PAINT message.
Public Const SB_SETTEXT As Long = (WM_USER + 1)
Public Const sbSimpleIdx As Long = 255
 

'The SB_GETTEXT message retrieves the text from the specified
'part of a status window.
'wParam = Zero-based index of the part from which to retrieve text.
'lParam = Pointer to the null-terminated string buffer that receives the text.
'Returns a 32-bit value that consists of two 16-bit values. The
'low word specifies the length, in characters, of the text. The high
'word specifies the type of operation used to draw the text. The
'type can be one of the following values:
'    0, SBT_NOBORDERS, SBT_POPOUT, SBT_RTLREADING
'If the text has the SBT_OWNERDRAW drawing type, this message
'returns the 32-bit value associated with the text instead of
'the length and operation type.
Public Const SB_GETTEXT As Long = (WM_USER + 2)
 

'The SB_GETTEXTLENGTH message retrieves the length, in characters,
'of the text from the specified part of a status window.
'wParam = Zero-based index of the part from which to retrieve text.
'lParam = 0;
'Returns a 32-bit value that consists of two 16-bit values. The
'low word specifies the length, in characters, of the text. The
'high word specifies the type of operation used to draw the text.
'The type can be one of the following values:
'0, SBT_NOBORDERS, SBT_OWNERDRAW, SBT_POPOUT, SBT_RTLREADING
 Public Const SB_GETTEXTLENGTH As Long = (WM_USER + 3)
 

'The SB_SETPARTS message sets the number of parts in a status
'window and the coordinate of the right edge of each part.
'wParam = Number of parts to set. The number of parts cannot be
'         greater than 255.
'lParam = Pointer to an integer array that has the same number
'         of elements as parts specified by nParts. Each element
'         in the array specifies the position, in client
'         coordinates, of the right edge of the corresponding part.
'         If an element is  - 1, the position of the right edge for
'         that part extends to the right edge of the window.
'Rtns True on success, False if failure
Public Const SB_SETPARTS As Long = (WM_USER + 4)
 

'The SB_GETPARTS message retrieves a count of the parts in
'a status window. The message also retrieves the coordinate of
'the right edge of the specified number of parts.
'wParam = Number of parts for which to retrieve coordinates.
'         If this parameter is greater than the number of
'         parts in the window, the message retrieves coordinates
'         for existing parts only.
'lParam = Pointer to an integer array that has the same number
'         of elements as parts specified by nParts. Each element
'         in the array receives the client coordinate of the right
'         edge of the corresponding part. If an element is set to - 1,
'         the position of the right edge for that part extends to the
'         right edge of the window. To retrieve the current number
'         of parts, set this parameter to zero.
'Returns the number of parts in the window if successful, or zero otherwise.
Public Const SB_GETPARTS As Long = (WM_USER + 6)
 

'The SB_GETBORDERS message retrieves the current widths of the horizontal
'and vertical borders of a status window. This message was once used for
'header windows as well
'The SB_GETBORDERS message retrieves the current width of the horizontal
'and vertical borders of a status or header window. These measurements
'determine the spacing between the outer edge of the window and the
'rectangles within the window that contain text, and the spacing between
'rectangles
'wParam = 0;
'lParam = Pointer to an integer array that has three elements. The
'         first element receives the width of the horizontal border,
'         the second receives the width of the vertical border, and
'         the third receives the width of the border between rectangles.
'Rtns True on success, False if failure
'The borders determine the spacing between the outside edge of the window
'and the rectangles within the window that contain text. The borders also
'determine the spacing between rectangles.
Public Const SB_GETBORDERS As Long = (WM_USER + 7)


Public Const SBB_HORIZONTAL As Long = 0 'horz border width
Public Const SBB_VERTICAL As Long = 1   'vert border width
Public Const SBB_DIVIDER As Long = 2    'vert part divider width
 

'The SB_SETMINHEIGHT message sets the minimum height of a
'status window's drawing area.
'wParam = Minimum height, in pixels, of the window.
'lParam = 0;
'No return value.
'The minimum height is the sum of wParam and twice (?) the
'width, in pixels, of the vertical border of the status window.
'An application must send the WM_SIZE message to the status window
'to redraw the window. The wParam and lParam parameters of the
'WM_SIZE message should be set to zero.

'The SB_SETMINHEIGHT message sets the minimum height for a status
'bar or header window. The minimum height of the window is the sum
'of the minimum height (wParam) and the height of the vertical
'border of the window.
Public Const SB_SETMINHEIGHT As Long = (WM_USER + 8)
 

'The SB_SIMPLE message specifies whether a status window
'displays simple text or displays all window parts set by
'a previous SB_SETPARTS message. The string that a status
'bar displays in simple mode is maintained separately
'from the strings it displays when it is in multiple-part mode
'wParam = Display type flag. If this parameter is TRUE, the window
'displays simple text. If it is FALSE, it displays multiple parts.
'lParam = 0;
'Returns FALSE if an error occurs.
'If the status window is being changed from non-simple to simple,
'or vice versa, the window is immediately redrawn.
Public Const SB_SIMPLE As Long = (WM_USER + 9)
 

'The SB_GETRECT message retrieves the bounding rectangle of
'a part in a status window.
'wParam = Zero-based index of the part whose bounding rectangle
'         is to be retrieved.
'lParam = Pointer to a RECT structure that receives the bounding
'         rectangle.
'Rtns True on success, False if failure
Public Const SB_GETRECT As Long = (WM_USER + 10)
 

'>= IE3 only!!
'The SB_ISSIMPLE message checks a status window control to
'determine if it is in simple mode.
'wParam = 0;
'lParam = 0;
'Returns nonzero if the status window control is in simple mode,
'or zero otherwise.
Public Const SB_ISSIMPLE As Long = (WM_USER + 14)
 
'Public Const SB_SETTEXTA As Long = (WM_USER + 1)
'Public Const SB_GETTEXTA As Long = (WM_USER + 2)
'Public Const SB_GETTEXTLENGTHA As Long = (WM_USER + 3)
'Public Const SB_SETPARTS As Long = (WM_USER + 4)
'Public Const SB_GETPARTS As Long = (WM_USER + 6)
'Public Const SB_GETBORDERS As Long = (WM_USER + 7)
'Public Const SB_SETMINHEIGHT As Long = (WM_USER + 8)
'Public Const SB_SIMPLE As Long = (WM_USER + 9)
'Public Const SB_GETRECT As Long = (WM_USER + 10)
Public Const SB_SETTEXTW As Long = (WM_USER + 11)
Public Const SB_GETTEXTLENGTHW As Long = (WM_USER + 12)
Public Const SB_GETTEXTW As Long = (WM_USER + 13)
'Public Const SB_ISSIMPLE As Long = (WM_USER + 14)
Public Const SB_SETICON As Long = (WM_USER + 15)
Public Const SB_SETTIPTEXTA As Long = (WM_USER + 16)
Public Const SB_SETTIPTEXTW As Long = (WM_USER + 17)
Public Const SB_GETTIPTEXTA As Long = (WM_USER + 18)
Public Const SB_GETTIPTEXTW As Long = (WM_USER + 19)
Public Const SB_GETICON As Long = (WM_USER + 20)
 
 
'Status Bar Standard Window Messages

'Initializes the status window.
Public Const WM_CREATE As Long = &H1

'Frees resources allocated for the status window.
Public Const WM_DESTROY As Long = &H2

'Resizes the status window based on the current width
'of the parent window client area and the height of
'the current font of the status window.
Public Const WM_SIZE As Long = &H5

'Copies the specified text into the first part of a
'status window, using the default drawing
'operation (specified as zero). It returns TRUE if
'successful or FALSE otherwise.
Public Const WM_SETTEXT As Long = &HC

'Copies the text from the first part of a status
'window to a buffer. (If in simple mode, copies the
'simple mode text.) It returns a 32-bit value that
'specifies the length, in characters, of the text
'and the technique used to draw the text.
Public Const WM_GETTEXT As Long = &HD

'Returns a 32-bit value that specifies the length,
'in characters, of the text in the first part of a
'status window and the technique used to draw the text.
'(If in simple mode, returns simple mode text information.)
Public Const WM_GETTEXTLENGTH As Long = &HE

'Paints the invalid region of the status window. If
'the wParam parameter is non-NULL, the control assumes
'that the value is an HDC and paints using that device context.
Public Const WM_PAINT As Long = &HF

'Selects the font handle into the device context
'for the status window.
Public Const WM_SETFONT As Long = &H30

'Returns the handle to the current font with which
'the status window draws its text.
Public Const WM_GETFONT = &H31

'Returns the HTBOTTOMRIGHT value if the mouse cursor
'is in the sizing grip, causing the system to display
'the sizing cursor. If the mouse cursor is not in the
'sizing grip, the status window passes this message
'to the DefWindowProc function.
Public Const WM_NCHITTEST As Long = &H84



'Common Control Styles

'Following are the common control styles. Except
'where noted, these styles apply to header controls,
'toolbar controls, and status windows.
      
'Causes the control to position itself at the top of
'the parent window client area and sets the width to
'be the same as the parent window width. Toolbars have
'this style by default.
Public Const CCS_TOP As Long = &H1&

'Causes the control to resize and move itself horizontally,
'but not vertically, in response to a WM_SIZE message. If
'CCS_NORESIZE is used, this style does not apply. Header
'windows have this style by default.
Public Const CCS_NOMOVEY As Long = &H2&

'Causes the control to position itself at the bottom of
'the parent window client area and sets the width to
'be the same as the parent window width. Status windows
'have this style by default.
Public Const CCS_BOTTOM As Long = &H3&

'Prevents the control from using the default width and
'height when setting its initial size or a new size.
'Instead, the control uses the width and height
'specified in the request for creation or sizing.
Public Const CCS_NORESIZE As Long = &H4&

'Prevents the control from automatically moving to the
'top or bottom of the parent window. Instead, the control
'keeps its position within the parent window despite
'changes to the size of the parent. If CCS_TOP or CCS_BOTTOM
'is also used, the height is adjusted to the default, but
'the position and width remain unchanged.
Public Const CCS_NOPARENTALIGN As Long = &H8&

'Prevents a one-pixel highlight from being drawn at the
'top of the control.
Public Const CCS_NOHILITE As Long = &H10&

'Enables a toolbar's built-in customization features, which
'allow the user to drag a button to a new position or to remove
'a button by dragging it off the toolbar. In addition, the user
'can double-click the toolbar to display the Customize Toolbar
'dialog box, allowing the user to add, delete, and rearrange
'toolbar buttons.
Public Const CCS_ADJUSTABLE As Long = &H20&

'Prevents a two-pixel highlight from being drawn at the top of
'the control.
Public Const CCS_NODIVIDER As Long = &H40&


'At some point early on in the status bar's life, a SB_SETBORDERS message
'may also have been supported but apparently is not anymore.
'From "Win32 Common Controls, Part 2: Status Bars and Toolbars"
'(Nancy Cluts. Microsoft Developer Network Technology Group, March 1994):
'
'"The SB_SETBORDERS message sets the widths of the horizontal and vertical
'borders of a status bar or header window. These borders determine the spacing
'between the outer edge of the window and the rectangles within the window that
'contain text, and the spacing between rectangles.
'wParam = 0, not used

'lParam = The address of an integer array that has three elements.
'The first element specifies the width of the horizontal border,
'the second specifies the width of the vertical border,
'and the third specifies the width of the border between rectangles.
'If an element is set to –1, the default width for the border is used.

'Rtns True on success, False if failure
'Public Const SB_SETBORDERS = (WM_USER + ?)".

