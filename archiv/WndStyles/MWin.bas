Attribute VB_Name = "MWin"
Option Explicit

#If VBA7 = 0 Then
    Public Enum LongPtr
        [_]
    End Enum
#End If

'Animation Control
Public Enum EAnimCStyle
    ACS_CENTER = &H1      ' Centers the animation in the animation control's window.
    ACS_TRANSPARENT = &H2 ' Allows you to match an animation's background color to that of the underlying window, creating a "transparent" background. The parent of the animation control must not have the WS_CLIPCHILDREN style (see Window Styles). The control sends a WM_CTLCOLORSTATIC message to its parent. Use SetBkColor to set the background color for the device context to an appropriate value. The control interprets the upper-left pixel of the first frame as the animation's default background color. It will remap all pixels with that color to the value you supplied in response to WM_CTLCOLORSTATIC.
    ACS_AUTOPLAY = &H4    ' Starts playing the animation as soon as the AVI clip is opened.
    ACS_TIMER = &H8       ' By default, the control creates a thread to play the AVI clip. If you set this flag, the control plays the clip without creating a thread; internally the control uses a Win32 timer to synchronize playback. Comctl32.dll version 6 and later: This style is not supported. By default, the control plays the AVI clip without creating a thread. Note: Comctl32.dll version 6 is not redistributable. To use Comctl32.dll version 6, specify it in a manifest. For more information on manifests, see Enabling Visual Styles.
End Enum

'PushButton, CheckBox, OptionButton
Public Enum EButtonStyle
Private Const BS_SOLID           As Long = 0&
Private Const BS_PUSHBUTTON      As Long = &H0&
Private Const BS_TEXT            As Long = &H0&
Private Const BS_NULL            As Long = 1&
Private Const BS_DEFPUSHBUTTON   As Long = &H1&
Private Const BS_HATCHED         As Long = 2&
Private Const BS_CHECKBOX        As Long = &H2&
Private Const BS_PATTERN         As Long = 3&
Private Const BS_AUTOCHECKBOX    As Long = &H3&
Private Const BS_INDEXED         As Long = 4&
Private Const BS_RADIOBUTTON     As Long = &H4&
Private Const BS_DIBPATTERN      As Long = 5&
Private Const BS_3STATE          As Long = &H5&
Private Const BS_DIBPATTERNPT    As Long = 6&
Private Const BS_AUTO3STATE      As Long = &H6&
Private Const BS_PATTERN8X8      As Long = 7&
Private Const BS_GROUPBOX        As Long = &H7&
Private Const BS_DIBPATTERN8X8   As Long = 8&
Private Const BS_USERBUTTON      As Long = &H8&
Private Const BS_MONOPATTERN     As Long = 9&
Private Const BS_AUTORADIOBUTTON As Long = &H9&
Private Const BS_OWNERDRAW       As Long = &HB&
Private Const BS_LEFTTEXT        As Long = &H20&
Private Const BS_RIGHTBUTTON     As Long = BS_LEFTTEXT
Private Const BS_ICON            As Long = &H40&
Private Const BS_BITMAP          As Long = &H80&
Private Const BS_LEFT            As Long = &H100&
Private Const BS_RIGHT           As Long = &H200&
Private Const BS_CENTER          As Long = &H300&
Private Const BS_TOP             As Long = &H400&
Private Const BS_BOTTOM          As Long = &H800&
Private Const BS_VCENTER         As Long = &HC00&
Private Const BS_HOLLOW          As Long = BS_NULL ' 1
Private Const BS_PUSHLIKE        As Long = &H1000&
Private Const BS_MULTILINE       As Long = &H2000&
Private Const BS_NOTIFY          As Long = &H4000&
Private Const BS_FLAT            As Long = &H8000&
End Enum
'ComboBox
Private Const CBS_SIMPLE            As Long = &H1&
Private Const CBS_DROPDOWN          As Long = &H2&
Private Const CBS_DROPDOWNLIST      As Long = &H3&
Private Const CBS_OWNERDRAWFIXED    As Long = &H10&
Private Const CBS_OWNERDRAWVARIABLE As Long = &H20&
Private Const CBS_AUTOHSCROLL       As Long = &H40&
Private Const CBS_OEMCONVERT        As Long = &H80&
Private Const CBS_SORT              As Long = &H100&
Private Const CBS_HASSTRINGS        As Long = &H200&
Private Const CBS_NOINTEGRALHEIGHT  As Long = &H400&
Private Const CBS_DISABLENOSCROLL   As Long = &H800&
Private Const CBS_UPPERCASE         As Long = &H2000&
Private Const CBS_LOWERCASE         As Long = &H4000&

'DateTimePicker
Private Const DTS_SHORTDATEFORMAT        As Long = &H0&
Private Const DTS_UPDOWN                 As Long = &H1&
Private Const DTS_SHOWNONE               As Long = &H2&
Private Const DTS_LONGDATEFORMAT         As Long = &H4&
Private Const DTS_TIMEFORMAT             As Long = &H9&
Private Const DTS_SHORTDATECENTURYFORMAT As Long = &HC&
Private Const DTS_APPCANPARSE            As Long = &H10&
Private Const DTS_RIGHTALIGN             As Long = &H20&

'ListBox, DragListBox
Private Const LBS_NOTIFY            As Long = &H1&
Private Const LBS_SORT              As Long = &H2&
Private Const LBS_NOREDRAW          As Long = &H4&
Private Const LBS_MULTIPLESEL       As Long = &H8&
Private Const LBS_OWNERDRAWFIXED    As Long = &H10&
Private Const LBS_OWNERDRAWVARIABLE As Long = &H20&
Private Const LBS_HASSTRINGS        As Long = &H40&
Private Const LBS_USETABSTOPS       As Long = &H80&
Private Const LBS_NOINTEGRALHEIGHT  As Long = &H100&
Private Const LBS_MULTICOLUMN       As Long = &H200&
Private Const LBS_WANTKEYBOARDINPUT As Long = &H400&
Private Const LBS_EXTENDEDSEL       As Long = &H800&
Private Const LBS_DISABLENOSCROLL   As Long = &H1000&
Private Const LBS_NODATA            As Long = &H2000&
Private Const LBS_NOSEL             As Long = &H4000&
Private Const LBS_STANDARD          As Long = (LBS_NOTIFY Or LBS_SORT Or WS_VSCROLL Or WS_BORDER)

'TextBox
Private Const ES_LEFT             As Long = &H0&
Private Const ES_CENTER           As Long = &H1&
Private Const ES_SYSTEM_REQUIRED  As Long = (&H1)
Private Const ES_DISPLAY_REQUIRED As Long = &H2&
Private Const ES_RIGHT            As Long = &H2&
Private Const ES_MULTILINE        As Long = &H4&
Private Const ES_USER_PRESENT     As Long = (&H4)
Private Const ES_NOOLEDRAGDROP    As Long = &H8&
Private Const ES_UPPERCASE        As Long = &H8&
Private Const ES_LOWERCASE        As Long = &H10&
Private Const ES_PASSWORD         As Long = &H20&
Private Const ES_AUTOVSCROLL      As Long = &H40&
Private Const ES_AUTOHSCROLL      As Long = &H80&
Private Const ES_NOHIDESEL        As Long = &H100&
Private Const ES_OEMCONVERT       As Long = &H400&
Private Const ES_READONLY         As Long = &H800&
Private Const ES_WANTRETURN       As Long = &H1000&
Private Const ES_DISABLENOSCROLL  As Long = &H2000&
Private Const ES_NUMBER           As Long = &H2000&
Private Const ES_SUNKEN           As Long = &H4000&
Private Const ES_SAVESEL          As Long = &H8000&
Private Const ES_SELFIME          As Long = &H40000
Private Const ES_NOIME            As Long = &H80000
Private Const ES_VERTICAL         As Long = &H400000
Private Const ES_EX_NOCALLOLEINIT As Long = &H1000000
Private Const ES_SELECTIONBAR     As Long = &H1000000
Private Const ES_CONTINUOUS       As Long = &H80000000

'Header
Private Const HDS_HORZ      As Long = &H0&
Private Const HDS_BUTTONS   As Long = &H2&
Private Const HDS_HOTTRACK  As Long = &H4&
Private Const HDS_HIDDEN    As Long = &H8&
Private Const HDS_DRAGDROP  As Long = &H40&
Private Const HDS_FULLDRAG  As Long = &H80&
Private Const HDS_FILTERBAR As Long = &H100&
Private Const HDS_FLAT      As Long = &H200&

'Hot Key, IP Address

'ListBox
Private Const LBS_NOTIFY            As Long = &H1&
Private Const LBS_SORT              As Long = &H2&
Private Const LBS_NOREDRAW          As Long = &H4&
Private Const LBS_MULTIPLESEL       As Long = &H8&
Private Const LBS_OWNERDRAWFIXED    As Long = &H10&
Private Const LBS_OWNERDRAWVARIABLE As Long = &H20&
Private Const LBS_HASSTRINGS        As Long = &H40&
Private Const LBS_USETABSTOPS       As Long = &H80&
Private Const LBS_NOINTEGRALHEIGHT  As Long = &H100&
Private Const LBS_MULTICOLUMN       As Long = &H200&
Private Const LBS_WANTKEYBOARDINPUT As Long = &H400&
Private Const LBS_EXTENDEDSEL       As Long = &H800&
Private Const LBS_DISABLENOSCROLL   As Long = &H1000&
Private Const LBS_NODATA            As Long = &H2000&
Private Const LBS_NOSEL             As Long = &H4000&
Private Const LBS_STANDARD          As Long = (LBS_NOTIFY Or LBS_SORT Or WS_VSCROLL Or WS_BORDER)

'ListView
Private Const LVS_ALIGNTOP            As Long = &H0&
Private Const LVS_ICON                As Long = &H0&
Private Const LVS_REPORT              As Long = &H1&
Private Const LVS_SMALLICON           As Long = &H2&
Private Const LVS_LIST                As Long = &H3&
Private Const LVS_TYPEMASK            As Long = &H3&
Private Const LVS_SINGLESEL           As Long = &H4&
Private Const LVS_SHOWSELALWAYS       As Long = &H8&
Private Const LVS_SORTASCENDING       As Long = &H10&
Private Const LVS_SORTDESCENDING      As Long = &H20&
Private Const LVS_SHAREIMAGELISTS     As Long = &H40&
Private Const LVS_NOLABELWRAP         As Long = &H80&
Private Const LVS_AUTOARRANGE         As Long = &H100&
Private Const LVS_EDITLABELS          As Long = &H200&
Private Const LVS_OWNERDRAWFIXED      As Long = &H400&
Private Const LVS_ALIGNLEFT           As Long = &H800&
Private Const LVS_ALIGNMASK           As Long = &HC00&
Private Const LVS_OWNERDATA           As Long = &H1000&
Private Const LVS_NOSCROLL            As Long = &H2000&
Private Const LVS_NOCOLUMNHEADER      As Long = &H4000&
Private Const LVS_NOSORTHEADER        As Long = &H8000&
Private Const LVS_TYPESTYLEMASK       As Long = &HFC00&
Private Const LVS_EX_GRIDLINES        As Long = &H1&
Private Const LVS_EX_SUBITEMIMAGES    As Long = &H2&
Private Const LVS_EX_CHECKBOXES       As Long = &H4&
Private Const LVS_EX_TRACKSELECT      As Long = &H8&
Private Const LVS_EX_HEADERDRAGDROP   As Long = &H10&
Private Const LVS_EX_FULLROWSELECT    As Long = &H20&
Private Const LVS_EX_ONECLICKACTIVATE As Long = &H40&
Private Const LVS_EX_TWOCLICKACTIVATE As Long = &H80&
Private Const LVS_EX_FLATSB           As Long = &H100&
Private Const LVS_EX_REGIONAL         As Long = &H200&
Private Const LVS_EX_INFOTIP          As Long = &H400&
Private Const LVS_EX_UNDERLINEHOT     As Long = &H800&
Private Const LVS_EX_UNDERLINECOLD    As Long = &H1000&
Private Const LVS_EX_MULTIWORKAREAS   As Long = &H2000&
Private Const LVS_EX_LABELTIP         As Long = &H4000&
Private Const LVS_EX_BORDERSELECT     As Long = &H8000&
Private Const LVS_EX_DOUBLEBUFFER     As Long = &H10000
Private Const LVS_EX_HIDELABELS       As Long = &H20000
Private Const LVS_EX_SINGLEROW        As Long = &H40000
Private Const LVS_EX_SNAPTOGRID       As Long = &H80000

'MonthCalendar
Private Const MCS_COMMAND_ENABLE            As Long = 13
Private Const MCS_COMMAND_DISABLE           As Long = 14
Private Const MCS_COMMAND_SET_CONFIG        As Long = 15
Private Const MCS_COMMAND_GET_CONFIG        As Long = 16
Private Const MCS_COMMAND_START             As Long = 17
Private Const MCS_COMMAND_STOP              As Long = 18
Private Const MCS_COMMAND_CONNECT           As Long = 19
Private Const MCS_COMMAND_RENAME            As Long = 20
Private Const MCS_COMMAND_REFRESH_STATUS    As Long = 21
Private Const MCS_CREATE_ONE_PER_NETCARD    As Long = &H1
Private Const MCS_CREATE_CONFIGS_BY_DEFAULT As Long = &H10
Private Const MCS_CREATE_PMODE_NOT_REQUIRED As Long = &H100
Private Const MCS_DAYSTATE                  As Long = &H1
Private Const MCS_MULTISELECT               As Long = &H2
Private Const MCS_WEEKNUMBERS               As Long = &H4
Private Const MCS_NOTODAYCIRCLE             As Long = &H8
Private Const MCS_NOTODAY                   As Long = &H10

'Pager
Private Const PGS_VERT       As Long = &H0
Private Const PGS_HORZ       As Long = &H1
Private Const PGS_AUTOSCROLL As Long = &H2
Private Const PGS_DRAGNDROP  As Long = &H4

'ProgressBar
Private Const PBS_SMOOTH        As Long = &H1
Private Const PBS_MARQUEE       As Long = &H2
Private Const PBS_VERTICAL      As Long = &H4
Private Const PBS_SMOOTHREVERSE As Long = &H8

'Rebar
Private Const PBS_SMOOTH          As Long = &H1&
Private Const PBS_VERTICAL        As Long = &H4&
Private Const RBS_TOOLTIPS        As Long = &H100&
Private Const RBS_VARHEIGHT       As Long = &H200&
Private Const RBS_BANDBORDERS     As Long = &H400&
Private Const RBS_FIXEDORDER      As Long = &H800&
Private Const RBS_REGISTERDROP    As Long = &H1000&
Private Const RBS_AUTOSIZE        As Long = &H2000&
Private Const RBS_VERTICALGRIPPER As Long = &H4000&
Private Const RBS_DBLCLKTOGGLE    As Long = &H8000&

'ScrollBar
Private Const SBS_HORZ                    As Long = &H0&
Private Const SBS_VERT                    As Long = &H1&
Private Const SBS_TOPALIGN                As Long = &H2&
Private Const SBS_SIZEBOXTOPLEFTALIGN     As Long = &H2&
Private Const SBS_LEFTALIGN               As Long = &H2&
Private Const SBS_RIGHTALIGN              As Long = &H4&
Private Const SBS_BOTTOMALIGN             As Long = &H4&
Private Const SBS_SIZEBOXBOTTOMRIGHTALIGN As Long = &H4&
Private Const SBS_SIZEBOX                 As Long = &H8&
Private Const SBS_SIZEGRIP                As Long = &H10&

'Static = Label
Private Const SS_LEFT            As Long = &H0&
Private Const SS_CENTER          As Long = &H1&
Private Const SS_RIGHT           As Long = &H2&
Private Const SS_ICON            As Long = &H3&
Private Const SS_BLACKRECT       As Long = &H4&
Private Const SS_GRAYRECT        As Long = &H5&
Private Const SS_WHITERECT       As Long = &H6&
Private Const SS_BLACKFRAME      As Long = &H7&
Private Const SS_GRAYFRAME       As Long = &H8&
Private Const SS_WHITEFRAME      As Long = &H9&
Private Const SS_USERITEM        As Long = &HA&
Private Const SS_SIMPLE          As Long = &HB&
Private Const SS_LEFTNOWORDWRAP  As Long = &HC&
Private Const SS_OWNERDRAW       As Long = &HD&
Private Const SS_BITMAP          As Long = &HE&
Private Const SS_ENHMETAFILE     As Long = &HF&
Private Const SS_ETCHEDHORZ      As Long = &H10&
Private Const SS_ETCHEDVERT      As Long = &H11&
Private Const SS_ETCHEDFRAME     As Long = &H12&
Private Const SS_TYPEMASK        As Long = &H1F&
Private Const SS_REALSIZECONTROL As Long = &H40&
Private Const SS_NOPREFIX        As Long = &H80&
Private Const SS_NOTIFY          As Long = &H100&
Private Const SS_CENTERIMAGE     As Long = &H200&
Private Const SS_RIGHTJUST       As Long = &H400&
Private Const SS_REALSIZEIMAGE   As Long = &H800&
Private Const SS_SUNKEN          As Long = &H1000&
Private Const SS_ELLIPSISMASK    As Long = &HC000&
Private Const SS_WORDELLIPSIS    As Long = &HC000&
Private Const SS_ENDELLIPSIS     As Long = &H4000&
Private Const SS_PATHELLIPSIS    As Long = &H8000&
Private Const SS_MAJOR_VERSION   As Long = 7
Private Const SS_MINOR_VERSION   As Long = 0
Private Const SS_LEVEL_VERSION   As Long = 0
Private Const SS_MINIMUM_VERSION As String = "7.00.00.0000"

'StatusBar
Private Const SBARS_SIZEGRIP   As Long = &H100&
Private Const SBARS_TOOLTIPS   As Long = &H800&

Private Const SBT_NOBORDERS    As Long = &H100&
Private Const SBT_POPOUT       As Long = &H200&
Private Const SBT_RTLREADING   As Long = &H400&
Private Const SBT_NOTABPARSING As Long = &H800&
Private Const SBT_TOOLTIPS     As Long = &H800&
Private Const SBT_OWNERDRAW    As Long = &H1000&

'SysLink
Private Const LWS_TRANSPARENT    As Long = 1
Private Const LWS_IGNORERETURN   As Long = 1
Private Const LWS_NOPREFIX       As Long = 1
Private Const LWS_USEVISUALSTYLE As Long = 1

'TabControl
Private Const TCS_RIGHTJUSTIFY      As Long = &H0
Private Const TCS_SINGLELINE        As Long = &H0
Private Const TCS_TABS              As Long = &H0
Private Const TCS_SCROLLOPPOSITE    As Long = &H1
Private Const TCS_EX_FLATSEPARATORS As Long = &H1
Private Const TCS_EX_REGISTERDROP   As Long = &H2
Private Const TCS_BOTTOM            As Long = &H2
Private Const TCS_RIGHT             As Long = &H2
Private Const TCS_MULTISELECT       As Long = &H4
Private Const TCS_FLATBUTTONS       As Long = &H8
Private Const TCS_FORCEICONLEFT     As Long = &H10
Private Const TCS_FORCELABELLEFT    As Long = &H20
Private Const TCS_HOTTRACK          As Long = &H40
Private Const TCS_VERTICAL          As Long = &H80
Private Const TCS_BUTTONS           As Long = &H100
Private Const TCS_MULTILINE         As Long = &H200
Private Const TCS_FIXEDWIDTH        As Long = &H400
Private Const TCS_RAGGEDRIGHT       As Long = &H800
Private Const TCS_FOCUSONBUTTONDOWN As Long = &H1000
Private Const TCS_OWNERDRAWFIXED    As Long = &H2000
Private Const TCS_TOOLTIPS          As Long = &H4000
Private Const TCS_FOCUSNEVER        As Long = &H8000

'Toolbar
Private Const TBSTYLE_BUTTON          As Long = &H0&
Private Const TBSTYLE_SEP             As Long = &H1&
Private Const TBSTYLE_EX_DRAWDDARROWS As Long = &H1&
Private Const TBSTYLE_CHECK           As Long = &H2&
Private Const TBSTYLE_GROUP           As Long = &H4&
Private Const TBSTYLE_DROPDOWN        As Long = &H8&
Private Const TBSTYLE_EX_MIXEDBUTTONS As Long = &H8&
Private Const TBSTYLE_AUTOSIZE        As Long = &H10&
Private Const TBSTYLE_EX_HIDECLIPPEDBUTTONS As Long = &H10&
Private Const TBSTYLE_NOPREFIX        As Long = &H20&
Private Const TBSTYLE_EX_DOUBLEBUFFER As Long = &H80&
Private Const TBSTYLE_TOOLTIPS        As Long = &H100&
Private Const TBSTYLE_WRAPABLE        As Long = &H200&
Private Const TBSTYLE_ALTDRAG         As Long = &H400&
Private Const TBSTYLE_FLAT            As Long = &H800&
Private Const TBSTYLE_LIST            As Long = &H1000&
Private Const TBSTYLE_CUSTOMERASE     As Long = &H2000&
Private Const TBSTYLE_REGISTERDROP    As Long = &H4000&
Private Const TBSTYLE_TRANSPARENT     As Long = &H8000&
Private Const TBSTYLE_CHECKGROUP      As Long = (TBSTYLE_GROUP Or TBSTYLE_CHECK)

'Tooltip
Private Const TTS_ALWAYSTIP As Long = &H1
Private Const TTS_NOPREFIX  As Long = &H2
Private Const TTS_NOANIMATE As Long = &H10
Private Const TTS_NOFADE    As Long = &H20
Private Const TTS_BALLOON   As Long = &H40

'Trackbar
Private Const TBS_RIGHT          As Long = &H0
Private Const TBS_HORZ           As Long = &H0
Private Const TBS_BOTTOM         As Long = &H0
Private Const TBS_AUTOTICKS      As Long = &H1
Private Const TBS_VERT           As Long = &H2
Private Const TBS_LEFT           As Long = &H4
Private Const TBS_TOP            As Long = &H4
Private Const TBS_BOTH           As Long = &H8
Private Const TBS_NOTICKS        As Long = &H10
Private Const TBS_ENABLESELRANGE As Long = &H20
Private Const TBS_FIXEDLENGTH    As Long = &H40
Private Const TBS_NOTHUMB        As Long = &H80
Private Const TBS_TOOLTIPS       As Long = &H100
Private Const TBS_REVERSED       As Long = &H200

'TreeView
Private Const TVS_HASBUTTONS      As Long = &H1&
Private Const TVS_HASLINES        As Long = &H2&
Private Const TVS_LINESATROOT     As Long = &H4&
Private Const TVS_EDITLABELS      As Long = &H8&
Private Const TVS_DISABLEDRAGDROP As Long = &H10&
Private Const TVS_SHOWSELALWAYS   As Long = &H20&
Private Const TVS_RTLREADING      As Long = &H40&
Private Const TVS_NOTOOLTIPS      As Long = &H80&
Private Const TVS_CHECKBOXES      As Long = &H100&
Private Const TVS_TRACKSELECT     As Long = &H200&
Private Const TVS_SINGLEEXPAND    As Long = &H400&
Private Const TVS_INFOTIP         As Long = &H800&
Private Const TVS_FULLROWSELECT   As Long = &H1000&
Private Const TVS_NOSCROLL        As Long = &H2000&
Private Const TVS_NONEVENHEIGHT   As Long = &H4000&
Private Const TVS_NOHSCROLL       As Long = &H8000&

'UpDown
Private Const UDS_WRAP        As Long = &H1
Private Const UDS_SETBUDDYINT As Long = &H2
Private Const UDS_ALIGNRIGHT  As Long = &H4
Private Const UDS_ALIGNLEFT   As Long = &H8
Private Const UDS_AUTOBUDDY   As Long = &H10
Private Const UDS_ARROWKEYS   As Long = &H20
Private Const UDS_HORZ        As Long = &H40
Private Const UDS_NOTHOUSANDS As Long = &H80
Private Const UDS_HOTTRACK    As Long = &H100


'dwStyle
'https://learn.microsoft.com/en-us/windows/win32/winmsg/window-styles
Public Enum EWndStyle
    WS_TILED = &H0&                 '  Das Fenster ist ein überlappende Fenster. Ein überlappendes Fenster hat eine Titelleiste und einen Rahmen. Identisch mit dem WS_OVERLAPPED Stil.
    WS_OVERLAPPED = &H0&            '  Das Fenster ist ein überlappende Fenster. Ein überlappendes Fenster hat eine Titelleiste und einen Rahmen. Identisch mit dem WS_TILED-Stil .

    WS_MAXIMIZEBOX = &H10000        '  Das Fenster verfügt über eine Schaltfläche zum Maximieren. Kann nicht mit dem WS_EX_CONTEXTHELP-Stil kombiniert werden. Die WS_SYSMENU Formatvorlage muss ebenfalls angegeben werden.
    WS_TABSTOP = &H10000            '  Das Fenster ist ein Steuerelement, das den Tastaturfokus erhalten kann, wenn der Benutzer die TAB-TASTE drückt. Durch Drücken der TAB-TASTE wird der Tastaturfokus auf das nächste Steuerelement mit der WS_TABSTOP-Formatvorlage geändert.Sie können diesen Stil aktivieren und deaktivieren, um die Navigation im Dialogfeld zu ändern. Um diesen Stil zu ändern, nachdem ein Fenster erstellt wurde, verwenden Sie die SetWindowLong-Funktion . Damit vom Benutzer erstellte Fenster und moduslose Dialoge mit Tabstopps funktionieren, ändern Sie die Nachrichtenschleife so, dass die IsDialogMessage-Funktion aufgerufen wird.
    WS_GROUP = &H20000              '  Das Fenster ist das erste Steuerelement einer Gruppe von Steuerelementen. Die Gruppe besteht aus diesem ersten Steuerelement und allen danach definierten Steuerelementen bis zum nächsten Steuerelement mit dem WS_GROUP Stil. Das erste Steuerelement in jeder Gruppe hat in der Regel den WS_TABSTOP Stil, sodass der Benutzer von Gruppe zu Gruppe wechseln kann. Der Benutzer kann anschließend den Tastaturfokus von einem Steuerelement in der Gruppe auf das nächste Steuerelement in der Gruppe ändern, indem er die Richtungstasten verwendet.Sie können diesen Stil aktivieren und deaktivieren, um die Navigation im Dialogfeld zu ändern. Um diesen Stil zu ändern, nachdem ein Fenster erstellt wurde, verwenden Sie die SetWindowLong-Funktion .
    WS_MINIMIZEBOX = &H20000        '  Das Fenster verfügt über eine Schaltfläche zum Minimieren. Kann nicht mit dem WS_EX_CONTEXTHELP-Stil kombiniert werden. Die WS_SYSMENU Formatvorlage muss ebenfalls angegeben werden.
    WS_SIZEBOX = &H40000            '  Das Fenster verfügt über einen Rahmen zur Größenanpassung. Identisch mit dem WS_THICKFRAME Stil.
    WS_THICKFRAME = &H40000         '  Das Fenster verfügt über einen Rahmen zur Größenanpassung. Identisch mit dem WS_SIZEBOX Stil.
    WS_SYSMENU = &H80000            '  Das Fenster verfügt über ein Fenstermenü auf der Titelleiste. Die WS_CAPTION Formatvorlage muss ebenfalls angegeben werden.

    WS_HSCROLL = &H100000           '  Das Fenster verfügt über eine horizontale Bildlaufleiste.
    WS_VSCROLL = &H200000           '  Das Fenster verfügt über eine vertikale Bildlaufleiste.
    WS_DLGFRAME = &H400000          '  Das Fenster verfügt über einen Rahmen eines Stils, der in der Regel mit Dialogfeldern verwendet wird. Ein Fenster mit dieser Formatvorlage kann keine Titelleiste aufweisen.
    WS_BORDER = &H800000            '  Das Fenster verfügt über einen dünnen Rahmen
    WS_CAPTION = &HC00000           '  Das Fenster verfügt über eine Titelleiste (einschließlich des WS_BORDER Stils).
                 
    WS_MAXIMIZE = &H1000000         '  Das Fenster wird anfänglich maximiert.
    WS_CLIPCHILDREN = &H2000000     '  Schließt den Bereich aus, der von untergeordneten Fenstern belegt wird, wenn das Zeichnen innerhalb des übergeordneten Fensters erfolgt. Diese Formatvorlage wird beim Erstellen des übergeordneten Fensters verwendet.
    WS_CLIPSIBLINGS = &H4000000     '  Schneidet untergeordnete Fenster relativ zueinander ab; Das heißt, wenn ein bestimmtes untergeordnetes Fenster eine WM_PAINT-Meldung empfängt, wird vom WS_CLIPSIBLINGS-Format alle anderen überlappenden untergeordneten Fenster aus dem Bereich des zu aktualisierenden untergeordneten Fensters heraus geklammert. Wenn WS_CLIPSIBLINGS nicht angegeben ist und sich untergeordnete Fenster überschneiden, ist es beim Zeichnen innerhalb des Clientbereichs eines untergeordneten Fensters möglich, innerhalb des Clientbereichs eines benachbarten untergeordneten Fensters zu zeichnen.
    WS_DISABLED = &H8000000         '  Das Fenster ist zunächst deaktiviert. Ein deaktiviertes Fenster kann keine Eingaben vom Benutzer empfangen. Um dies zu ändern, nachdem ein Fenster erstellt wurde, verwenden Sie die Funktion EnableWindow .

    WS_VISIBLE = &H10000000         '  Das Fenster ist zunächst sichtbar. Diese Formatvorlage kann mithilfe der ShowWindow - oder SetWindowPos-Funktion aktiviert und deaktiviert werden.
    WS_ICONIC = &H20000000          '  Das Fenster wird zunächst minimiert. Identisch mit dem WS_MINIMIZE Stil.
    WS_MINIMIZE = &H20000000        '  Das Fenster wird zunächst minimiert. Identisch mit dem WS_ICONIC Stil.
    WS_CHILD = &H40000000           '  Das Fenster ist ein untergeordnetes Fenster. Ein Fenster mit dieser Formatvorlage kann keine Menüleiste aufweisen. Diese Formatvorlage kann nicht mit der WS_POPUP-Formatvorlage verwendet werden.
    WS_CHILDWINDOW = &H40000000     '  Identisch mit dem WS_CHILD Stil.
    WS_POPUP = &H80000000           '  Das Fenster ist ein Popupfenster. Dieser Stil kann nicht mit dem WS_CHILD-Stil verwendet werden.
    WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)       'Das Fenster ist ein Popupfenster. Die Formatvorlagen WS_CAPTION und WS_POPUPWINDOW müssen kombiniert werden, um das Fenstermenü sichtbar zu machen.
    WS_TILEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)           ' Das Fenster ist ein überlappende Fenster. Identisch mit dem WS_OVERLAPPEDWINDOW Stil.
    WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX) 'Das Fenster ist ein überlappende Fenster. Identisch mit dem WS_TILEDWINDOW Stil.
    
    VBFormStyle_BSNone = WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
    VBFormStyle_FixedSingle = WS_SYSMENU Or WS_DLGFRAME Or WS_BORDER Or WS_CAPTION Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
    VBFormStyle_Sizable = WS_MAXIMIZEBOX Or WS_TABSTOP Or WS_GROUP Or WS_MINIMIZEBOX Or WS_SIZEBOX Or WS_THICKFRAME Or WS_SYSMENU Or WS_DLGFRAME Or WS_BORDER Or WS_CAPTION Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
    VBFormStyle_FixedDialog = WS_SYSMENU Or WS_DLGFRAME Or WS_BORDER Or WS_CAPTION Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
    VBFormStyle_FixedToolWindow = WS_SYSMENU Or WS_DLGFRAME Or WS_BORDER Or WS_CAPTION Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
    VBFormStyle_SizableToolWindow = WS_SIZEBOX Or WS_THICKFRAME Or WS_SYSMENU Or WS_DLGFRAME Or WS_BORDER Or WS_CAPTION Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE

'BorderStyle: vbBSNone = 0
'    Style:   WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
'    StyleEx:
'
'BorderStyle: vbFixedSingle = 1
'    Style:   WS_SYSMENU Or WS_DLGFRAME Or WS_BORDER Or WS_CAPTION Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
'    StyleEx: WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW
'
'BorderStyle: vbSizable = 2
'    Style:   WS_MAXIMIZEBOX Or WS_TABSTOP Or WS_GROUP Or WS_MINIMIZEBOX Or WS_SIZEBOX Or WS_THICKFRAME Or WS_SYSMENU Or WS_DLGFRAME Or WS_BORDER Or WS_CAPTION Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
'    StyleEx: WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW
'
'BorderStyle: vbFixedDialog = 3
'    Style:   WS_SYSMENU Or WS_DLGFRAME Or WS_BORDER Or WS_CAPTION Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
'    StyleEx: WS_EX_DLGMODALFRAME Or WS_EX_WINDOWEDGE
'
'BorderStyle: vbFixedToolWindow = 4
'    Style:   WS_SYSMENU Or WS_DLGFRAME Or WS_BORDER Or WS_CAPTION Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
'    StyleEx: WS_EX_TOOLWINDOW Or WS_EX_WINDOWEDGE
'
'BorderStyle: vbSizableToolWindow = 5
'    Style:   WS_SIZEBOX Or WS_THICKFRAME Or WS_SYSMENU Or WS_DLGFRAME Or WS_BORDER Or WS_CAPTION Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
'    StyleEx: WS_EX_TOOLWINDOW Or WS_EX_WINDOWEDGE
        
End Enum

'dwExStyle
'https://learn.microsoft.com/en-us/windows/win32/winmsg/extended-window-styles
Public Enum EWndStyleEx
    WS_EX_LEFT = &H0&                      '   Das Fenster verfügt über generische linksbündige Eigenschaften. Dies ist die Standardeinstellung.
    WS_EX_LTRREADING = &H0&                '   Der Fenstertext wird mithilfe von Eigenschaften der Lesereihenfolge von links nach rechts angezeigt. Dies ist die Standardeinstellung.
    WS_EX_RIGHTSCROLLBAR = &H0&            '   Die vertikale Bildlaufleiste (sofern vorhanden) befindet sich rechts neben dem Clientbereich. Dies ist die Standardeinstellung.
    WS_EX_DLGMODALFRAME = &H1&             '   Das Fenster hat einen doppelten Rahmen; Das Fenster kann optional mit einer Titelleiste erstellt werden, indem die WS_CAPTION-Formatvorlage im dwStyle-Parameter angegeben wird.

    WS_EX_NOPARENTNOTIFY = &H4&            '   Das untergeordnete Fenster, das mit dieser Formatvorlage erstellt wurde, sendet die WM_PARENTNOTIFY Nachricht nicht an das übergeordnete Fenster, wenn es erstellt oder zerstört wird.
    WS_EX_TOPMOST = &H8&                   '   Das Fenster sollte über allen nicht obersten Fenstern platziert werden und darüber bleiben, auch wenn das Fenster deaktiviert ist. Verwenden Sie zum Hinzufügen oder Entfernen dieses Stils die SetWindowPos-Funktion.

    WS_EX_ACCEPTFILES = &H10&              '   Das Fenster akzeptiert Drag-Drop-Dateien.
    WS_EX_TRANSPARENT = &H20&              '   Das Fenster sollte erst gezeichnet werden, wenn gleichgeordnete Elemente unter dem Fenster (die von demselben Faden erstellt wurden) gezeichnet wurden. Das Fenster wird transparent angezeigt, da die Bits der zugrunde liegenden gleichgeordneten Fenster bereits gezeichnet wurden. Um Transparenz ohne diese Einschränkungen zu erzielen, verwenden Sie die SetWindowRgn-Funktion .
    WS_EX_MDICHILD = &H40&                 '   Das Fenster ist ein untergeordnetes MDI-Fenster.
    WS_EX_TOOLWINDOW = &H80&               '   Das Fenster soll als unverankerte Symbolleiste verwendet werden. Ein Toolfenster hat eine Titelleiste, die kürzer ist als eine normale Titelleiste, und der Fenstertitel wird mit einer kleineren Schriftart gezeichnet. Ein Toolfenster wird nicht in der Taskleiste oder im Dialogfeld angezeigt, das angezeigt wird, wenn der Benutzer ALT+TAB drückt. Wenn ein Toolfenster über ein Systemmenü verfügt, wird sein Symbol nicht auf der Titelleiste angezeigt. Sie können das Systemmenü jedoch anzeigen, indem Sie mit der rechten Maustaste klicken oder ALT+LEERZEICHEN eingeben.

    WS_EX_WINDOWEDGE = &H100&              '   Das Fenster hat einen Rahmen mit einer erhöhten Kante.
    WS_EX_CLIENTEDGE = &H200&              '   Das Fenster hat einen Rahmen mit einem gesunkenen Rand.
    WS_EX_CONTEXTHELP = &H400&             '   Die Titelleiste des Fensters enthält ein Fragezeichen. Wenn der Benutzer auf das Fragezeichen klickt, wird der Cursor zu einem Fragezeichen geändert. Wenn der Benutzer dann auf ein untergeordnetes Fenster klickt, erhält das untergeordnete Element eine WM_HELP Nachricht. Das untergeordnete Fenster sollte die Nachricht an die Prozedur des übergeordneten Fensters übergeben, die die WinHelp-Funktion mithilfe des Befehls HELP_WM_HELP aufrufen sollte. Die Hilfeanwendung zeigt ein Popupfenster an, das normalerweise Hilfe für das untergeordnete Fenster enthält. WS_EX_CONTEXTHELP können nicht mit dem format WS_MAXIMIZEBOX oder WS_MINIMIZEBOX verwendet werden.

    WS_EX_RIGHT = &H1000&                  '   Das Fenster verfügt über generische "rechtsbündige" Eigenschaften. Dies hängt von der Fensterklasse ab. Diese Formatvorlage wirkt sich nur dann aus, wenn die Shellsprache Hebräisch, Arabisch oder eine andere Sprache ist, die die Ausrichtung der Lesereihenfolge unterstützt. andernfalls wird die Formatvorlage ignoriert. Die Verwendung des WS_EX_RIGHT-Stils für statische Steuerelemente oder Bearbeitungssteuerelemente hat die gleiche Auswirkung wie die Verwendung des SS_RIGHT bzw . ES_RIGHT Stils. Die Verwendung dieses Stils mit Schaltflächensteuerelementen hat die gleiche Auswirkung wie die Verwendung von BS_RIGHT und BS_RIGHTBUTTON Formatvorlagen.
    WS_EX_RTLREADING = &H2000&             '   Wenn die Shellsprache Hebräisch, Arabisch oder eine andere Sprache ist, die die Ausrichtung der Lesereihenfolge unterstützt, wird der Fenstertext mithilfe von Eigenschaften der Lesereihenfolge von rechts nach links angezeigt. Bei anderen Sprachen wird der Stil ignoriert.
    WS_EX_LEFTSCROLLBAR = &H4000&          '   Wenn die Shellsprache Hebräisch, Arabisch oder eine andere Sprache ist, die die Ausrichtung der Lesereihenfolge unterstützt, befindet sich die vertikale Scrollleiste (falls vorhanden) links vom Clientbereich. Bei anderen Sprachen wird die Formatvorlage ignoriert.

    WS_EX_CONTROLPARENT = &H10000          '   Das Fenster selbst enthält untergeordnete Fenster, die an der Navigation im Dialogfeld teilnehmen sollen. Wenn diese Formatvorlage angegeben ist, wird der Dialog-Manager in untergeordnete Elemente dieses Fensters zurückgesetzt, wenn Navigationsvorgänge wie die TAB-TASTE, eine Pfeiltaste oder eine mnemonische Tastatur ausgeführt werden.
    WS_EX_STATICEDGE = &H20000             '   Das Fenster verfügt über eine dreidimensionale Rahmenart, die für Elemente verwendet werden soll, die keine Benutzereingaben akzeptieren.
    WS_EX_APPWINDOW = &H40000              '   Erzwingt ein Fenster der obersten Ebene auf der Taskleiste, wenn das Fenster sichtbar ist.
    WS_EX_LAYERED = &H80000                '   Das Fenster ist ein mehrschichtiges Fenster. Diese Formatvorlage kann nicht verwendet werden, wenn das Fenster eine Klassenart von CS_OWNDC oder CS_CLASSDC aufweist. Windows 8: Der WS_EX_LAYERED-Stil wird für Fenster der obersten Ebene und untergeordnete Fenster unterstützt. Frühere Windows-Versionen unterstützen WS_EX_LAYERED nur für Fenster der obersten Ebene.

    WS_EX_NOINHERITLAYOUT = &H100000       '   Das Fenster gibt sein Fensterlayout nicht an die untergeordneten Fenster weiter.
    WS_EX_NOREDIRECTIONBITMAP = &H200000   '   Das Fenster wird nicht auf eine Umleitungsoberfläche gerendert. Dies gilt für Fenster, die keinen sichtbaren Inhalt haben oder andere Mechanismen als Oberflächen verwenden, um ihr Visuelles bereitzustellen.
    WS_EX_LAYOUTRTL = &H400000             '   Wenn die Shellsprache Hebräisch, Arabisch oder eine andere Sprache ist, die die Ausrichtung der Lesereihenfolge unterstützt, befindet sich der horizontale Ursprung des Fensters am rechten Rand. Zunehmende horizontale Werte gehen nach links vor.

    WS_EX_COMPOSITED = &H2000000           '   Zeichnet alle untergeordneten Elemente eines Fensters in der Reihenfolge von unten nach oben mit Doppelpufferung. Die Unter-nach-Oben-Malreihenfolge ermöglicht es einem absteigenden Fenster, Transluzenzeffekte (Alpha) und Transparenzeffekte (Farbtaste) zu erhalten, aber nur, wenn im absteigenden Fenster auch das WS_EX_TRANSPARENT Bit festgelegt ist. Durch die Doppelpufferung können das Fenster und seine Absteigenden ohne Flackern bemalt werden. Dies kann nicht verwendet werden, wenn das Fenster eine Klassenart von CS_OWNDC oder CS_CLASSDC aufweist. Windows 2000: Dieser Stil wird nicht unterstützt.
    WS_EX_NOACTIVATE = &H8000000           '   Ein Fenster der obersten Ebene, das mit dieser Formatvorlage erstellt wurde, wird nicht zum Vordergrundfenster, wenn der Benutzer darauf klickt. Das System bringt dieses Fenster nicht in den Vordergrund, wenn der Benutzer das Vordergrundfenster minimiert oder schließt. Das Fenster sollte nicht durch programmgesteuerten Zugriff oder über die Tastaturnavigation durch barrierefreie Technologien wie die Sprachausgabe aktiviert werden. Verwenden Sie zum Aktivieren des Fensters die Funktion SetActiveWindow oder SetForegroundWindow . Das Fenster wird standardmäßig nicht auf der Taskleiste angezeigt. Um zu erzwingen, dass das Fenster auf der Taskleiste angezeigt wird, verwenden Sie die WS_EX_APPWINDOW Stil.
    WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)                     '   Das Fenster ist ein überlappende Fenster.
    WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)       '   Das Fenster ist ein Palettenfenster, bei dem es sich um ein modusloses Dialogfeld handelt, in dem ein Array von Befehlen angezeigt wird.
    
    VBFormStyleEx_BSNone = 0
    VBFormStyleEx_FixedSingle = WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW
    VBFormStyleEx_Sizable = WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW
    VBFormStyleEx_FixedDialog = WS_EX_DLGMODALFRAME Or WS_EX_WINDOWEDGE
    VBFormStyleEx_FixedToolWindow = WS_EX_TOOLWINDOW Or WS_EX_WINDOWEDGE
    VBFormStyleEx_SizableToolWindow = WS_EX_TOOLWINDOW Or WS_EX_WINDOWEDGE
    
End Enum
Private Const GWL_STYLE      As Long = -16&  ' Ruft die Fensterstile ab.
Private Const GWL_EXSTYLE    As Long = -20&  ' Ruft die erweiterten Fensterstile ab.
Private Declare Function GetWindowLongW Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

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

'Nur für Controls
Public Function EWndStyle_ToStr(ByVal e As EWndStyle) As String
    Dim sOr As String: sOr = " Or " 'vbTab
    Dim s As String, e2 As EWndStyle: e2 = e
    If e And WS_TILED Then s = s & IIf(Len(s), sOr, "") & "WS_TILED":               e2 = e2 Xor WS_TILED
    If e And WS_OVERLAPPED Then s = s & IIf(Len(s), sOr, "") & "WS_OVERLAPPED":     e2 = e2 Xor WS_OVERLAPPED

    'If e And WS_MAXIMIZEBOX Then s = s & IIf(Len(s), sOr, "") & "WS_MAXIMIZEBOX":   e2 = e2 xOr WS_MAXIMIZEBOX
    If e And WS_TABSTOP Then s = s & IIf(Len(s), sOr, "") & "WS_TABSTOP":           e2 = e2 Xor WS_TABSTOP
    If e And WS_GROUP Then s = s & IIf(Len(s), sOr, "") & "WS_GROUP":               e2 = e2 Xor WS_GROUP
    'If e And WS_MINIMIZEBOX Then s = s & IIf(Len(s), sOr, "") & "WS_MINIMIZEBOX":   e2 = e2 xOr WS_MINIMIZEBOX
    'If e And WS_SIZEBOX Then s = s & IIf(Len(s), sOr, "") & "WS_SIZEBOX":           e2 = e2 xOr WS_SIZEBOX
    If e And WS_THICKFRAME Then s = s & IIf(Len(s), sOr, "") & "WS_THICKFRAME":     e2 = e2 Xor WS_THICKFRAME
    'If e And WS_SYSMENU Then s = s & IIf(Len(s), sOr, "") & "WS_SYSMENU":           e2 = e2 xOr WS_SYSMENU

    If e And WS_HSCROLL Then s = s & IIf(Len(s), sOr, "") & "WS_HSCROLL":           e2 = e2 Xor WS_HSCROLL
    If e And WS_VSCROLL Then s = s & IIf(Len(s), sOr, "") & "WS_VSCROLL":           e2 = e2 Xor WS_VSCROLL
    If e And WS_DLGFRAME Then s = s & IIf(Len(s), sOr, "") & "WS_DLGFRAME":         e2 = e2 Xor WS_DLGFRAME
    If e And WS_BORDER Then s = s & IIf(Len(s), sOr, "") & "WS_BORDER":             e2 = e2 Xor WS_BORDER
    'If e And WS_CAPTION Then s = s & IIf(Len(s), sOr, "") & "WS_CAPTION":           e2 = e2 Xor WS_CAPTION

    'If e And WS_MAXIMIZE Then s = s & IIf(Len(s), sOr, "") & "WS_MAXIMIZE":         e2 = e2 Or WS_MAXIMIZE
    If e And WS_CLIPCHILDREN Then s = s & IIf(Len(s), sOr, "") & "WS_CLIPCHILDREN": e2 = e2 Xor WS_CLIPCHILDREN
    If e And WS_CLIPSIBLINGS Then s = s & IIf(Len(s), sOr, "") & "WS_CLIPSIBLINGS": e2 = e2 Xor WS_CLIPSIBLINGS
    If e And WS_DISABLED Then s = s & IIf(Len(s), sOr, "") & "WS_DISABLED":         e2 = e2 Xor WS_DISABLED

    If e And WS_VISIBLE Then s = s & IIf(Len(s), sOr, "") & "WS_VISIBLE":           e2 = e2 Xor WS_VISIBLE
    If e And WS_ICONIC Then s = s & IIf(Len(s), sOr, "") & "WS_ICONIC":             e2 = e2 Xor WS_ICONIC
    'If e And WS_MINIMIZE Then s = s & IIf(Len(s), sOr, "") & "WS_MINIMIZE":         e2 = e2 Or WS_MINIMIZE
    If e And WS_CHILD Then s = s & IIf(Len(s), sOr, "") & "WS_CHILD":               e2 = e2 Xor WS_CHILD
    'If e And WS_CHILDWINDOW Then s = s & IIf(Len(s), sOr, "") & "WS_CHILDWINDOW":   e2 = e2 Or WS_CHILDWINDOW
    If e And WS_POPUP Then s = s & IIf(Len(s), sOr, "") & "WS_POPUP":               e2 = e2 Xor WS_POPUP
    
    EWndStyle_ToStr = "&H" & Hex(e) & " = " & s & IIf(e2, sOr & "&H" & Hex(e2), "")
End Function

'Nur für Controls
Public Function EWndStyleEx_ToStr(ByVal e As EWndStyleEx) As String
    Dim sOr As String: sOr = " Or " ' vbTab
    Dim s As String, e2 As EWndStyleEx: e2 = e
    If e And WS_EX_LEFT Then s = s & IIf(Len(s), sOr, "") & "WS_EX_LEFT":                       e2 = e2 Xor WS_EX_LEFT
    If e And WS_EX_LTRREADING Then s = s & IIf(Len(s), sOr, "") & "WS_EX_LTRREADING":           e2 = e2 Xor WS_EX_LTRREADING
    If e And WS_EX_RIGHTSCROLLBAR Then s = s & IIf(Len(s), sOr, "") & "WS_EX_RIGHTSCROLLBAR":   e2 = e2 Xor WS_EX_RIGHTSCROLLBAR
    If e And WS_EX_DLGMODALFRAME Then s = s & IIf(Len(s), sOr, "") & "WS_EX_DLGMODALFRAME":     e2 = e2 Xor WS_EX_DLGMODALFRAME
    
    If e And WS_EX_NOPARENTNOTIFY Then s = s & IIf(Len(s), sOr, "") & "WS_EX_NOPARENTNOTIFY":   e2 = e2 Xor WS_EX_NOPARENTNOTIFY
    If e And WS_EX_TOPMOST Then s = s & IIf(Len(s), sOr, "") & "WS_EX_TOPMOST":                 e2 = e2 Xor WS_EX_TOPMOST
    
    If e And WS_EX_ACCEPTFILES Then s = s & IIf(Len(s), sOr, "") & "WS_EX_ACCEPTFILES":         e2 = e2 Xor WS_EX_ACCEPTFILES
    If e And WS_EX_TRANSPARENT Then s = s & IIf(Len(s), sOr, "") & "WS_EX_TRANSPARENT":         e2 = e2 Xor WS_EX_TRANSPARENT
    If e And WS_EX_MDICHILD Then s = s & IIf(Len(s), sOr, "") & "WS_EX_MDICHILD":               e2 = e2 Xor WS_EX_MDICHILD
    If e And WS_EX_TOOLWINDOW Then s = s & IIf(Len(s), sOr, "") & "WS_EX_TOOLWINDOW":           e2 = e2 Xor WS_EX_TOOLWINDOW
    
    If e And WS_EX_WINDOWEDGE Then s = s & IIf(Len(s), sOr, "") & "WS_EX_WINDOWEDGE":           e2 = e2 Xor WS_EX_WINDOWEDGE
    If e And WS_EX_CLIENTEDGE Then s = s & IIf(Len(s), sOr, "") & "WS_EX_CLIENTEDGE":           e2 = e2 Xor WS_EX_CLIENTEDGE
    If e And WS_EX_CONTEXTHELP Then s = s & IIf(Len(s), sOr, "") & "WS_EX_CONTEXTHELP":         e2 = e2 Xor WS_EX_CONTEXTHELP
    
    If e And WS_EX_RIGHT Then s = s & IIf(Len(s), sOr, "") & "WS_EX_RIGHT":                     e2 = e2 Xor WS_EX_RIGHT
    If e And WS_EX_RTLREADING Then s = s & IIf(Len(s), sOr, "") & "WS_EX_RTLREADING":           e2 = e2 Xor WS_EX_RTLREADING
    If e And WS_EX_LEFTSCROLLBAR Then s = s & IIf(Len(s), sOr, "") & "WS_EX_LEFTSCROLLBAR":     e2 = e2 Xor WS_EX_LEFTSCROLLBAR
    
    If e And WS_EX_CONTROLPARENT Then s = s & IIf(Len(s), sOr, "") & "WS_EX_CONTROLPARENT":     e2 = e2 Xor WS_EX_CONTROLPARENT
    If e And WS_EX_STATICEDGE Then s = s & IIf(Len(s), sOr, "") & "WS_EX_STATICEDGE":           e2 = e2 Xor WS_EX_STATICEDGE
    If e And WS_EX_APPWINDOW Then s = s & IIf(Len(s), sOr, "") & "WS_EX_APPWINDOW":             e2 = e2 Xor WS_EX_APPWINDOW
    If e And WS_EX_LAYERED Then s = s & IIf(Len(s), sOr, "") & "WS_EX_LAYERED":                 e2 = e2 Xor WS_EX_LAYERED
    
    If e And WS_EX_NOINHERITLAYOUT Then s = s & IIf(Len(s), sOr, "") & "WS_EX_NOINHERITLAYOUT": e2 = e2 Xor WS_EX_NOINHERITLAYOUT
    If e And WS_EX_NOREDIRECTIONBITMAP Then s = s & IIf(Len(s), sOr, "") & "WS_EX_NOREDIRECTIONBITMAP": e2 = e2 Xor WS_EX_NOREDIRECTIONBITMAP
    If e And WS_EX_LAYOUTRTL Then s = s & IIf(Len(s), sOr, "") & "WS_EX_LAYOUTRTL":             e2 = e2 Xor WS_EX_LAYOUTRTL
    
    If e And WS_EX_COMPOSITED Then s = s & IIf(Len(s), sOr, "") & "WS_EX_COMPOSITED":           e2 = e2 Xor WS_EX_COMPOSITED
    If e And WS_EX_NOACTIVATE Then s = s & IIf(Len(s), sOr, "") & "WS_EX_NOACTIVATE":           e2 = e2 Xor WS_EX_NOACTIVATE
    EWndStyleEx_ToStr = "&H" & Hex(e) & " = " & s & IIf(e2, sOr & "&H" & Hex(e2), "")
End Function
