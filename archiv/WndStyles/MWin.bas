Attribute VB_Name = "MWin"
Option Explicit

#If VBA7 = 0 Then
    Public Enum LongPtr
        [_]
    End Enum
#End If

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
