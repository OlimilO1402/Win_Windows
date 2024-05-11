VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EClsStyle
    CS_VREDRAW = &H1&             '    Zeichnet das gesamte Fenster neu, wenn eine Bewegung oder Größenanpassung die Höhe des Clientbereichs ändert.
    CS_HREDRAW = &H2&             '    Zeichnet das gesamte Fenster neu, wenn durch eine Verschiebung oder Größenanpassung die Breite des Clientbereichs geändert wird.
    CS_DBLCLKS = &H8&             '    Sendet eine Doppelklicknachricht an die Fensterprozedur, wenn der Benutzer mit der Maus doppelklicken, während sich der Cursor in einem Fenster befindet, das zur -Klasse gehört.
    CS_OWNDC = &H20&              '    Weist jedem Fenster in der -Klasse einen eindeutigen Gerätekontext zu.
    CS_CLASSDC = &H40&            '    Weist einen Gerätekontext zu, der von allen Fenstern in der -Klasse gemeinsam verwendet werden soll. Da Fensterklassen prozessspezifisch sind, ist es möglich, dass mehrere Threads einer Anwendung ein Fenster derselben Klasse erstellen. Es ist auch möglich, dass die Threads versuchen, den Gerätekontext gleichzeitig zu verwenden. In diesem Fall lässt das System nur einen Thread zu, um den Zeichnungsvorgang erfolgreich abzuschließen.
    CS_PARENTDC = &H80&           '    Legt das Beschneidungsrechteck des untergeordneten Fensters auf das des übergeordneten Fensters fest, damit das untergeordnete Element auf dem übergeordneten Fenster zeichnen kann. Ein Fenster mit dem CS_PARENTDC Stilbits empfängt einen regulären Gerätekontext aus dem Gerätekontextcache des Systems. Es gibt dem untergeordneten Element nicht die Gerätekontext- oder Gerätekontexteinstellungen des übergeordneten Elements. Die Angabe CS_PARENTDC verbessert die Leistung einer Anwendung.
    CS_NOCLOSE = &H200&           '    Deaktiviert Schließen im Fenstermenü.
    CS_SAVEBITS = &H800&          '    Speichert als Bitmap den Teil des Bildschirmbilds, der von einem Fenster dieser Klasse verdeckt wird. Wenn das Fenster entfernt wird, verwendet das System die gespeicherte Bitmap, um das Bildschirmbild wiederherzustellen, einschließlich anderer verdeckter Fenster. Daher sendet das System keine WM_PAINT Nachrichten an Fenster, die verdeckt wurden, wenn der von der Bitmap verwendete Arbeitsspeicher nicht verworfen wurde und wenn das gespeicherte Bild durch andere Bildschirmaktionen nicht ungültig gemacht wurde. Dieser Stil ist nützlich für kleine Fenster (z. B. Menüs oder Dialogfelder), die kurz angezeigt und dann entfernt werden, bevor andere Bildschirmaktivitäten stattfinden. Durch diese Formatvorlage wird die zum Anzeigen des Fensters erforderliche Zeit erhöht, da das System zunächst Arbeitsspeicher zum Speichern der Bitmap zuweisen muss.
    CS_BYTEALIGNCLIENT = &H1000&  '    Richtet den Clientbereich des Fensters an einer Bytegrenze (in x-Richtung) aus. Diese Formatvorlage wirkt sich auf die Breite des Fensters und seine horizontale Platzierung auf der Anzeige aus.
    CS_BYTEALIGNWINDOW = &H2000&  '    Richtet das Fenster an einer Bytegrenze (in x-Richtung) aus. Diese Formatvorlage wirkt sich auf die Breite des Fensters und seine horizontale Platzierung auf der Anzeige aus.
    CS_GLOBALCLASS = &H4000&      '    Gibt an, dass die Fensterklasse eine globale Anwendungsklasse ist. Weitere Informationen finden Sie im Abschnitt "Globale Anwendungsklassen" unter Informationen zu Fensterklassen.
    CS_DROPSHADOW = &H20000       '    Aktiviert den Schlagschatteneffekt für ein Fenster. Der Effekt wird über SPI_SETDROPSHADOW ein- und ausgeschaltet. In der Regel ist dies für kleine, kurzlebige Fenster wie Menüs aktiviert, um ihre Z-Reihenfolge-Beziehung zu anderen Fenstern hervorzuheben. Fenster, die aus einer Klasse mit diesem Stil erstellt werden, müssen Fenster der obersten Ebene sein. Sie dürfen keine Kinderfenster sein.
    VBFormClassStyleDefault = CS_VREDRAW Or CS_PARENTDC Or CS_OWNDC Or CS_HREDRAW
End Enum

'https://learn.microsoft.com/de-de/windows/win32/api/winuser/ns-winuser-wndclassexw
'typedef struct tagWNDCLASSEXW {
'  UINT      cbSize;
'  UINT      style;
'  WNDPROC   lpfnWndProc;
'  int       cbClsExtra;
'  int       cbWndExtra;
'  HINSTANCE hInstance;
'  HICON     hIcon;
'  HCURSOR   hCursor;
'  HBRUSH    hbrBackground;
'  LPCWSTR   lpszMenuName;
'  LPCWSTR   lpszClassName;
'  HICON     hIconSm;
'} WNDCLASSEXW, *PWNDCLASSEXW, *NPWNDCLASSEXW, *LPWNDCLASSEXW;
Private Type WNDCLASSEXW
    cbSize        As Long      'UINT
    Style         As EClsStyle 'UINT
    lpfnWndProc   As LongPtr   'WNDPROC
    cbClsExtra    As Long      'int
    cbWndExtra    As Long      'int
    hInstance     As LongPtr   'HINSTANCE
    hIcon         As LongPtr   'HICON
    hCursor       As LongPtr   'HCURSOR
    hbrBackground As LongPtr   'HBRUSH
    lpszMenuName  As LongPtr   'LPCSTR
    lpszClassName As LongPtr   'LPCSTR
    hIconSm       As LongPtr   'HICON the small icon
End Type '} WNDCLASSEXA, *PWNDCLASSEXA, *NPWNDCLASSEXA, *LPWNDCLASSEXA;

'dwStyle
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

Public Enum EWndShowZPos
    SWP_NOSIZE = &H1&            '    Behält die aktuelle Größe bei (ignoriert die Parameter cx und cy ).
    SWP_NOMOVE = &H2&            '    Behält die aktuelle Position bei (ignoriert die X - und Y-Parameter ).
    SWP_NOZORDER = &H4&          '    Behält die aktuelle Z-Reihenfolge bei (ignoriert den Parameter hWndInsertAfter ).
    SWP_NOREDRAW = &H8&          '    Änderungen werden nicht neu gezeichnet. Wenn dieses Flag festgelegt ist, erfolgt keine Neulackierung jeglicher Art. Dies gilt für den Clientbereich, den Nichtclientbereich (einschließlich Titelleiste und Bildlaufleisten) und alle Teile des übergeordneten Fensters, die aufgrund des Verschiebens des Fensters aufgedeckt wurden. Wenn dieses Flag festgelegt ist, muss die Anwendung alle Teile des Fensters und des übergeordneten Fensters, die neu ausgezeichnet werden müssen, explizit für ungültig erklären oder neu zeichnen.
    SWP_NOACTIVATE = &H10&       '    Aktiviert das Fenster nicht. Wenn dieses Flag nicht festgelegt ist, wird das Fenster aktiviert und an den Anfang der obersten oder nicht obersten Gruppe verschoben (abhängig von der Einstellung des hWndInsertAfter-Parameters ).
    SWP_DRAWFRAME = &H20&        '    Zeichnet einen Rahmen (definiert in der Klassenbeschreibung des Fensters) um das Fenster.
    SWP_FRAMECHANGED = &H20&     '    Wendet mithilfe der SetWindowLong-Funktion neue Frameformatvorlagen an. Sendet eine WM_NCCALCSIZE Nachricht an das Fenster, auch wenn die Größe des Fensters nicht geändert wird. Wenn dieses Flag nicht angegeben ist, wird WM_NCCALCSIZE nur gesendet, wenn die Größe des Fensters geändert wird.
    SWP_SHOWWINDOW = &H40&       '    Zeigt das Fenster an.
    SWP_HIDEWINDOW = &H80&       '    Blendet das Fenster aus.
    SWP_NOCOPYBITS = &H100&      '    Verwirft den gesamten Inhalt des Clientbereichs. Wenn dieses Flag nicht angegeben ist, werden die gültigen Inhalte des Clientbereichs gespeichert und wieder in den Clientbereich kopiert, nachdem das Fenster dimensioniert oder neu positioniert wurde.
    SWP_NOOWNERZORDER = &H200&   '    Ändert die Position des Besitzerfensters in der Z-Reihenfolge nicht.
    SWP_NOREPOSITION = &H200&    '    Identisch mit dem flag SWP_NOOWNERZORDER .
    SWP_ASYNCWINDOWPOS = &H4000& '    Wenn der aufrufende Thread und der Thread, der das Fenster besitzt, an verschiedene Eingabewarteschlangen angefügt sind, sendet das System die Anforderung an den Thread, der das Fenster besitzt. Dadurch wird verhindert, dass der aufrufende Thread seine Ausführung blockiert, während andere Threads die Anforderung verarbeiten.
    SWP_NOSENDCHANGING = &H400&  '    Verhindert, dass das Fenster die WM_WINDOWPOSCHANGING-Nachricht empfängt.
    SWP_DEFERERASE = &H2000&     '    Verhindert die Generierung der WM_SYNCPAINT Nachricht.
End Enum

Public Enum EWndShow
    SW_HIDE = &H0&              ' Blendet das Fenster aus und aktiviert ein anderes Fenster.
    SW_SHOWNORMAL = &H1&
    SW_NORMAL = &H1&            ' Aktiviert und zeigt ein Fenster an. Wenn das Fenster minimiert, maximiert oder angeordnet ist, wird es vom System auf seine ursprüngliche Größe und Position wiederhergestellt. Eine Anwendung sollte dieses Flag angeben, wenn das Fenster zum ersten Mal angezeigt wird.
    SW_SHOWMINIMIZED = &H2&     ' Aktiviert das Fenster und zeigt es als minimiertes Fenster an.
    SW_SHOWMAXIMIZED = &H3&
    SW_MAXIMIZE = &H3&          ' Aktiviert das Fenster und zeigt es als maximiertes Fenster an.
    SW_SHOWNOACTIVATE = &H4&    ' Zeigt ein Fenster in seiner neuesten Größe und Position an. Dieser Wert ähnelt SW_SHOWNORMAL, mit der Ausnahme, dass das Fenster nicht aktiviert ist.
    SW_SHOW = &H5&              ' Aktiviert das Fenster und zeigt es in seiner aktuellen Größe und Position an.
    SW_MINIMIZE = &H6&          ' Minimiert das angegebene Fenster und aktiviert das nächste Fenster der obersten Ebene in der Reihenfolge Z.
    SW_SHOWMINNOACTIVE = &H7&   ' Zeigt das Fenster als minimiertes Fenster an. Dieser Wert ähnelt SW_SHOWMINIMIZED, außer dass das Fenster nicht aktiviert ist.
    SW_SHOWNA = &H8&            ' Zeigt das Fenster in seiner aktuellen Größe und Position an. Dieser Wert ähnelt SW_SHOW, mit der Ausnahme, dass das Fenster nicht aktiviert ist.
    SW_RESTORE = &H9&           ' Aktiviert das Fenster und zeigt es an. Wenn das Fenster minimiert, maximiert oder angeordnet ist, wird es vom System auf seine ursprüngliche Größe und Position wiederhergestellt. Eine Anwendung sollte dieses Flag beim Wiederherstellen eines minimierten Fensters angeben.
    SW_SHOWDEFAULT = &H10&      ' Legt den Showstatus basierend auf dem SW_ Wert fest, der in der STARTUPINFO-Struktur angegeben ist, die von dem Programm, das die Anwendung gestartet hat, an die CreateProcess-Funktion übergeben wird.
    SW_FORCEMINIMIZE = &H11&    ' Minimiert ein Fenster, auch wenn der Thread, der das Fenster besitzt, nicht reagiert. Dieses Flag sollte nur verwendet werden, wenn Fenster aus einem anderen Thread minimiert werden.
End Enum


Private Const IDC_ARROW              As Long = 32512&
Private Const IDC_IBEAM              As Long = 32513&
Private Const IDC_WAIT               As Long = 32514&
Private Const IDC_CROSS              As Long = 32515&
Private Const IDC_UPARROW            As Long = 32516&

Private Const IDC_SIZE               As Long = 32640&
Private Const IDC_ICON               As Long = 32641&
Private Const IDC_SIZENWSE           As Long = 32642&
Private Const IDC_SIZENESW           As Long = 32643&
Private Const IDC_SIZEWE             As Long = 32644&
Private Const IDC_SIZENS             As Long = 32645&
Private Const IDC_SIZEALL            As Long = 32646&
'Private Const IDC_                   As Long = 32647&
Private Const IDC_NO                 As Long = 32648&
Private Const IDC_HAND               As Long = 32649&
Private Const IDC_APPSTARTING        As Long = 32650&
Private Const IDC_HELP               As Long = 32651&

'https://learn.microsoft.com/de-de/windows/win32/api/winuser/ns-winuser-createstructw
'typedef struct tagCREATESTRUCTW {
'  LPVOID    lpCreateParams;
'  HINSTANCE hInstance;
'  HMENU     hMenu;
'  HWND      hwndParent;
'  int       cy;
'  int       cx;
'  int       y;
'  int       x;
'  LONG      style;
'  LPCWSTR   lpszName;
'  LPCWSTR   lpszClass;
'  DWORD     dwExStyle;
'} CREATESTRUCTW, *LPCREATESTRUCTW;
'contains every argument CreateWindowExW needs in the order needed for the stack
Private Type CREATESTRUCTW
    lpCreateParams As LongPtr
    hInstance      As LongPtr
    hMenu          As LongPtr
    hwndParent     As LongPtr
    cy             As Long
    cx             As Long
    Y              As Long
    X              As Long
    Style          As EWndStyle
    lpszName       As LongPtr
    lpszClass      As LongPtr
    dwExStyle      As EWndStyleEx
End Type


'https://www.codeproject.com/Articles/7969/How-to-make-a-Message-Only-Window
Private Const HWND_MESSAGE As Long = -3&
'Private Const WndMsgClassName As String = "Message"

'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-createwindoww
'HWND CreateWindowW(
'  [in, optional]  lpClassName,
'  [in, optional]  lpWindowName,
'  [in]            dwStyle,
'  [in]            x,
'  [in]            y,
'  [in]            nWidth,
'  [in]            nHeight,
'  [in, optional]  hWndParent,
'  [in, optional]  hMenu,
'  [in, optional]  hInstance,
'  [in, optional]  lpParam
');
#If VBA7 Then
    
    Private Declare PtrSafe Function RegisterClassExW Lib "user32" (pcWndClassEx As WNDCLASSEXW) As Long

    Private Declare PtrSafe Function UnregisterClassW Lib "user32" (ByVal lpClassName As LongPtr, ByVal hInstance As LongPtr) As Long

    Private Declare PtrSafe Function CreateWindowExW Lib "user32" ( _
        ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, _
        ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As LongPtr, _
        ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
    
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    
    Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal EWndShow As Long) As Long
    
    Private Declare PtrSafe Function SetWindowTextW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr) As Long
    
    Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    
    Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

    Private Declare PtrSafe Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As LongPtr) As LongPtr

    Private Declare PtrSafe Function LoadCursorW Lib "user32" (ByVal hInstance As LongPtr, ByVal lpCursorName As LongPtr) As LongPtr
    
#Else

    'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nc-winuser-wndproc
    'LRESULT Wndproc( HWND unnamedParam1, UINT unnamedParam2, WPARAM unnamedParam3, LPARAM unnamedParam4 )
    'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-registerclassexw
    'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-registerclassexw
    'ATOM RegisterClassExA( [in] const WNDCLASSEXA *unnamedParam1 );
    Private Declare Function RegisterClassExW Lib "user32" (pcWndClassEx As WNDCLASSEXW) As Long
    'https://learn.microsoft.com/de-de/windows/win32/winmsg/using-window-classes
    
    'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-unregisterclassw
    'BOOL UnregisterClassW( [in] LPCWSTR lpClassName, [in, optional] HINSTANCE hInstance);
    Private Declare Function UnregisterClassW Lib "user32" (ByVal lpClassName As LongPtr, ByVal hInstance As LongPtr) As Long
    
    'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-createwindowexw
    'HWND CreateWindowExW(
    '  [in]           DWORD     dwExStyle,
    '  [in, optional] LPCWSTR   lpClassName,
    '  [in, optional] LPCWSTR   lpWindowName,
    '  [in]           DWORD     dwStyle,
    '  [in]           int       X,
    '  [in]           int       Y,
    '  [in]           int       nWidth,
    '  [in]           int       nHeight,
    '  [in, optional] HWND      hWndParent,
    '  [in, optional] HMENU     hMenu,
    '  [in, optional] HINSTANCE hInstance,
    '  [in, optional] LPVOID    lpParam
    ');
    Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, _
        ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As LongPtr, _
        ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
    
    'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-setwindowpos
    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As EWndShow) As Long
    
    'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-movewindow
    'BOOL MoveWindow( [in] HWND hWnd, [in] int  X, [in] int  Y, [in] int  nWidth, [in] int  nHeight, [in] BOOL bRepaint );
    Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    
    'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-showwindow
    'BOOL ShowWindow( [in] HWND hWnd, [in] int  nCmdShow );
    Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As EWndShow) As Long
    
    'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-setwindowtextw
    'BOOL SetWindowTextW( [in] HWND hWnd, [in, optional] LPCWSTR lpString);
    Private Declare Function SetWindowTextW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr) As Long
    
    'BOOL DestroyWindow( [in] HWND hWnd );
    Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    
    
    'https://learn.microsoft.com/de-de/windows/win32/api/errhandlingapi/nf-errhandlingapi-getlasterror
    Private Declare Function GetLastError Lib "kernel32" () As Long
    
    'https://learn.microsoft.com/de-de/windows/win32/api/libloaderapi/nf-libloaderapi-getmodulehandlew
    'HMODULE GetModuleHandleW( [in, optional] LPCWSTR lpModuleName );
    Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As LongPtr) As LongPtr
    
    'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-loadcursorw
    'HCURSOR LoadCursorW( [in, optional] HINSTANCE hInstance, [in] LPCWSTR lpCursorName);
    Private Declare Function LoadCursorW Lib "user32" (ByVal hInstance As LongPtr, ByVal lpCursorName As LongPtr) As LongPtr
    
#End If

Private m_Name         As String
Private m_Caption      As String
Private Const ClsName  As String = "ThunderVB64uWindow"
Private m_ClsName      As String
Private m_WndCls       As WNDCLASSEXW
Private m_ClsAtom      As LongPtr
Private m_Window       As CREATESTRUCTW
Private m_hWnd         As LongPtr
Private m_hDC          As LongPtr
Private m_Tag          As Variant
Private m_ScaleMode    As ScaleModeConstants
Private m_WindowState  As EWndShow
Private m_Controls     As Collection
Private m_ActControl   As Object
Private m_Font         As StdFont
Private m_Icon         As StdPicture 'IPictureDisp
Private m_MouseIcon    As StdPicture 'IPictureDisp
Private m_MousePointer As Long ' = IDC_

Public Event Activate()
Public Event Click()
Public Event DblClick()
Public Event Deactivate()
Public Event DragDrop(Source As Control, X As Single, Y As Single)
Public Event DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Public Event GotFocus()
Public Event Initialize()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event LinkClose()
Public Event LinkError(LinkErr As Integer)
Public Event LinkExecute(CmdStr As String, Cancel As Integer)
Public Event LinkOpen(Cancel As Integer)
Public Event Load()
Public Event LostFocus()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Public Event Paint()
Public Event QueryUnload(Cancel As Integer, UnloadMode As Integer)
Public Event Resize()
Public Event Terminate()
Public Event Unload(Cancel As Integer)

Private Function GethInstance() As LongPtr
    GethInstance = GetModuleHandleW(StrPtr(vbNullString))
End Function

Private Sub Class_Initialize()
    m_WindowState = SW_NORMAL
    m_Name = "Window"
    m_ClsName = ClsName
    m_Caption = "Form"
    Dim cap As String: cap = m_Caption & vbNullChar
    With m_Window
        .lpszName = StrPtr(cap)
        .hInstance = GethInstance 'App.hInstance
        .Style = EWndStyle.VBFormStyle_Sizable
        .dwExStyle = EWndStyle.VBFormStyle_Sizable
        '         WS_MAXIMIZEBOX Or WS_TABSTOP Or WS_GROUP Or WS_MINIMIZEBOX Or _
        '         WS_SIZEBOX Or WS_THICKFRAME Or WS_SYSMENU Or WS_DLGFRAME Or WS_BORDER Or _
        '         WS_CAPTION Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
        .dwExStyle = WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW
        'default position and size
        .X = 33
        .Y = 26
        .cx = 320 'px VB-default
        .cy = 240 'px VB-default
    End With
    m_MousePointer = IDC_ARROW
End Sub

Private Sub Class_Terminate()
    RaiseEvent Terminate
    Destroy
End Sub

Private Sub Destroy()
    Dim hr As Long
    
    hr = DestroyWindow(m_hWnd)
    If hr Then
        MWindow.Windows_Delete m_hWnd
    Else
        'ErrHandler "Class_Terminate", "Could not destroy the Window " & m_hWnd
    End If
    
    hr = UnregisterClassW(m_ClsAtom, GethInstance)
    If hr Then
        MWindow.Atoms_Delete m_ClsName
    Else
        'ErrHandler "Class_Terminate", "Could not unregister the class atom " & m_ClsAtom
    End If
End Sub

Friend Sub New_(ByVal Name As String, Optional ByVal Caption As String, Optional ByVal Style As EWndStyle = EWndStyle.VBFormStyle_Sizable, Optional ByVal StyleEx As EWndStyleEx = EWndStyleEx.VBFormStyleEx_Sizable, Optional ByVal ClsStyle As EClsStyle = EClsStyle.VBFormClassStyleDefault)
    m_Name = Name
    m_ClsName = ClsName & Name
    m_Caption = IIf(Len(Caption), Caption, m_Name)
    Dim cap As String: cap = m_Caption & vbNullChar
    With m_Window
        .lpszName = StrPtr(cap)
        .Style = Style
        .dwExStyle = StyleEx
    End With
    With m_WndCls
        .Style = ClsStyle
    End With
End Sub

Private Function Register() As Boolean
    'preparing the wnd-class-struct
    Dim cln As String: cln = m_ClsName & vbNullChar
    Debug.Print m_ClsName
    With m_WndCls
        .cbSize = LenB(m_WndCls) '48?
        .lpszClassName = StrPtr(cln)
        .lpfnWndProc = FncPtr(AddressOf MWindow.WndProc)
        .hInstance = GethInstance 'App.hInstance
        '.cbClsExtra
        '.cbWndExtra
        '.hbrBackground
        If m_MouseIcon Is Nothing Then
            .hCursor = LoadCursorW(0, m_MousePointer)
        Else
            If m_MouseIcon.Handle = 0 Then
                .hCursor = LoadCursorW(0, m_MousePointer)
            End If
        End If
        '.hIcon
        '.hIconSm
        '.lpszMenuName
    End With
    
    'first check if already registered
    Dim ClsName As String: ClsName = m_ClsName
    If MWindow.Atoms_Contains(ClsName) Then
        'OK already registered in this process
        'so we just assign the atom an leave
        m_ClsAtom = MWindow.Atoms_Item(ClsName)
        Register = CBool(m_ClsAtom)
        m_WndCls.lpszClassName = m_ClsAtom
        Exit Function
    End If
    'class not registered until now
    'so first try to register with the original name
    ClsName = ClsName & vbNullChar
    
    m_WndCls.lpszClassName = StrPtr(ClsName)
    m_ClsAtom = RegisterClassExW(m_WndCls)
    Register = CBool(m_ClsAtom)
    If Register Then
        MWindow.Atoms_Add m_ClsName, m_ClsAtom
        Exit Function
    End If
    'let's try again with an altered name
    Randomize Timer
    Dim sr As String: sr = Rnd * 2147483647
    ClsName = m_ClsName & sr & vbNullChar
    
    m_WndCls.lpszClassName = StrPtr(ClsName)
    m_ClsAtom = RegisterClassExW(m_WndCls)
    Register = CBool(m_ClsAtom)
    If Register Then
        MWindow.Atoms_Add m_ClsName, m_ClsAtom
        Exit Function
    End If
    'we did not succeed
    ErrHandler "Register", "m_ClsAtom=0"
End Function

Public Sub Load()
'Try: On Error GoTo Catch
    If Not Register Then
        'MsgBox "Could not register class: " & m_ClsName
        'Exit Sub
    End If
    With m_Window
        .lpszClass = m_ClsAtom
        m_hWnd = CreateWindowExW(.dwExStyle, .lpszClass, .lpszName, .Style, .X, .Y, .cx, .cy, .hwndParent, .hMenu, .hInstance, ByVal .lpCreateParams)
        'Debug.Print "Windows_Add"
        MWindow.Windows_Add m_hWnd, Me
        'MWindow.WindowAdded = True
        If m_hWnd = 0 Then
            'GoTo Catch
            ErrHandler "Load", "hWnd = 0"
        End If
    End With
    RaiseEvent Load
    RaiseEvent Resize
    RaiseEvent Activate
    RaiseEvent GotFocus
'Paint()
'    Exit Sub
'Catch:
End Sub

Public Sub Unload()
    Dim Cancel  As Integer
    RaiseEvent Unload(Cancel)
    If Cancel Then Exit Sub
    Destroy
End Sub

'Public Sub Show(Optional Modal As FormShowConstants, Optional OwnerForm As Object): End Sub
Public Sub Show(Optional Modal As FormShowConstants, Optional OwnerForm As Object)
    If m_hWnd = 0 Then Me.Load
    If m_hWnd = 0 Then Exit Sub
    'Dim hr As Long: hr = ShowWindow(m_hWnd, EWndShow.SW_SHOWNORMAL)
    Dim hr As Long: hr = ShowWindow(m_hWnd, m_WindowState)
    
End Sub

' v ############################## v '    OM TODO:    ' v ############################## v '
'Drawing functions:
Public Sub Cls() 'clear screen
    '
End Sub
Public Sub CCircle() 'draws a Circle
    '
End Sub
Public Sub Line() 'draws a line
    '
End Sub
Public Function Point(X As Single, Y As Single) As Long 'moves to point
    '
End Function
Public Sub PPSet(Step As Integer, X As Single, Y As Single, Color As Long)
    '
End Sub

Public Property Get ActiveControl() As Control
    'Gibt das Steuerelement mit dem Fokus zurück.

End Property

'0 - 2D
'1 - 3D
Public Property Get Appearance() As Integer
    'Gibt an oder legt fest, ob ein Objekt zur Laufzeit mit 3D-Effekten ausgegeben wird oder nicht.

End Property
Public Property Let Appearance(ByVal Value As Integer)
    '
End Property

Public Property Get AutoRedraw() As Boolean
    ' Gibt die Ausgabe einer Grafikmethode an eine beständige Bitmap zurück oder legt diese fest.
End Property
Public Property Let AutoRedraw(ByVal Value As Boolean)
    '
End Property

Public Property Get BackColor() As Long
    ' Gibt die Hintergrundfarbe zurück, die verwendet wird, um Text und Grafik in einem Objekt anzuzeigen, oder legt diese fest.

End Property
Public Property Let BackColor(ByVal Value As Long)
    '
End Property

'Enum FormBorderStyleConstants
'    vbBSNone = 0
'    vbFixedSingle = 1
'    vbSizable = 2
'    vbFixedDialog = 3
'    vbFixedToolWindow = 4
'    vbSizableToolWindow = 5
'End Enum
Public Property Get BorderStyle() As FormBorderStyleConstants
    'WS_SIZEBOX    = &H40000         '  Das Fenster verfügt über einen Rahmen zur Größenanpassung. Identisch mit dem WS_THICKFRAME Stil.
    'WS_THICKFRAME = &H40000         '  Das Fenster verfügt über einen Rahmen zur Größenanpassung. Identisch mit dem WS_SIZEBOX Stil.
    'WS_DLGFRAME   = &H400000        '  Das Fenster verfügt über einen Rahmen eines Stils, der in der Regel mit Dialogfeldern verwendet wird.
                                     '  Ein Fenster mit dieser Formatvorlage kann keine Titelleiste aufweisen.
    'WS_BORDER     = &H800000        '  Das Fenster verfügt über einen dünnen Rahmen
    Dim bs As FormBorderStyleConstants
    Dim ws As EWndStyle: ws = m_Window.Style
    Select Case ws
    Case EWndStyle.VBFormStyle_BSNone: bs = vbBSNone
    Case EWndStyle.VBFormStyle_FixedSingle: bs = vbFixedSingle
    Case EWndStyle.VBFormStyle_Sizable: bs = vbSizable
    Case EWndStyle.VBFormStyle_FixedDialog: bs = vbFixedDialog
    Case EWndStyle.VBFormStyle_FixedToolWindow: bs = vbFixedToolWindow
    Case EWndStyle.VBFormStyle_SizableToolWindow: bs = vbSizableToolWindow
    End Select
    BorderStyle = bs
End Property
Public Property Let BorderStyle(ByVal Value As FormBorderStyleConstants)
    With m_Window
        Select Case Value
        Case FormBorderStyleConstants.vbBSNone:            .Style = EWndStyle.VBFormStyle_BSNone:            .dwExStyle = EWndStyleEx.VBFormStyleEx_BSNone
        Case FormBorderStyleConstants.vbFixedSingle:       .Style = EWndStyle.VBFormStyle_FixedSingle:       .dwExStyle = EWndStyleEx.VBFormStyleEx_FixedSingle
        Case FormBorderStyleConstants.vbSizable:           .Style = EWndStyle.VBFormStyle_Sizable:           .dwExStyle = EWndStyleEx.VBFormStyleEx_Sizable
        Case FormBorderStyleConstants.vbFixedDialog:       .Style = EWndStyle.VBFormStyle_FixedDialog:       .dwExStyle = EWndStyleEx.VBFormStyleEx_FixedDialog
        Case FormBorderStyleConstants.vbFixedToolWindow:   .Style = EWndStyle.VBFormStyle_FixedToolWindow:   .dwExStyle = EWndStyleEx.VBFormStyleEx_FixedToolWindow
        Case FormBorderStyleConstants.vbSizableToolWindow: .Style = EWndStyle.VBFormStyle_SizableToolWindow: .dwExStyle = EWndStyleEx.VBFormStyleEx_SizableToolWindow
        End Select
    End With
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property
Public Property Let Caption(ByVal Value As String)
    m_Caption = Value
    Dim hr As Long: hr = SetWindowTextW(m_hWnd, StrPtr(m_Caption & vbNullChar))
End Property

Public Property Get Text() As String
    Text = Me.Caption
End Property
Public Property Let Text(ByVal Value As String)
    Me.Caption = Value
End Property

Public Property Get ClassName() As String
    ClassName = m_ClsName
End Property
    
Public Property Get ClipControls() As Boolean
    '
End Property

Public Property Get Controls() As Object
    Set Controls = m_Controls
End Property

Public Property Get Count() As Integer
    'schreibgeschützt
    'Gibt die Anzahl der Objekte in einer Auflistung zurück.
    Count = m_Controls.Count
End Property

Public Property Get CurrentX() As Single:                      End Property
Public Property Get CurrentY() As Single:                      End Property

Public Property Get DrawMode() As DrawModeConstants
    '
End Property

Public Property Get DrawStyle() As DrawStyleConstants:         End Property
Public Property Get DrawWidth() As Integer:                    End Property
Public Property Get Enabled() As Boolean:                      End Property
Public Property Get FillColor() As Long:                       End Property
Public Property Get FillStyle() As FillStyleConstants:         End Property

' v ############################## v '    Font    ' v ############################## v '
Public Property Get Font() As StdFont
    'Gibt ein Font-Objekt zurück.
    Set Font = m_Font
End Property
Public Property Set Font(ByVal Value As StdFont)
    'm_Window.h
    Set m_Font = Value
End Property

Public Property Get FontBold() As Boolean
    'Gibt Schriftstile für Fettschrift zurück oder legt diese fest.
    FontBold = m_Font.Bold
End Property
Public Property Let FontBold(ByVal Value As Boolean)
    m_Font.Bold = Value
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = m_Font.Italic
End Property
Public Property Let FontItalic(ByVal Value As Boolean)
    m_Font.Italic = Value
End Property

Public Property Get FontName() As String
    FontName = m_Font.Name
End Property
Public Property Let FontName(ByVal Value As String)
    m_Font.Name = Value
End Property

Public Property Get FontSize() As Single
    FontSize = m_Font.Size
End Property
Public Property Let FontSize(ByVal Value As Single)
    m_Font.Size = Value
End Property

Public Property Get FontStrikethru() As Boolean
    FontStrikethru = m_Font.Strikethrough
End Property
Public Property Let FontStrikethru(ByVal Value As Boolean)
    m_Font.Strikethrough = Value
End Property

Public Property Get FontTransparent() As Boolean
    '
End Property
Public Property Let FontTransparent(ByVal Value As Boolean)
    'm_font.tr
End Property

Public Property Get FontUnderline() As Boolean
    FontUnderline = m_Font.Underline
End Property
Public Property Let FontUnderline(ByVal Value As Boolean)
    m_Font.Underline = Value
End Property
' ^ ############################## ^ '    Font    ' ^ ############################## ^ '


Public Property Get ForeColor() As Long:                       End Property

Public Property Get HasDC() As Boolean
    '
    
End Property
Public Property Get hDC() As LongPtr
    '
    
End Property
'Public Property Get Height() As Single : End Property
Public Property Get HelpContextID() As Long: End Property

Public Sub Hide()
    'Mitglied von VB.Form
    'Blendet ein MDIForm- oder Form-Objekt aus, ohne es zu entladen.
    If m_hWnd Then
        Dim hr As Long: hr = ShowWindow(m_hWnd, EWndShow.SW_HIDE)
    End If
End Sub

'Public Property Get hWnd() As LongPtr: End Property
Public Property Get Icon() As StdPicture ' IPictureDisp
    Set Icon = m_Icon
End Property
Public Property Set Icon(ByVal Value As StdPicture) ' IPictureDisp)
    Set m_Icon = Value
    'Debug.Print m_Icon.Handle
    m_WndCls.hIcon = m_Icon.Handle
    m_WndCls.hIconSm = m_Icon.Handle
End Property

Public Property Get Image() As IPictureDisp: End Property
Public Property Get KeyPreview() As Boolean: End Property
'Public Property Get Left() As Single : End Property
Public Property Get LinkMode() As Integer: End Property
Public Property Get LinkTopic() As String: End Property

'Public Enum EWndStyle
'    WS_TILED = &H0&                 '  Das Fenster ist ein überlappende Fenster. Ein überlappendes Fenster hat eine Titelleiste und einen Rahmen. Identisch mit dem WS_OVERLAPPED Stil.
'    WS_OVERLAPPED = &H0&            '  Das Fenster ist ein überlappende Fenster. Ein überlappendes Fenster hat eine Titelleiste und einen Rahmen. Identisch mit dem WS_TILED-Stil .
'
'    WS_MAXIMIZEBOX = &H10000        '  Das Fenster verfügt über eine Schaltfläche zum Maximieren. Kann nicht mit dem WS_EX_CONTEXTHELP-Stil kombiniert werden. Die WS_SYSMENU Formatvorlage muss ebenfalls angegeben werden.
'    WS_TABSTOP = &H10000            '  Das Fenster ist ein Steuerelement, das den Tastaturfokus erhalten kann, wenn der Benutzer die TAB-TASTE drückt. Durch Drücken der TAB-TASTE wird der Tastaturfokus auf das nächste Steuerelement mit der WS_TABSTOP-Formatvorlage geändert.Sie können diesen Stil aktivieren und deaktivieren, um die Navigation im Dialogfeld zu ändern. Um diesen Stil zu ändern, nachdem ein Fenster erstellt wurde, verwenden Sie die SetWindowLong-Funktion . Damit vom Benutzer erstellte Fenster und moduslose Dialoge mit Tabstopps funktionieren, ändern Sie die Nachrichtenschleife so, dass die IsDialogMessage-Funktion aufgerufen wird.
'    WS_GROUP = &H20000              '  Das Fenster ist das erste Steuerelement einer Gruppe von Steuerelementen. Die Gruppe besteht aus diesem ersten Steuerelement und allen danach definierten Steuerelementen bis zum nächsten Steuerelement mit dem WS_GROUP Stil. Das erste Steuerelement in jeder Gruppe hat in der Regel den WS_TABSTOP Stil, sodass der Benutzer von Gruppe zu Gruppe wechseln kann. Der Benutzer kann anschließend den Tastaturfokus von einem Steuerelement in der Gruppe auf das nächste Steuerelement in der Gruppe ändern, indem er die Richtungstasten verwendet.Sie können diesen Stil aktivieren und deaktivieren, um die Navigation im Dialogfeld zu ändern. Um diesen Stil zu ändern, nachdem ein Fenster erstellt wurde, verwenden Sie die SetWindowLong-Funktion .
'    WS_MINIMIZEBOX = &H20000        '  Das Fenster verfügt über eine Schaltfläche zum Minimieren. Kann nicht mit dem WS_EX_CONTEXTHELP-Stil kombiniert werden. Die WS_SYSMENU Formatvorlage muss ebenfalls angegeben werden.
'    WS_SIZEBOX = &H40000            '  Das Fenster verfügt über einen Rahmen zur Größenanpassung. Identisch mit dem WS_THICKFRAME Stil.
'    WS_THICKFRAME = &H40000         '  Das Fenster verfügt über einen Rahmen zur Größenanpassung. Identisch mit dem WS_SIZEBOX Stil.
'    WS_SYSMENU = &H80000            '  Das Fenster verfügt über ein Fenstermenü auf der Titelleiste. Die WS_CAPTION Formatvorlage muss ebenfalls angegeben werden.
'
'    WS_HSCROLL = &H100000           '  Das Fenster verfügt über eine horizontale Bildlaufleiste.
'    WS_VSCROLL = &H200000           '  Das Fenster verfügt über eine vertikale Bildlaufleiste.
'    WS_DLGFRAME = &H400000          '  Das Fenster verfügt über einen Rahmen eines Stils, der in der Regel mit Dialogfeldern verwendet wird. Ein Fenster mit dieser Formatvorlage kann keine Titelleiste aufweisen.
'    WS_BORDER = &H800000            '  Das Fenster verfügt über einen dünnen Rahmen
'    WS_CAPTION = &HC00000           '  Das Fenster verfügt über eine Titelleiste (einschließlich des WS_BORDER Stils).
'
'    WS_MAXIMIZE = &H1000000         '  Das Fenster wird anfänglich maximiert.
'    WS_CLIPCHILDREN = &H2000000     '  Schließt den Bereich aus, der von untergeordneten Fenstern belegt wird, wenn das Zeichnen innerhalb des übergeordneten Fensters erfolgt. Diese Formatvorlage wird beim Erstellen des übergeordneten Fensters verwendet.
'    WS_CLIPSIBLINGS = &H4000000     '  Schneidet untergeordnete Fenster relativ zueinander ab; Das heißt, wenn ein bestimmtes untergeordnetes Fenster eine WM_PAINT-Meldung empfängt, wird vom WS_CLIPSIBLINGS-Format alle anderen überlappenden untergeordneten Fenster aus dem Bereich des zu aktualisierenden untergeordneten Fensters heraus geklammert. Wenn WS_CLIPSIBLINGS nicht angegeben ist und sich untergeordnete Fenster überschneiden, ist es beim Zeichnen innerhalb des Clientbereichs eines untergeordneten Fensters möglich, innerhalb des Clientbereichs eines benachbarten untergeordneten Fensters zu zeichnen.
'    WS_DISABLED = &H8000000         '  Das Fenster ist zunächst deaktiviert. Ein deaktiviertes Fenster kann keine Eingaben vom Benutzer empfangen. Um dies zu ändern, nachdem ein Fenster erstellt wurde, verwenden Sie die Funktion EnableWindow .'
'
'    WS_VISIBLE = &H10000000         '  Das Fenster ist zunächst sichtbar. Diese Formatvorlage kann mithilfe der ShowWindow - oder SetWindowPos-Funktion aktiviert und deaktiviert werden.
'    WS_ICONIC = &H20000000          '  Das Fenster wird zunächst minimiert. Identisch mit dem WS_MINIMIZE Stil.
'    WS_MINIMIZE = &H20000000        '  Das Fenster wird zunächst minimiert. Identisch mit dem WS_ICONIC Stil.
'    WS_CHILD = &H40000000           '  Das Fenster ist ein untergeordnetes Fenster. Ein Fenster mit dieser Formatvorlage kann keine Menüleiste aufweisen. Diese Formatvorlage kann nicht mit der WS_POPUP-Formatvorlage verwendet werden.
'    WS_CHILDWINDOW = &H40000000     '  Identisch mit dem WS_CHILD Stil.
'    WS_POPUP = &H80000000           '  Das Fenster ist ein Popupfenster. Dieser Stil kann nicht mit dem WS_CHILD-Stil verwendet werden.
'    WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)       'Das Fenster ist ein Popupfenster. Die Formatvorlagen WS_CAPTION und WS_POPUPWINDOW müssen kombiniert werden, um das Fenstermenü sichtbar zu machen.
'    WS_TILEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)           ' Das Fenster ist ein überlappende Fenster. Identisch mit dem WS_OVERLAPPEDWINDOW Stil.
'    WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX) 'Das Fenster ist ein überlappende Fenster. Identisch mit dem WS_TILEDWINDOW Stil.
'    VBFormStyle = WS_MAXIMIZEBOX Or WS_TABSTOP Or WS_GROUP Or WS_MINIMIZEBOX Or WS_SIZEBOX Or WS_THICKFRAME Or WS_SYSMENU Or WS_DLGFRAME Or WS_BORDER Or WS_CAPTION Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
'End Enum
Public Property Get ControlBox() As Boolean
    ControlBox = m_Window.Style And WS_SYSMENU
End Property
Public Property Let ControlBox(ByVal Value As Boolean)
    m_Window.Style = m_Window.Style Or WS_SYSMENU
    If Not Value Then m_Window.Style = m_Window.Style Xor WS_SYSMENU
End Property

Public Property Get MinButton() As Boolean
    MinButton = m_Window.Style And WS_MINIMIZEBOX
End Property
Public Property Let MinButton(ByVal Value As Boolean)
    m_Window.Style = m_Window.Style Or WS_MINIMIZEBOX
    If Not Value Then m_Window.Style = m_Window.Style Xor WS_MINIMIZEBOX
End Property

Public Property Get MaxButton() As Boolean
    MaxButton = m_Window.Style And WS_MAXIMIZEBOX
End Property
Public Property Let MaxButton(ByVal Value As Boolean)
    m_Window.Style = m_Window.Style Or WS_MAXIMIZEBOX
    If Not Value Then m_Window.Style = m_Window.Style Xor WS_MAXIMIZEBOX
End Property

Public Property Get Visible() As Boolean
    'Mitglied von VB.Form
    'Gibt einen Wert zurück, der bestimmt, ob ein Objekt sichtbar oder versteckt ist, oder legt diesen fest.
    Visible = m_Window.Style And WS_VISIBLE
End Property
Public Property Let Visible(ByVal Value As Boolean)
    m_Window.Style = m_Window.Style Or WS_VISIBLE
    If Not Value Then m_Window.Style = m_Window.Style Xor WS_VISIBLE
    If m_hWnd Then
        Dim sw As EWndShow: sw = IIf(Value, EWndShow.SW_SHOW, EWndShow.SW_HIDE)
        Dim hr As Long:     hr = ShowWindow(m_hWnd, sw)
    End If
End Property

Public Property Get ScrollBars() As ScrollBarConstants
    ScrollBars = ScrollBarConstants.vbSBNone
    If m_Window.Style And WS_HSCROLL Then ScrollBars = ScrollBars And ScrollBarConstants.vbHorizontal
    If m_Window.Style And WS_VSCROLL Then ScrollBars = ScrollBars And ScrollBarConstants.vbVertical
End Property
Public Property Let ScrollBars(ByVal Value As ScrollBarConstants)
    m_Window.Style = m_Window.Style Or WS_HSCROLL
    m_Window.Style = m_Window.Style Or WS_VSCROLL
    If Value = ScrollBarConstants.vbHorizontal Then
        m_Window.Style = m_Window.Style Xor WS_VSCROLL
    End If
    If Value = ScrollBarConstants.vbVertical Then
        m_Window.Style = m_Window.Style Xor WS_HSCROLL
    End If
    If Value = ScrollBarConstants.vbSBNone Then
        m_Window.Style = m_Window.Style Xor WS_HSCROLL
        m_Window.Style = m_Window.Style Xor WS_VSCROLL
    End If
End Property



Public Property Get MDIChild() As Boolean
    MDIChild = m_Window.dwExStyle And WS_EX_MDICHILD
End Property
Public Property Let MDIChild(ByVal Value As Boolean)
    m_Window.dwExStyle = m_Window.dwExStyle Or WS_EX_MDICHILD
    If Not Value Then m_Window.dwExStyle = m_Window.dwExStyle Xor WS_EX_MDICHILD
End Property

Public Property Get MouseIcon() As StdPicture 'IPictureDisp
    Set MouseIcon = m_MouseIcon
End Property
Public Property Set MouseIcon(ByVal Value As StdPicture) 'IPictureDisp)
    Set m_MouseIcon = Value
    If m_MouseIcon Is Nothing Then Exit Property
    m_WndCls.hCursor = m_MouseIcon.Handle
End Property

'Public Const IDC_ARROW              As Long = 32512&
'Public Const IDC_IBEAM              As Long = 32513&
'Public Const IDC_WAIT               As Long = 32514&
'Public Const IDC_CROSS              As Long = 32515&
'Public Const IDC_UPARROW            As Long = 32516&
'Public Const IDC_SIZE               As Long = 32640&
'Public Const IDC_ICON               As Long = 32641&
'Public Const IDC_SIZENWSE           As Long = 32642&
'Public Const IDC_SIZENESW           As Long = 32643&
'Public Const IDC_SIZEWE             As Long = 32644&
'Public Const IDC_SIZENS             As Long = 32645&
'Public Const IDC_SIZEALL            As Long = 32646&
''Public Const IDC_                   As Long = 32647&
'Public Const IDC_NO                 As Long = 32648&
'Public Const IDC_HAND               As Long = 32649&
'Public Const IDC_APPSTARTING        As Long = 32650&
'Public Const IDC_HELP               As Long = 32651&
'Enum MousePointerConstants
'    vbDefault = 0
'    vbArrow = 1
'    vbCrosshair = 2
'    vbIbeam = 3
'    vbIconPointer = 4
'    vbSizePointer = 5
'    vbSizeNESW = 6
'    vbSizeNS = 7
'    vbSizeNWSE = 8
'    vbSizeWE = 9
'    vbUpArrow = 10
'    vbHourglass = 11
'    vbNoDrop = 12
'    vbArrowHourglass = 13
'    vbArrowQuestion = 14
'    vbSizeAll = 15
'    vbCustom = 99 ' (&H63)
'End Enum

Public Property Get MousePointer() As Integer 'MousePointerConstants
    Dim idmp As Long: idmp = m_MousePointer
    Dim mp   As MousePointerConstants
    Select Case idmp
    Case 0:               mp = MousePointerConstants.vbArrow
    Case IDC_ARROW:       mp = MousePointerConstants.vbArrow
    Case IDC_IBEAM:       mp = MousePointerConstants.vbIbeam
    Case IDC_CROSS:       mp = MousePointerConstants.vbCrosshair
    Case IDC_UPARROW:     mp = MousePointerConstants.vbUpArrow
    Case IDC_SIZE:        mp = MousePointerConstants.vbSizeAll
    Case IDC_ICON:        mp = MousePointerConstants.vbIconPointer
    Case IDC_SIZENWSE:    mp = MousePointerConstants.vbSizeNWSE
    Case IDC_SIZENESW:    mp = MousePointerConstants.vbSizeNESW
    Case IDC_SIZEWE:      mp = MousePointerConstants.vbSizeWE
    Case IDC_SIZENS:      mp = MousePointerConstants.vbSizeNS
    Case IDC_SIZEALL:     mp = MousePointerConstants.vbSizeAll
    Case IDC_WAIT:        mp = MousePointerConstants.vbHourglass
    Case IDC_NO:          mp = MousePointerConstants.vbNoDrop
    Case IDC_APPSTARTING: mp = MousePointerConstants.vbArrowHourglass
    Case IDC_HELP:        mp = MousePointerConstants.vbArrowQuestion
    Case Else:            mp = MousePointerConstants.vbCustom
    End Select
End Property
Public Property Let MousePointer(ByVal Value As Integer) 'MousePointerConstants)
    Dim mp   As MousePointerConstants: mp = Value
    Dim idmp As Long
    Select Case mp
    Case MousePointerConstants.vbArrow:          idmp = IDC_ARROW
    Case MousePointerConstants.vbIbeam:          idmp = IDC_IBEAM
    Case MousePointerConstants.vbCrosshair:      idmp = IDC_CROSS
    Case MousePointerConstants.vbUpArrow:        idmp = IDC_UPARROW
    Case MousePointerConstants.vbSizePointer:    idmp = IDC_SIZEALL 'IDC_SIZE
    Case MousePointerConstants.vbIconPointer:    idmp = IDC_ICON
         If m_MouseIcon Is Nothing Then idmp = IDC_ARROW Else If m_MouseIcon.Handle = 0 Then idmp = IDC_ARROW
    Case MousePointerConstants.vbSizeNWSE:       idmp = IDC_SIZENWSE
    Case MousePointerConstants.vbSizeNESW:       idmp = IDC_SIZENESW
    Case MousePointerConstants.vbSizeWE:         idmp = IDC_SIZEWE
    Case MousePointerConstants.vbSizeNS:         idmp = IDC_SIZENS
    Case MousePointerConstants.vbSizeAll:        idmp = IDC_SIZEALL
    Case MousePointerConstants.vbHourglass:      idmp = IDC_WAIT
    Case MousePointerConstants.vbNoDrop:         idmp = IDC_NO
    Case MousePointerConstants.vbArrowHourglass: idmp = IDC_APPSTARTING
    Case MousePointerConstants.vbArrowQuestion:  idmp = IDC_HELP
    Case Else:
    End Select
    m_MousePointer = idmp
End Property

'Public Sub Move(): End Sub
Public Property Get Moveable() As Boolean: End Property

Public Sub OLEDrag(): End Sub
Public Property Get OLEDropMode() As Integer: End Property
Public Sub PaintPicture(): End Sub
Public Property Get Palette() As IPictureDisp: End Property
Public Property Get PaletteMode() As Integer: End Property
Public Property Get Picture() As IPictureDisp: End Property


Public Sub PopupMenu(Menu As Object, Optional Flags, Optional X, Optional Y, Optional DefaultMenu): End Sub
Public Sub PrintForm(): End Sub

Public Sub Refresh()
    RaiseEvent Paint
End Sub

Public Property Get RightToLeft() As Boolean
    'WS_EX_RTLREADING
End Property
Public Property Let RightToLeft(ByVal Value As Boolean)
    'WS_EX_RTLREADING
End Property

Public Sub SScale(): End Sub

Public Property Get ScaleMode() As Integer
    'Gibt einen Wert zurück, der die Maßeinheiten für Objektkoordinaten bestimmt, wenn Grafikmethoden verwendet oder Steuerelemente positioniert werden,
    'oder legt diesen fest.
    ScaleMode = m_ScaleMode
'Enum VBRUN.ScaleModeConstants
    'vbUser              =  0
    'vbTwips             =  1
    'vbPoints            =  2
    'vbPixels            =  3
    'vbCharacters        =  4
    'vbInches            =  5
    'vbMillimeters       =  6
    'vbCentimeters       =  7
    'vbHimetric          =  8
    'vbContainerPosition =  9
    'vbContainerSize     = 10
'End Enum
End Property
Public Property Let ScaleMode(ByVal Value As Integer)
    'Gibt einen Wert zurück, der die Maßeinheiten für Objektkoordinaten bestimmt, wenn Grafikmethoden verwendet oder Steuerelemente positioniert werden,
    'oder legt diesen fest.
    m_ScaleMode = Value
End Property
Public Property Get ScaleLeft() As Single: End Property
Public Property Get ScaleTop() As Single: End Property
Public Property Get ScaleWidth() As Single: End Property
Public Property Get ScaleHeight() As Single: End Property

Public Function ScaleX(Width As Single, Optional FromScale, Optional ToScale) As Single: End Function
Public Function ScaleY(Height As Single, Optional FromScale, Optional ToScale): End Function

Public Sub SetFocus(): End Sub


'Public Sub Show(Optional Modal As FormShowConstants, Optional OwnerForm As Object): End Property
Public Property Get ShowInTaskbar() As Boolean
    'schreibgeschützt
    'Mitglied von VB.Form
    'Bestimmt, ob ein Formular oder ein MDI-Formular in der Task-Leiste von Windows 95 angezeigt wird.
End Property
Public Property Let ShowInTaskbar(ByVal Value As Boolean)
    '
End Property

'Enum StartUpPositionConstants
'    vbStartUpManual = 0
'    vbStartUpOwner = 1
'    vbStartUpScreen = 2
'    vbStartUpWindowsDefault = 3
'End Enum

Public Property Get StartUpPosition() As Integer
    'schreibgeschützt
    'Mitglied von VB.Form
    'Gibt einen Wert zurück oder legt einen Wert fest, der die Position eines Formulars angibt, wenn es zum ersten Mal angezeigt wird.
End Property
Public Property Let StartUpPosition(ByVal Value As Integer)
    '
End Property

Public Property Get Tag() As Variant
    'Mitglied von VB.Form
    'Speichert alle zusätzlichen Daten, die Sie für Ihr Programm benötigen.
    Tag = m_Tag
End Property
Public Property Let Tag(ByVal Value As Variant)
     m_Tag = Value
End Property
Public Property Set Tag(ByVal Value As Variant)
     Set m_Tag = Value
End Property

Public Function TextHeight(Str As String) As Single
    'Mitglied von VB.Form
    'Gibt zurück, mit welcher Höhe eine Zeichenfolge in der aktuellen Schriftart ausgegeben würde.
End Function

Public Function TextWidth(Str As String) As Single
    'Mitglied von VB.Form
    'Gibt zurück, mit welcher Breite eine Zeichenfolge in der aktuellen Schriftart ausgegeben würde.
End Function

'Public Property Get Top() As Single: End Property

Public Sub ValidateControls()
    'Mitglied von VB.Form
    'Überprüfen Sie den Inhalt des letzten Steuerelements auf dem Formular, bevor Sie das Formular verlassen.
End Sub

Public Property Get WhatsThisButton() As Boolean
    'schreibgeschützt
    'Mitglied von VB.Form
    'Gibt zurück oder legt fest, ob die Schaltfläche "Direkthilfe" in der Titelleiste eines Formulars oder MDI-Formulars angezeigt wird.
End Property
Public Property Get WhatsThisHelp() As Boolean
    'schreibgeschützt
    'Mitglied von VB.Form
    'Gibt zurück oder legt fest, ob für die kontextbezogene Hilfe das von der Windows 95-Hilfe
    'bereitgestellte Direkthilfe-Popup oder das Hilfe-Hauptfenster verwendet wird.
End Property

Public Sub WhatsThisMode()
    'Bereitet die Anwendung so vor, daß eine Direkthilfe für das ausgewählte Objekt angezeigt wird.
    
End Sub

'Public Property Get Width() As Single: End Property
'Enum FormWindowStateConstants
'    vbNormal = 0
'    vbMinimized = 1
'    vbMaximized = 2
'End Enum
Public Property Get WindowState() As Integer 'FormWindowStateConstants
    'Mitglied von VB.Form
    'Gibt den Anzeigezustand eines Formularfensters zur Laufzeit zurück oder legt diesen fest.
    Dim ws As FormWindowStateConstants
    Select Case m_WindowState
    Case SW_NORMAL:   ws = FormWindowStateConstants.vbNormal
    Case SW_MINIMIZE: ws = FormWindowStateConstants.vbMinimized
    Case SW_MAXIMIZE: ws = FormWindowStateConstants.vbMaximized
    End Select
    WindowState = ws
End Property
Public Property Let WindowState(ByVal Value As Integer)  'FormWindowStateConstants)
    Dim ws As FormWindowStateConstants: ws = Value
    Select Case ws
    Case vbNormal:    m_WindowState = SW_NORMAL
    Case vbMinimized: m_WindowState = SW_MINIMIZE
    Case vbMaximized: m_WindowState = SW_MAXIMIZE
    End Select
    Dim hr As Long
    With m_Window
        If m_hWnd Then
            hr = ShowWindow(m_hWnd, m_WindowState)
        End If
    End With
End Property
' ^ ############################## ^ '    OM TODO:    ' ^ ############################## ^ '

Public Sub ZOrder(Optional Position)
    'Legt ein angegebenes Objekt an der ersten oder letzten Stelle der Z-Reihenfolge innerhalb der zugehörigen Grafikstufe ab.
    'EWndShowZPos ? ?
    With m_Window
        SetWindowPos m_hWnd, 0&, .X, .Y, .cx, .cy, CLng(Position)
    End With
End Sub

'Public Property Get Caption() As String
'    Caption = m_Name 'm_Window.lpszName
'End Property
'Public Property Let Caption(ByVal Value As String)
'    m_Caption = Value
'End Property

Public Property Get Style() As EWndStyle
    Style = m_Window.Style
End Property
Public Property Let Style(ByVal Value As EWndStyle)
    m_Window.Style = Value
    If m_hWnd Then MWindow.WindowStyle(m_hWnd) = Value
End Property
Public Function Style_ToStr() As String
    Style_ToStr = MWindow.EWndStyle_ToStr(Me.Style)
End Function

Public Property Get StyleEx() As EWndStyleEx
    StyleEx = m_Window.dwExStyle
End Property
Public Property Let StyleEx(ByVal Value As EWndStyleEx)
    m_Window.dwExStyle = Value
    If m_hWnd Then MWindow.WindowStyleEx(m_hWnd) = Value
End Property
Public Function StyleEx_ToStr() As String
    StyleEx_ToStr = MWindow.EWndStyleEx_ToStr(Me.StyleEx)
End Function

Public Property Get hWnd() As LongPtr
    hWnd = m_hWnd
End Property

Public Property Get Left() As Single
    Left = m_Window.X
End Property
Public Property Let Left(ByVal Value As Single)
    With m_Window
        .X = Value
        MoveWindow m_hWnd, .X, .Y, .cx, .cy, True
    End With
End Property

Public Property Get Top() As Single
    Top = m_Window.Y
End Property
Public Property Let Top(ByVal Value As Single)
    With m_Window
        .Y = Value
        MoveWindow m_hWnd, .X, .Y, .cx, .cy, True
    End With
End Property

Public Property Get Width() As Single
    Width = m_Window.cx
End Property
Public Property Let Width(ByVal Value As Single)
    With m_Window
        .cx = Value
        MoveWindow m_hWnd, .X, .Y, .cx, .cy, True
    End With
End Property

Public Property Get Height() As Single
    Height = m_Window.cy
End Property
Public Property Let Height(ByVal Value As Single)
    With m_Window
        .cy = Value
        MoveWindow m_hWnd, .X, .Y, .cx, .cy, True
    End With
End Property

Public Sub Move(ByVal Left As Single, Optional ByVal Top, Optional ByVal Width, Optional ByVal Height)
    With m_Window
        .X = CLng(Left)
        If Not IsMissing(Top) Then .Y = CLng(Top)
        If Not IsMissing(Width) Then .cx = CLng(Width)
        If Not IsMissing(Height) Then .cy = CLng(Height)
        'SetWindowPos m_hWnd, 0&, .Left, .Top, .Width, .Height
        MoveWindow m_hWnd, .X, .Y, .cx, .cy, True
    End With
End Sub

Friend Sub OnActivate():                                                               RaiseEvent Activate:                                                                                                      End Sub
Friend Sub OnClick():                                                                  RaiseEvent Click:                                                                                                         End Sub
Friend Sub OnDblClick():                                                               RaiseEvent DblClick:                                                                                                      End Sub
Friend Sub OnDeactivate():                                                             RaiseEvent Deactivate:                                                                                                    End Sub
Friend Sub OnDragDrop(Source As Control, X As Single, Y As Single):                    RaiseEvent DragDrop(Source, X, Y):                                                                                        End Sub
Friend Sub OnDragOver(Source As Control, X As Single, Y As Single, State As Integer):  RaiseEvent DragOver(Source, X, Y, State):                                                                                 End Sub
Friend Sub OnGotFocus():                                                               RaiseEvent GotFocus:                                                                                                      End Sub
Friend Sub OnInitialize():                                                             RaiseEvent Initialize:                                                                                                    End Sub
Friend Sub OnKeyDown(KeyCode As Integer, Shift As Integer):                            RaiseEvent KeyDown(KeyCode, Shift):                                                                                       End Sub
Friend Sub OnKeyPress(KeyAscii As Integer):                                            RaiseEvent KeyPress(KeyAscii):                                                                                            End Sub
Friend Sub OnKeyUp(KeyCode As Integer, Shift As Integer):                              RaiseEvent KeyUp(KeyCode, Shift):                                                                                         End Sub
Friend Sub OnLinkClose():                                                              RaiseEvent LinkClose:                                                                                                     End Sub
Friend Sub OnLinkError(LinkErr As Integer):                                            RaiseEvent LinkError(LinkErr):                                                                                            End Sub
Friend Sub OnLinkExecute(CmdStr As String, Cancel As Integer):                         RaiseEvent LinkExecute(CmdStr, Cancel):                                                                                   End Sub
Friend Sub OnLinkOpen(Cancel As Integer):                                              RaiseEvent LinkOpen(Cancel):                                                                                              End Sub
Friend Sub OnLoad():                                                                   RaiseEvent Load:                                                                                                          End Sub
Friend Sub OnLostFocus():                                                              RaiseEvent LostFocus:                                                                                                     End Sub
Friend Sub OnMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single): RaiseEvent MouseDown(Button, Shift, X, Y):                                                                                End Sub
Friend Sub OnMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single): RaiseEvent MouseMove(Button, Shift, X, Y):                                                                                End Sub
Friend Sub OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single):   RaiseEvent MouseUp(Button, Shift, X, Y):                                                                                  End Sub
Friend Sub OnOLECompleteDrag(Effect As Long):                                          RaiseEvent OLECompleteDrag(Effect):                                                                                       End Sub
Friend Sub OnOLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single):                   RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y):        End Sub
Friend Sub OnOLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer): RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State): End Sub
Friend Sub OnOLEGiveFeedback(Effect As Long, DefaultCursors As Boolean):               RaiseEvent OLEGiveFeedback(Effect, DefaultCursors):                                                                       End Sub
Friend Sub OnOLESetData(Data As DataObject, DataFormat As Integer):                    RaiseEvent OLESetData(Data, DataFormat):                                                                                  End Sub
Friend Sub OnOLEStartDrag(Data As DataObject, AllowedEffects As Long):                 RaiseEvent OLEStartDrag(Data, AllowedEffects):                                                                            End Sub
Friend Sub OnPaint():                                                                  RaiseEvent Paint:                                                                                                         End Sub
Friend Sub OnQueryUnload(Cancel As Integer, UnloadMode As Integer):                    RaiseEvent QueryUnload(Cancel, UnloadMode):                                                                               End Sub
Friend Sub OnResize():                                                                 RaiseEvent Resize:                                                                                                        End Sub
Friend Sub OnTerminate():                                                              RaiseEvent Terminate:                                                                                                     End Sub
Friend Sub OnUnload(Cancel As Integer):                                                RaiseEvent Unload(Cancel):                                                                                                End Sub

'' v #################### v '     Local ErrHandler      ' v #################### v '
Private Function ErrHandler(ByVal FuncName As String, _
                            Optional ByVal AddInfo As String, _
                            Optional WinApiError, _
                            Optional bLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKCancel, _
                            Optional bRetry As Boolean) As VbMsgBoxResult

    If bRetry Then

        ErrHandler = MessErrorRetry(TypeName(Me), FuncName, AddInfo, WinApiError, bErrLog)

    Else

        ErrHandler = MessError(TypeName(Me), FuncName, AddInfo, WinApiError, bLoud, bErrLog, vbDecor)

    End If

End Function

