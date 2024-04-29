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
    VBFormStyle = WS_MAXIMIZEBOX Or WS_TABSTOP Or WS_GROUP Or WS_MINIMIZEBOX Or WS_SIZEBOX Or WS_THICKFRAME Or WS_SYSMENU Or WS_DLGFRAME Or WS_BORDER Or WS_CAPTION Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
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
    VBFormStyleEx = WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW
End Enum

Public Enum EClsStyle
    CS_VREDRAW = &H1             '    Zeichnet das gesamte Fenster neu, wenn eine Bewegung oder Größenanpassung die Höhe des Clientbereichs ändert.
    CS_HREDRAW = &H2             '    Zeichnet das gesamte Fenster neu, wenn durch eine Verschiebung oder Größenanpassung die Breite des Clientbereichs geändert wird.
    CS_DBLCLKS = &H8             '    Sendet eine Doppelklicknachricht an die Fensterprozedur, wenn der Benutzer mit der Maus doppelklicken, während sich der Cursor in einem Fenster befindet, das zur -Klasse gehört.
    CS_OWNDC = &H20              '    Weist jedem Fenster in der -Klasse einen eindeutigen Gerätekontext zu.
    CS_CLASSDC = &H40            '    Weist einen Gerätekontext zu, der von allen Fenstern in der -Klasse gemeinsam verwendet werden soll. Da Fensterklassen prozessspezifisch sind, ist es möglich, dass mehrere Threads einer Anwendung ein Fenster derselben Klasse erstellen. Es ist auch möglich, dass die Threads versuchen, den Gerätekontext gleichzeitig zu verwenden. In diesem Fall lässt das System nur einen Thread zu, um den Zeichnungsvorgang erfolgreich abzuschließen.
    CS_PARENTDC = &H80           '    Legt das Beschneidungsrechteck des untergeordneten Fensters auf das des übergeordneten Fensters fest, damit das untergeordnete Element auf dem übergeordneten Fenster zeichnen kann. Ein Fenster mit dem CS_PARENTDC Stilbits empfängt einen regulären Gerätekontext aus dem Gerätekontextcache des Systems. Es gibt dem untergeordneten Element nicht die Gerätekontext- oder Gerätekontexteinstellungen des übergeordneten Elements. Die Angabe CS_PARENTDC verbessert die Leistung einer Anwendung.
    CS_NOCLOSE = &H200           '    Deaktiviert Schließen im Fenstermenü.
    CS_SAVEBITS = &H800          '    Speichert als Bitmap den Teil des Bildschirmbilds, der von einem Fenster dieser Klasse verdeckt wird. Wenn das Fenster entfernt wird, verwendet das System die gespeicherte Bitmap, um das Bildschirmbild wiederherzustellen, einschließlich anderer verdeckter Fenster. Daher sendet das System keine WM_PAINT Nachrichten an Fenster, die verdeckt wurden, wenn der von der Bitmap verwendete Arbeitsspeicher nicht verworfen wurde und wenn das gespeicherte Bild durch andere Bildschirmaktionen nicht ungültig gemacht wurde. Dieser Stil ist nützlich für kleine Fenster (z. B. Menüs oder Dialogfelder), die kurz angezeigt und dann entfernt werden, bevor andere Bildschirmaktivitäten stattfinden. Durch diese Formatvorlage wird die zum Anzeigen des Fensters erforderliche Zeit erhöht, da das System zunächst Arbeitsspeicher zum Speichern der Bitmap zuweisen muss.
    CS_BYTEALIGNCLIENT = &H1000  '    Richtet den Clientbereich des Fensters an einer Bytegrenze (in x-Richtung) aus. Diese Formatvorlage wirkt sich auf die Breite des Fensters und seine horizontale Platzierung auf der Anzeige aus.
    CS_BYTEALIGNWINDOW = &H2000  '    Richtet das Fenster an einer Bytegrenze (in x-Richtung) aus. Diese Formatvorlage wirkt sich auf die Breite des Fensters und seine horizontale Platzierung auf der Anzeige aus.
    CS_GLOBALCLASS = &H4000      '    Gibt an, dass die Fensterklasse eine globale Anwendungsklasse ist. Weitere Informationen finden Sie im Abschnitt "Globale Anwendungsklassen" unter Informationen zu Fensterklassen.
    CS_DROPSHADOW = &H20000      '    Aktiviert den Schlagschatteneffekt für ein Fenster. Der Effekt wird über SPI_SETDROPSHADOW ein- und ausgeschaltet. In der Regel ist dies für kleine, kurzlebige Fenster wie Menüs aktiviert, um ihre Z-Reihenfolge-Beziehung zu anderen Fenstern hervorzuheben. Fenster, die aus einer Klasse mit diesem Stil erstellt werden, müssen Fenster der obersten Ebene sein. Sie dürfen keine Kinderfenster sein.
End Enum


Public Enum EWndShowZPos
    SWP_ASYNCWINDOWPOS = &H4000& '    Wenn der aufrufende Thread und der Thread, der das Fenster besitzt, an verschiedene Eingabewarteschlangen angefügt sind, sendet das System die Anforderung an den Thread, der das Fenster besitzt. Dadurch wird verhindert, dass der aufrufende Thread seine Ausführung blockiert, während andere Threads die Anforderung verarbeiten.
    SWP_DEFERERASE = &H2000&     '    Verhindert die Generierung der WM_SYNCPAINT Nachricht.
    SWP_DRAWFRAME = &H20&        '    Zeichnet einen Rahmen (definiert in der Klassenbeschreibung des Fensters) um das Fenster.
    SWP_FRAMECHANGED = &H20&     '    Wendet mithilfe der SetWindowLong-Funktion neue Frameformatvorlagen an. Sendet eine WM_NCCALCSIZE Nachricht an das Fenster, auch wenn die Größe des Fensters nicht geändert wird. Wenn dieses Flag nicht angegeben ist, wird WM_NCCALCSIZE nur gesendet, wenn die Größe des Fensters geändert wird.
    SWP_HIDEWINDOW = &H80&       '    Blendet das Fenster aus.
    SWP_NOACTIVATE = &H10&       '    Aktiviert das Fenster nicht. Wenn dieses Flag nicht festgelegt ist, wird das Fenster aktiviert und an den Anfang der obersten oder nicht obersten Gruppe verschoben (abhängig von der Einstellung des hWndInsertAfter-Parameters ).
    SWP_NOCOPYBITS = &H100&      '    Verwirft den gesamten Inhalt des Clientbereichs. Wenn dieses Flag nicht angegeben ist, werden die gültigen Inhalte des Clientbereichs gespeichert und wieder in den Clientbereich kopiert, nachdem das Fenster dimensioniert oder neu positioniert wurde.
    SWP_NOMOVE = &H2&            '    Behält die aktuelle Position bei (ignoriert die X - und Y-Parameter ).
    SWP_NOOWNERZORDER = &H200&   '    Ändert die Position des Besitzerfensters in der Z-Reihenfolge nicht.
    SWP_NOREDRAW = &H8&          '    Änderungen werden nicht neu gezeichnet. Wenn dieses Flag festgelegt ist, erfolgt keine Neulackierung jeglicher Art. Dies gilt für den Clientbereich, den Nichtclientbereich (einschließlich Titelleiste und Bildlaufleisten) und alle Teile des übergeordneten Fensters, die aufgrund des Verschiebens des Fensters aufgedeckt wurden. Wenn dieses Flag festgelegt ist, muss die Anwendung alle Teile des Fensters und des übergeordneten Fensters, die neu ausgezeichnet werden müssen, explizit für ungültig erklären oder neu zeichnen.
    SWP_NOREPOSITION = &H200&    '    Identisch mit dem flag SWP_NOOWNERZORDER .
    SWP_NOSENDCHANGING = &H400&  '    Verhindert, dass das Fenster die WM_WINDOWPOSCHANGING-Nachricht empfängt.
    SWP_NOSIZE = &H1&            '    Behält die aktuelle Größe bei (ignoriert die Parameter cx und cy ).
    SWP_NOZORDER = &H4&          '    Behält die aktuelle Z-Reihenfolge bei (ignoriert den Parameter hWndInsertAfter ).
    SWP_SHOWWINDOW = &H40&       '    Zeigt das Fenster an.
End Enum

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
Private Type CREATESTRUCTW
    lpCreateParams As LongPtr
    hInstance      As LongPtr
    hMenu          As LongPtr
    hwndParent     As LongPtr
    cy             As Long
    cx             As Long
    y              As Long
    x              As Long
    Style          As EWndStyle
    lpszName       As LongPtr
    lpszClass      As LongPtr
    dwExStyle      As EWndStyleEx
End Type

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
  hIconSm       As LongPtr   'HICON
End Type '} WNDCLASSEXA, *PWNDCLASSEXA, *NPWNDCLASSEXA, *LPWNDCLASSEXA;


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
Private Declare Function CreateWindowExW Lib "user32" ( _
    ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, _
    ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As LongPtr, _
    ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr

'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-setwindowpos
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-movewindow
'BOOL MoveWindow(
'  [in] HWND hWnd,
'  [in] int  X,
'  [in] int  Y,
'  [in] int  nWidth,
'  [in] int  nHeight,
'  [in] BOOL bRepaint
');
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'https://learn.microsoft.com/de-de/windows/win32/api/errhandlingapi/nf-errhandlingapi-getlasterror
Private Declare Function GetLastError Lib "kernel32" () As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nc-winuser-wndproc
'LRESULT Wndproc( HWND unnamedParam1, UINT unnamedParam2, WPARAM unnamedParam3, LPARAM unnamedParam4 )


'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-registerclassexw
'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-registerclassexw
'ATOM RegisterClassExA( [in] const WNDCLASSEXA *unnamedParam1 );
Private Declare Function RegisterClassExW Lib "user32" (pcWndClassEx As WNDCLASSEXW) As Long
'https://learn.microsoft.com/de-de/windows/win32/winmsg/using-window-classes


Private m_ClsName As String
Private m_WndCls  As WNDCLASSEXW
Private m_ClsAtom As Long
Private m_Name    As String
Private m_Window  As CREATESTRUCTW
Private m_hWnd    As LongPtr

Public Event Click()
Public Event Resize()

Private Sub Class_Initialize()
    m_ClsName = "ThunderVB64uWindow" & vbNullChar
    If Not Register Then
        MsgBox "Could not register class: " & m_ClsName
        Exit Sub
    End If
    m_Name = "VBForm" & vbNullChar
    With m_Window
        'entweder so
        '.lpszClass = m_ClsAtom
        'oder auch so:
        .lpszClass = StrPtr(m_ClsName)
        
        .lpszName = StrPtr(m_Name)
        .hInstance = App.hInstance
        .Style = WS_MAXIMIZEBOX Or WS_TABSTOP Or WS_GROUP Or WS_MINIMIZEBOX Or _
                 WS_SIZEBOX Or WS_THICKFRAME Or WS_SYSMENU Or WS_DLGFRAME Or WS_BORDER Or _
                 WS_CAPTION Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_VISIBLE
        .dwExStyle = WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW
        .x = 33
        .y = 26
        .cx = 320 'px VB-default
        .cy = 240 'px VB-default
    End With
End Sub

Friend Sub New_(Name As String, Optional ByVal Style As EWndStyle = EWndStyle.VBFormStyle, Optional ByVal StyleEx As EWndStyleEx = EWndStyleEx.VBFormStyleEx)
    m_Name = Name & vbNullChar
    With m_Window
        '.hInstance = App.hInstance
        .lpszName = StrPtr(m_Name)
        .Style = Style
        .dwExStyle = StyleEx
    End With
End Sub

Private Function Register() As Boolean
    With m_WndCls
        .cbSize = LenB(m_WndCls) '48?
        .lpszClassName = StrPtr(m_ClsName)
        .lpfnWndProc = FncPtr(AddressOf MWindow.WndProc)
        .hInstance = App.hInstance
        '.style = CS_VREDRAW Or CS_PARENTDC Or CS_OWNDC Or CS_HREDRAW
    End With
    m_ClsAtom = RegisterClassExW(m_WndCls)
    Register = CBool(m_ClsAtom)
End Function

Public Sub Load()
'Try: On Error GoTo Catch
    With m_Window
        m_hWnd = CreateWindowExW(.dwExStyle, .lpszClass, .lpszName, .Style, .x, .y, .cx, .cy, .hwndParent, .hMenu, .hInstance, ByVal .lpCreateParams)
        If m_hWnd = 0 Then
            'GoTo Catch
            ErrHandler "Load", "hWnd = 0"
        End If
    End With
'    Exit Sub
'Catch:
End Sub

Friend Sub OnClick()
    RaiseEvent Click
End Sub

Friend Sub OnResize()
    RaiseEvent Resize
End Sub

Public Property Get Caption() As String
    Caption = m_Window.lpszName
End Property

Public Property Get Style() As EWndStyle
    Style = m_Window.Style
End Property

Public Property Get StyleEx() As EWndStyle
    StyleEx = m_Window.dwExStyle
End Property

Public Property Get hWnd() As LongPtr
    hWnd = m_hWnd
End Property

'Public Property Get Caption() As String
'    Caption = m_Caption
'End Property
'Public Property Let Caption(ByVal Value As String)
'    m_Caption = Value
'End Property

Public Property Get Left() As Long
    Left = m_Rect.Left
End Property
Public Property Let Left(ByVal Value As Long)
    m_Rect.Left = Value
End Property

Public Property Get Top() As Long
    Top = m_Rect.Top
End Property
Public Property Let Top(ByVal Value As Long)
    m_Rect.Top = Value
End Property

Public Property Get Width() As Long
    Width = m_Rect.Width
End Property
Public Property Let Width(ByVal Value As Long)
    m_Rect.Width = Value
End Property

Public Property Get Height() As Long
    Height = m_Rect.Height
End Property
Public Property Let Height(ByVal Value As Long)
    m_Rect.Height = Value
End Property

Public Sub Move(ByVal Left As Single, Optional ByVal Top, Optional ByVal Width, Optional ByVal Height)
    With m_Rect
        .Left = CLng(Left)
        If Not IsMissing(Top) Then .Top = CLng(Top)
        If Not IsMissing(Width) Then .Width = CLng(Width)
        If Not IsMissing(Height) Then .Height = CLng(Height)
        'SetWindowPos m_hWnd, 0&, .Left, .Top, .Width, .Height
        MoveWindow m_hWnd, .Left, .Top, .Width, .Height
    End With
End Sub


Private Function EWndStyle_ToStr(e As EWndStyle) As String
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

Private Function EWndStyleEx_ToStr(e As EWndStyleEx) As String
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


