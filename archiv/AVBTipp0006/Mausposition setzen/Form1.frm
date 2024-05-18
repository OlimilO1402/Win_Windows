VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "www.ActiveVB.de"
   ClientHeight    =   10725
   ClientLeft      =   3090
   ClientTop       =   2310
   ClientWidth     =   20475
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   189.177
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   361.157
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Move Mouse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   18000
      Top             =   7680
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !

Option Explicit

'Deklaration: Globale Form API-Typen
Private Type POINTAPI
    x As Long
    y As Long
End Type

'Deklaration: Globale Form API-Funktionen

'https://learn.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-clienttoscreen
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

'
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private m_ScaleMode As ScaleModeConstants

Private Sub Form_Load()
    m_ScaleMode = Me.ScaleMode ' vbPixels
End Sub

Private Sub MoveMouse(ByVal x As Long, ByVal y As Long)
    'Deklaration: Lokale Prozedur-Variablen
    Dim udtP As POINTAPI
    
    'udtP.x = x
    'udtP.y = y
    
    'Mauszeiger versetzen
    ClientToScreen Me.hwnd, udtP
    SetCursorPos udtP.x + x, udtP.y + y
End Sub

Private Sub Command1_Click()
    
    'Die Maus soll in die Mitte des Rechts verschoben werden:
    Dim L As Long: L = Shape1.Left + Shape1.Width / 2
    Dim T As Long: T = Shape1.Top + Shape1.Height / 2
    
    L = Form1.ScaleX(L, Form1.ScaleMode, ScaleModeConstants.vbPixels)
    T = Form1.ScaleY(T, Form1.ScaleMode, ScaleModeConstants.vbPixels)
    
    MoveMouse L, T
End Sub

'Private Function SScaleX(ByVal Width As Single, optinoal ByVal FromScale As ScaleModeConstants, optional ByVal ToScale As ScaleModeConstants) As Single
Private Function SScaleX(ByVal Width As Single, Optional FromScale, Optional ToScale) As Single
    Dim frsc As ScaleModeConstants: If Not IsMissing(FromScale) Then frsc = CLng(FromScale) Else If Not IsMissing(FromScale) Then frsc = m_ScaleMode Else frsc = vbUser
    Dim tosc As ScaleModeConstants: If Not IsMissing(ToScale) Then tosc = CLng(ToScale) Else tosc = m_ScaleMode
    
    'Dim frsc As ScaleModeConstants: If Not IsMissing(FromScale) Then frsc = CLng(FromScale) Else frsc = m_ScaleMode
    'Dim tosc As ScaleModeConstants: If Not IsMissing(ToScale) Then tosc = CLng(ToScale) Else If Not IsMissing(FromScale) Then tosc = m_ScaleMode Else tosc = vbMillimeters
    
    
    'first scale everything to vbPixels
    Const dpi  As Double = 96#        ' dpi = dots per inch, dots = pixel!
    Const cmpi As Double = 2.54       ' centimeter per inch
    Const mmpi As Double = 25.4       ' millimeter per inch
    Const popi As Double = 72#        ' points per inch
    Const ppch As Double = 8#
    
    Const cmpp As Double = dpi / cmpi ' 37.79524
    Const mmpp As Double = dpi / mmpi '  3.779524
    Const chpp As Double = 1 / ppch   '  0.125
    Const ipp  As Double = 1 / dpi    '  0.01041667
    'Umrechnungsfaktor in pixel
    Dim f1 As Double
    Select Case frsc
    Case ScaleModeConstants.vbUser:        f1 = 1 / 26.45833 '0.01
    Case ScaleModeConstants.vbTwips:       f1 = 1 / Screen.TwipsPerPixelX '  0.06666667
    Case ScaleModeConstants.vbPoints:      f1 = dpi / popi                '  1.333333
    Case ScaleModeConstants.vbPixels:      f1 = 1                         '  1
    Case ScaleModeConstants.vbCharacters:  f1 = ppch                      '  8
    Case ScaleModeConstants.vbInches:      f1 = dpi                       ' 96
    Case ScaleModeConstants.vbMillimeters: f1 = dpi / mmpi                '  3.779524
    Case ScaleModeConstants.vbCentimeters: f1 = dpi / cmpi                ' 37.79524
    Case ScaleModeConstants.vbHimetric:    f1 = dpi / mmpi / 100          '  0.03779528
    End Select
    'Umrechnungsfaktor in ZielScale
    Dim f2 As Double
    Select Case tosc
    Case ScaleModeConstants.vbTwips:       f2 = Screen.TwipsPerPixelX     ' 15
    Case ScaleModeConstants.vbPoints:      f2 = popi / dpi                '  0.75
    Case ScaleModeConstants.vbPixels:      f2 = 1                         '  1
    Case ScaleModeConstants.vbCharacters:  f2 = chpp                      '  0.125
    Case ScaleModeConstants.vbInches:      f2 = 1 / dpi                   '  0.01041667
    Case ScaleModeConstants.vbMillimeters: f2 = mmpi / dpi                '  0.2645836
    Case ScaleModeConstants.vbCentimeters: f2 = cmpi / dpi                '  0.02645836
    Case ScaleModeConstants.vbHimetric:    f2 = mmpi / dpi * 100 '0     ' 26.45833
    End Select
    SScaleX = Width * f1 * f2
End Function

'Enum ScaleModeConstants
'   vbUser = 0
'   vbTwips = 1
'   vbPoints = 2
'   vbPixels = 3
'   vbCharacters = 4
'   vbInches = 5
'   vbMillimeters = 6
'   vbCentimeters = 7
'   vbHimetric = 8
'   vbContainerPosition = 9
'   vbContainerSize = 10
'End Enum

Private Sub Command2_Click()
    Dim pp1 As Single: pp1 = 1
    
    '0.5669292
    '1.76388868
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbUser, ScaleModeConstants.vbPixels)
    Debug.Print ScaleX(pp1) ', ScaleModeConstants.vbTwips) ', ScaleModeConstants.vbPixels)       '  6.666667E-02
    Debug.Print ScaleX(pp1) ', ScaleModeConstants.vbPoints) ', ScaleModeConstants.vbPixels)      '  1.333333
    Debug.Print ScaleX(pp1) ', ScaleModeConstants.vbPixels) ', ScaleModeConstants.vbPixels)      '  1
    Debug.Print ScaleX(pp1) ', ScaleModeConstants.vbCharacters) ', ScaleModeConstants.vbPixels)  '  8
    Debug.Print ScaleX(pp1) ', ScaleModeConstants.vbInches) ', ScaleModeConstants.vbPixels)      ' 96
    Debug.Print ScaleX(pp1) ', ScaleModeConstants.vbMillimeters) ', ScaleModeConstants.vbPixels) '  3.779524
    Debug.Print ScaleX(pp1) ', ScaleModeConstants.vbCentimeters) ', ScaleModeConstants.vbPixels) ' 37.79524
    Debug.Print ScaleX(pp1) ', ScaleModeConstants.vbHimetric) ', ScaleModeConstants.vbPixels)    '  3.779528E-02
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbContainerPosition, ScaleModeConstants.vbPixels)
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbContainerSize, ScaleModeConstants.vbPixels)
    Debug.Print ""
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbUser, ScaleModeConstants.vbPixels)
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbTwips) ', ScaleModeConstants.vbPixels)       '  6.666667E-02
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbPoints) ', ScaleModeConstants.vbPixels)      '  1.333333
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbPixels) ', ScaleModeConstants.vbPixels)      '  1
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbCharacters) ', ScaleModeConstants.vbPixels)  '  8
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbInches) ', ScaleModeConstants.vbPixels)      ' 96
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbMillimeters) ', ScaleModeConstants.vbPixels) '  3.779524
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbCentimeters) ', ScaleModeConstants.vbPixels) ' 37.79524
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbHimetric) ', ScaleModeConstants.vbPixels)    '  3.779528E-02
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbContainerPosition, ScaleModeConstants.vbPixels)
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbContainerSize, ScaleModeConstants.vbPixels)
    Debug.Print ""
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbUser, ScaleModeConstants.vbPixels)
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)       '  6.666667E-02
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbPoints, ScaleModeConstants.vbPixels)      '  1.333333
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbPixels)      '  1
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbCharacters, ScaleModeConstants.vbPixels)  '  8
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbInches, ScaleModeConstants.vbPixels)      ' 96
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbMillimeters, ScaleModeConstants.vbPixels) '  3.779524
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbCentimeters, ScaleModeConstants.vbPixels) ' 37.79524
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbHimetric, ScaleModeConstants.vbPixels)    '  3.779528E-02
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbContainerPosition, ScaleModeConstants.vbPixels)
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbContainerSize, ScaleModeConstants.vbPixels)
    Debug.Print ""
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbUser, ScaleModeConstants.vbPixels)
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips)       ' 15
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbPoints)      '  0.75
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbPixels)      '  1
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbCharacters)  '  0.125
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbInches)      '  1.041667E-02
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbMillimeters) '  0.2645836
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbCentimeters) '  2.645836E-02
    Debug.Print ScaleX(pp1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbHimetric)    ' 26.45833
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbContainerPosition, ScaleModeConstants.vbPixels)
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbContainerSize, ScaleModeConstants.vbPixels)
    
    Debug.Print ""
    Debug.Print ""
    
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbUser, ScaleModeConstants.vbPixels)
    Debug.Print SScaleX(pp1) ', ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)       '  6.666667E-02
    Debug.Print SScaleX(pp1) ', ScaleModeConstants.vbPoints, ScaleModeConstants.vbPixels)      '  1.333333
    Debug.Print SScaleX(pp1) ', ScaleModeConstants.vbPixels, ScaleModeConstants.vbPixels)      '  1
    Debug.Print SScaleX(pp1) ', ScaleModeConstants.vbCharacters, ScaleModeConstants.vbPixels)  '  8
    Debug.Print SScaleX(pp1) ', ScaleModeConstants.vbInches, ScaleModeConstants.vbPixels)      ' 96
    Debug.Print SScaleX(pp1) ', ScaleModeConstants.vbMillimeters, ScaleModeConstants.vbPixels) '  3.779524
    Debug.Print SScaleX(pp1) ', ScaleModeConstants.vbCentimeters, ScaleModeConstants.vbPixels) ' 37.79524
    Debug.Print SScaleX(pp1) ', ScaleModeConstants.vbHimetric, ScaleModeConstants.vbPixels)    '  3.779528E-02
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbContainerPosition, ScaleModeConstants.vbPixels)
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbContainerSize, ScaleModeConstants.vbPixels)
    Debug.Print ""
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbUser, ScaleModeConstants.vbPixels)
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbTwips) ', ScaleModeConstants.vbPixels)       '  6.666667E-02
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbPoints) ', ScaleModeConstants.vbPixels)      '  1.333333
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbPixels) ', ScaleModeConstants.vbPixels)      '  1
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbCharacters) ', ScaleModeConstants.vbPixels)  '  8
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbInches) ', ScaleModeConstants.vbPixels)      ' 96
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbMillimeters) ', ScaleModeConstants.vbPixels) '  3.779524
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbCentimeters) ', ScaleModeConstants.vbPixels) ' 37.79524
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbHimetric) ', ScaleModeConstants.vbPixels)    '  3.779528E-02
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbContainerPosition, ScaleModeConstants.vbPixels)
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbContainerSize, ScaleModeConstants.vbPixels)
    Debug.Print ""
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbUser, ScaleModeConstants.vbPixels)
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)       '  6.666667E-02
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbPoints, ScaleModeConstants.vbPixels)      '  1.333333
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbPixels)      '  1
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbCharacters, ScaleModeConstants.vbPixels)  '  8
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbInches, ScaleModeConstants.vbPixels)      ' 96
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbMillimeters, ScaleModeConstants.vbPixels) '  3.779524
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbCentimeters, ScaleModeConstants.vbPixels) ' 37.79524
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbHimetric, ScaleModeConstants.vbPixels)    '  3.779528E-02
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbContainerPosition, ScaleModeConstants.vbPixels)
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbContainerSize, ScaleModeConstants.vbPixels)
    Debug.Print ""
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbUser, ScaleModeConstants.vbPixels)
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips)       ' 15
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbPoints)      '  0.75
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbPixels)      '  1
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbCharacters)  '  0.125
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbInches)      '  1.041667E-02
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbMillimeters) '  0.2645836
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbCentimeters) '  2.645836E-02
    Debug.Print SScaleX(pp1, ScaleModeConstants.vbPixels, ScaleModeConstants.vbHimetric)    ' 26.45833
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbContainerPosition, ScaleModeConstants.vbPixels)
    'Debug.Print ScaleX(pp1, ScaleModeConstants.vbContainerSize, ScaleModeConstants.vbPixels)
    
    
End Sub
