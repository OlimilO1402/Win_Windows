VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "www.activevb.de"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   3060
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox Check1 
      Caption         =   "Dicker Rahmen"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Value           =   1  'Aktiviert
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Titelleiste"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Value           =   1  'Aktiviert
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Scrollbars"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Schließen Button"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Value           =   1  'Aktiviert
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Min/Max"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Aktiviert
      Width           =   2295
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

Private Declare Function SetWindowLong Lib "user32" Alias _
        "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex _
        As Long, ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias _
        "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex _
        As Long) As Long
        
Private Declare Function SetWindowPos Lib "user32" (ByVal _
        hWnd As Long, ByVal hWndInsertAfter As Long, ByVal _
        x As Long, ByVal y As Long, ByVal cx As Long, _
        ByVal cy As Long, ByVal wFlags As Long) As Long
       
Private Declare Function GetWindowRect Lib "user32" (ByVal _
        hWnd As Long, lpRect As Rect) As Long
        
Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Const SWP_FRAMECHANGED = &H20
Const GWL_STYLE = (-16)
Const WS_MAXIMIZEBOX = &H10000
Const WS_MINIMIZEBOX = &H20000
Const WS_THICKFRAME = &H40000
Const WS_SYSMENU = &H80000
Const WS_HSCROLL = &H100000
Const WS_VSCROLL = &H200000
Const WS_BORDER = &H800000

Private Sub Check1_Click(Index As Integer)
  Select Case Index
    Case 0: Call SetForm(WS_MAXIMIZEBOX, False)
            Call SetForm(WS_MINIMIZEBOX, True)
    Case 1: Call SetForm(WS_SYSMENU, True)
    Case 2: Call SetForm(WS_VSCROLL, True)
            Call SetForm(WS_HSCROLL, True)
    Case 3: Call SetForm(WS_BORDER, True)
    Case 4: Call SetForm(WS_THICKFRAME, True)
  End Select
End Sub

Private Sub SetForm(ToggleStyle&, FRefresh As Boolean)
  Dim lngStyle&, R As Rect
  
    lngStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
    If (lngStyle And ToggleStyle&) Then
      lngStyle = lngStyle - ToggleStyle&
    Else
      lngStyle = lngStyle Or ToggleStyle&
    End If
    Call SetWindowLong(Me.hWnd, GWL_STYLE, lngStyle)
    Call GetWindowRect(Me.hWnd, R)
    Call SetWindowPos(Me.hWnd, 0, R.Left, R.Top, _
         R.Right - R.Left, R.Bottom - R.Top, _
         SWP_FRAMECHANGED)
End Sub
