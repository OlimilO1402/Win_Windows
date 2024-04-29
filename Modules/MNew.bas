Attribute VB_Name = "MNew"
Option Explicit
'
'#If VBA7 = 0 Then
'    Public Enum LongPtr
'        [_]
'    End Enum
'#End If
'
'Public Type WindowRect
'    Left   As Long
'    Top    As Long
'    Width  As Long
'    Height As Long
'End Type


Public Function Window(ByVal Caption As String) As Window
    Set Window = New Window: Window.New_ Caption
End Function
