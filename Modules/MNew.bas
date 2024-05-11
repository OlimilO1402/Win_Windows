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


Public Function Window(ByVal Name As String, Optional ByVal Caption As String, Optional ByVal Style As EWndStyle = EWndStyle.VBFormStyle_Sizable, Optional ByVal StyleEx As EWndStyleEx = EWndStyleEx.VBFormStyleEx_Sizable) As Window
    Set Window = New Window: Window.New_ Name, Caption, Style, StyleEx
End Function
