Imports Microsoft.VisualBasic

Public Class Class1
#Region "Methods And Functions Declaration"


    Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" ( _
             ByVal lpClassName As String, _
             ByVal lpWindowName As String) As Int32

    Public Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" ( _
         ByVal hWnd1 As Int32, _
         ByVal hWnd2 As Int32, _
         ByVal lpsz1 As String, _
         ByVal lpsz2 As String) As Int32

    Public Declare Function SetForegroundWindow Lib "user32.dll" ( _
        ByVal hwnd As Int32) As Int32
#End Region
End Class
