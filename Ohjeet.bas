Attribute VB_Name = "Ohjeet"
Option Explicit

Private Declare PtrSafe Function ShellExecute _
Lib "shell32.dll" Alias "ShellExecuteA" ( _
ByVal hWnd As Long, _
ByVal Operation As String, _
ByVal Filename As String, _
Optional ByVal Parameters As String, _
Optional ByVal Directory As String, _
Optional ByVal WindowStyle As Long = vbMinimizedFocus _
) As Long

Public Sub OpenUrl()

    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", "https://github.com/takutta/Hoitoajat-Excel")

End Sub


