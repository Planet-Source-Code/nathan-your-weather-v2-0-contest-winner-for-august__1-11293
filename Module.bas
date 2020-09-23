Attribute VB_Name = "Main"
Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wparam As Integer, ByVal iparam As Long) As Long
Global r%
Global entry$
Global iniPath$
Public Sub FormMove(theform As Form)
    ReleaseCapture
    Call SendMessage(theform.hWnd, &HA1, 2, 0&)
End Sub
Function GetFromINI(AppName$, KeyName$, FileName$) As String
    Dim RetStr As String
    RetStr = String(255, Chr(0))
    GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
End Function
Sub systrayme()
    Mainfrm.Hide
    If Systemtrayfrm.SystemTray1.IsIconLoaded = False Then
        Systemtrayfrm.SystemTray1.Icon = Val(Mainfrm.Icon)
        Systemtrayfrm.SystemTray1.SysTrayText = "Your Weather v1.2"
        Systemtrayfrm.SystemTray1.Action = sys_Add
    End If
End Sub

