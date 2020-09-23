Attribute VB_Name = "Module1"
Option Explicit
Public Const MB_OK = &H0&
Public Const MB_OKCANCEL = &H1&
Public Const MB_ABORTRETRYIGNORE = &H2&
Public Const MB_YESNOCANCEL = &H3&
Public Const MB_YESNO = &H4&
Public Const MB_RETRYCANCEL = &H5&
Public Const MB_ICONMASK = &HF0&
Public Const MB_ICONCRITICAL = &H10&
Public Const MB_ICONQUESTION = &H20&
Public Const MB_ICONINFORMATION = &H40&
Public Const MB_ICONEXCLAMATION = &H30&
Public Const MB_SYSTEMMODAL = &H1000&
Public Const MB_TASKMODAL = &H2000&
Public Const IDOK = 1
Public Const IDCANCEL = 2
Public Const IDABORT = 3
Public Const IDRETRY = 4
Public Const IDIGNORE = 5
Public Const IDYES = 6
Public Const IDNO = 7
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const FLAGS = SWP_NOSIZE Or SWP_NOMOVE

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function MessageBox Lib "user32" Alias _
"MessageBoxA" (ByVal hwnd As Long, ByVal lpText As _
String, ByVal lpCaption As String, ByVal wType As Long) _
As Long
Public Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal _
X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As _
Long, ByVal wFlags As Long) As Long
Public Function FormOnTop(frm As Form)
Dim SetFrmOnTop As Long
    SetFrmOnTop = SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, _
                    0, 0, FLAGS)
End Function

Public Function FormNotOnTop(frm As Form)
Dim SetFrmNotOnTop As Long
    SetFrmNotOnTop = SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0, _
                        0, 0, 0, FLAGS)
End Function



