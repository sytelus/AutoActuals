Attribute VB_Name = "modGlobal"
Option Explicit
' -------------------------------
' GLOBAL VARIABLES
' -------------------------------
Public gobjErrors As New clsErrHandler          ' Global error handler
' END OF GLOBAL VARIABLES
' -------------------------------

' -------------------------------
' GLOBAL CONSTANTS
' -------------------------------
'App specific
Public Const gsFORM_CAPTION As String = "Data Access"
Public Const gnERROR_RESUME As Integer = 1
Public Const gnERROR_RESUME_NEXT As Integer = 2
Public Const gnERROR_EXIT_APP As Integer = 3
Public Const gnERROR_EXIT As Integer = 3  'included for backward compatibility in old apps
Public Const gnERROR_EXIT_PROC As Integer = 4
Public Const gnRESUME_NEXT As Integer = 2        ' for compatibilty with other code
Public Const gnCANCEL As Integer = 4             ' for compatibilily with other code
Public Const gnGO_GLOBAL_HANDLER As Integer = 5
Public Const gnERR_OBJECT_DOES_NOT_SUPPORT_THIS_PROPERTY = 438


Public Const gnRULE_TRUE As Integer = 1
Public Const gnRULE_FALSE As Integer = 0
Public Const gnRULE_EMPTY As Integer = 2


'------------------------------------------
'API Declarations and constants
'------------------------------------------
Public Const MAX_PATH = 260
Public Const SW_SHOW = 5
Public Const SW_NORMAL = 1
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SWP_HIDEWINDOW = &H80
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const GW_HWNDNEXT = 2
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_SETDROPPEDWIDTH As Long = 352
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const BIF_RETURNONLYFSDIRS = 1

Public Type BrowseInfo
    hWndOwner       As Long
    pIDLRoot        As Long
    pszDisplayName  As String
    lpszTitle       As String
    ulFlags         As Long
    lpfnCallBack    As Long
    lparam          As Long
    iImage           As Long
End Type






Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As Any, ByVal lpsz2 As Any) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
'Do not change type of lParam to Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lparam As Any) As Long
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long



