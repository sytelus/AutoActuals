Attribute VB_Name = "modSysTray"
Option Explicit

'Author:
'        Ben Baird <psyborg@cyberhighway.com>
'        Copyright (c) 1997, Ben Baird
'
'Purpose:
'        Demonstrates setting an icon in the taskbar's
'        system tray without the overhead of subclassing
'        to receive events.

Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Const WM_MOUSEMOVE = &H200
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const MAX_TOOLTIP As Integer = 64

Type NOTIFYICONDATA
 cbSize As Long
 hwnd As Long
 uID As Long
 uFlags As Long
 uCallbackMessage As Long
 hIcon As Long
 szTip As String * MAX_TOOLTIP
End Type

Public nfIconData As NOTIFYICONDATA

Public Sub AddIconInSysTray(frm As Form, ByVal vsHint As String)

    'Add the icon to the system tray...
    With nfIconData
     .hwnd = frm.hwnd
     .uID = frm.Icon
     .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
     .uCallbackMessage = WM_MOUSEMOVE
     .hIcon = frm.Icon.Handle
     .szTip = vsHint & Chr$(0)
     .cbSize = Len(nfIconData)
    End With
    
    Call Shell_NotifyIcon(NIM_ADD, nfIconData)
    
End Sub

Public Sub DeleteIconFromSysTray()
    
    Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
    
End Sub

Public Sub ModifySysTrayIcon(Optional frm As Variant, Optional ByVal vsHint As Variant)

    If Not IsMissing(frm) Then
    
        nfIconData.hIcon = frm.Icon.Handle
        
    End If
    
    If Not IsMissing(vsHint) Then
    
        nfIconData.szTip = vsHint & Chr$(0)
        
    End If

    Call Shell_NotifyIcon(NIM_MODIFY, nfIconData)

End Sub

