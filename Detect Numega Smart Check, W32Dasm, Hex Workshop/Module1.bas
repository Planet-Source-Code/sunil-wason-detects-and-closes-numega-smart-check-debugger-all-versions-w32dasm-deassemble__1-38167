Attribute VB_Name = "Module1"
Option Explicit

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function sendmessagebystring Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function getwindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)

Public Const WM_CLOSE = &H10
Public Const GW_HWNDNEXT = 2
Public Const GW_OWNER = 4
Public Const EXCEPTION_ACCESS_VIOLATION = &HC0000005

'*-------------------------------------
'Stores the truncated class names of
'applications which are to be monitored in the
'system while this application is running
Public Const NoOfAppClassMonitored As Integer = 4
Public AppClassName(NoOfAppClassMonitored) As String

Public Sub FillClassName()

    Dim i As Integer
    'Clear the Array
    For i = 0 To NoOfAppClassMonitored
        AppClassName(i) = ""
    Next i
    AppClassName(0) = "HexWorks" 'Hex Workshop 3.1
    AppClassName(1) = "OWL_Window" 'WinDasm
    AppClassName(2) = "NMSCMW" 'for various versions of Numega Smart Check
    AppClassName(3) = "Winamp" 'WinAmp Playlist, WinAmp Mnibrowser, WinAmp version x.x, Winamp Equaliser
    AppClassName(4) = "Notepad" 'Notepad

End Sub 'FillClassName()

Public Function AppPresent(TgtClassName As String, frmName As Object) As Long

'Check if the windows classname matches with
'our list in the array (TgtClassName i.e. target
'classname)

AppPresent = SearchAppsByClassName(TgtClassName, frmName)

End Function 'AppPresent(TgtClassName As String, frmName As Object) As Long

Function SearchAppsByClassName(TgtClassName As String, frmName As Object)
    'Searches all the applications running in
    'Windows environment either in the
    'foreground or in the background
    Dim ThisAppHandle As Long
    Dim NextAppHandle As Long
    Dim AppClassName As String
    
    'Get your current application owner's
    'handle
    ThisAppHandle = getwindow(frmName.hwnd, GW_OWNER)
    'Pass it on as a seed for the next handle
    'so all the windows in the z order can be
    'searched
    NextAppHandle = ThisAppHandle
    'Perform the iterations in the Z order till
    'all the running applications (either foreground
    'or background) have been run through
    Do While NextAppHandle <> 0
        DoEvents
        'Get the handle of the next window
        'in the Z order
        NextAppHandle = getwindow(NextAppHandle, GW_HWNDNEXT)
        'Retrieve its window's classname
        AppClassName = GetAppClassName(NextAppHandle)
        'If a part of TgtClassName is found
        'then retrieve this applications handle
        If InStr(1, LCase(Trim(AppClassName)), LCase(Trim(TgtClassName))) <> 0 Then
            SearchAppsByClassName = NextAppHandle
            Exit Do
            Exit Function
        End If
    Loop
    
End Function 'SearchAppsByClassName(TgtClassName As String, frmName As Object)

Private Function GetAppClassName(HwndWindow As Long) As String

Dim BufferAll&
Dim WindowClassName$
    'Get the Window's Class Name
    WindowClassName$ = String(100, Chr(0))
    BufferAll& = GetClassName(HwndWindow, WindowClassName$, 100)
    GetAppClassName = Left(WindowClassName$, BufferAll&)
DoEvents

End Function 'GetAppClassName(HwndWindow As Long) As String

Public Sub KillWin(hwnd As Long)

    'Close the selected application
    Dim dummy
    dummy = sendmessagebystring(hwnd, WM_CLOSE, 0, 0)
    If frmCheck.chkFound = 1 Then
        'Raise an exception
        RaiseException EXCEPTION_ACCESS_VIOLATION, 0, 0, 0
    End If
    
End Sub 'KillWin(hwnd As Long)

