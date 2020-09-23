Attribute VB_Name = "ModMsgBox"
'This module is a heavily modified version
'of something I downloaded ages ago.
'Apologies to the author for forgetting your name
'It places a temporary window hook on a
'specific Msgbox, and then automatically
'removes the hook after use.
Option Explicit
Private Const MB_YESNOCANCEL = &H3&
Private Const MB_YESNO = &H4&
Private Const MB_RETRYCANCEL = &H5&
Private Const MB_OKCANCEL = &H1&
Private Const MB_OK = &H0&
Private Const MB_ABORTRETRYIGNORE = &H2&
Private Const MB_ICONEXCLAMATION = &H30&
Private Const MB_ICONQUESTION = &H20&
Private Const MB_ICONASTERISK = &H40&
Private Const MB_ICONINFORMATION = MB_ICONASTERISK
Private Const IDOK = 1
Private Const IDCANCEL = 2
Private Const IDABORT = 3
Private Const IDRETRY = 4
Private Const IDIGNORE = 5
Private Const IDYES = 6
Private Const IDNO = 7
Private Const IDPROMPT = &HFFFF&
Private Const WH_CBT = 5
Private Const GWL_HINSTANCE = (-6)
Private Const HCBT_ACTIVATE = 5
Private Type MSGBOX_HOOK_PARAMS
   hwndOwner   As Long
   hHook       As Long
End Type
Private MSGHOOK As MSGBOX_HOOK_PARAMS
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Dim mbFlags As VbMsgBoxStyle
Dim mbFlags2 As VbMsgBoxStyle
Dim mTitle As String
Dim mPrompt As String
Dim But1 As String
Dim But2 As String
Dim But3 As String
Public Function MessageBoxH(hwndThreadOwner As Long, hwndOwner As Long, mbFlags As VbMsgBoxStyle) As Long
'This function calls the hook
Dim hInstance As Long
Dim hThreadId As Long
hInstance = GetWindowLong(hwndThreadOwner, GWL_HINSTANCE)
hThreadId = GetCurrentThreadId()
With MSGHOOK
   .hwndOwner = hwndOwner
   .hHook = SetWindowsHookEx(WH_CBT, AddressOf MsgBoxHookProc, hInstance, hThreadId)
End With
MessageBoxH = MessageBox(hwndOwner, Space$(120), Space$(120), mbFlags)
End Function
Public Function MsgBoxHookProc(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'This function catches the messagebox before it opens
'and changes the text of the buttons - then removes the hook
If uMsg = HCBT_ACTIVATE Then
SetWindowText wParam, mTitle
SetDlgItemText wParam, IDPROMPT, mPrompt
Select Case mbFlags
Case vbAbortRetryIgnore
   SetDlgItemText wParam, IDABORT, But1
   SetDlgItemText wParam, IDRETRY, But2
   SetDlgItemText wParam, IDIGNORE, But3
 Case vbYesNoCancel
   SetDlgItemText wParam, IDYES, But1
   SetDlgItemText wParam, IDNO, But2
   SetDlgItemText wParam, IDCANCEL, But3
 Case vbOKOnly
   SetDlgItemText wParam, IDOK, But1
 Case vbRetryCancel
   SetDlgItemText wParam, IDRETRY, But1
   SetDlgItemText wParam, IDCANCEL, But2
 Case vbYesNo
   SetDlgItemText wParam, IDYES, But1
   SetDlgItemText wParam, IDNO, But2
 Case vbOKCancel
   SetDlgItemText wParam, IDOK, But1
   SetDlgItemText wParam, IDCANCEL, But2
End Select
   UnhookWindowsHookEx MSGHOOK.hHook
End If
MsgBoxHookProc = False
End Function


Public Function BBmsgbox(mhwnd As Long, mMsgbox As VbMsgBoxStyle, Title As String, Prompt As String, Optional mMsgIcon As VbMsgBoxStyle, Optional ButA As String, Optional ButB As String, Optional ButC As String) As String
'This function sets your custom parameters and returns
'which button was pressed as a string
Dim mReturn As Long
mbFlags = mMsgbox
mbFlags2 = mMsgIcon
mTitle = Title
mPrompt = Prompt
But1 = ButA
But2 = ButB
But3 = ButC
mReturn = MessageBoxH(mhwnd, GetDesktopWindow(), mbFlags Or mbFlags2)
Select Case mReturn
    Case IDABORT
        BBmsgbox = But1
    Case IDRETRY
        BBmsgbox = But2
    Case IDIGNORE
        BBmsgbox = But3
    Case IDYES
        BBmsgbox = But1
    Case IDNO
        BBmsgbox = But2
    Case IDCANCEL
        BBmsgbox = But3
    Case IDOK
        BBmsgbox = But1
End Select
End Function
