Attribute VB_Name = "Module1"
Option Explicit
 Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, _
        ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function GetModuleHandle Lib "kernel32" Alias _
        "GetModuleHandleA" (ByVal lpModuleName As String) As Long
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
        (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, _
        ByVal dwThreadId As Long) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    Private Declare Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" _
        (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
        (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
    
    Private Const EM_SETPASSWORDCHAR = &HCC
Private Const WH_CBT = 5
Private Const HCBT_ACTIVATE = 5
Private Const HC_ACTION = 0
Private hHook As Long
Public Function NewProc(ByVal lngCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim RetVal
    Dim strClassName As String, lngBuffer As Long
    If lngCode < HC_ACTION Then
        NewProc = CallNextHookEx(hHook, lngCode, wParam, lParam)
        Exit Function
    End If

    strClassName = String$(256, " ")
    lngBuffer = 255
    If lngCode = HCBT_ACTIVATE Then 'A window has been activated
        RetVal = GetClassName(wParam, strClassName, lngBuffer)
        If Left$(strClassName, RetVal) = "#32770" Then
            'This changes the edit control so that it display the password character *.
            'You can change the Asc("*") as you please.
            SendDlgItemMessage wParam, &H1324, EM_SETPASSWORDCHAR, Asc("*"), &H0
        End If
    End If
    'This line will ensure that any other hooks that may be in place are
    'called correctly.
    CallNextHookEx hHook, lngCode, wParam, lParam
End Function

Function InputBoxDK(Prompt, Title) As String
   Dim lngModHwnd As Long
    Dim lngThreadID As Long
    lngThreadID = GetCurrentThreadId
    lngModHwnd = GetModuleHandle(vbNullString)
    hHook = SetWindowsHookEx(WH_CBT, AddressOf NewProc, lngModHwnd, lngThreadID)
    InputBoxDK = InputBox(Prompt, Title)
    UnhookWindowsHookEx hHook
End Function

Public Function UsePasswordForm() As Boolean
101:
   X = InputBoxDK("Enter your Password.", "Password Required")
   If StrPtr(X) = 0 Then
      'Cancel pressed
      UsePasswordForm = False
   Exit Function
   ElseIf X = "" Then
      MsgBox "Please enter a password", vbExclamation, "Alert"
   GoTo 101:
   Else
      'Ok pressed
      'Continue with your macro.
      'Password is stored in the variable "x"
      strsql = "Select password FROM Users Where (islock = 0 or islock is null) and (isAdministrator = 1 or isManager = 1) and password in ('" & EncryptStr(X, True) & "')"
      If CN.Execute(strsql).EOF Then
         MsgBox "Incorrect Password. Please Enter the Correct Paasword", vbExclamation, "Alert"
         GoTo 101:
      Else
         UsePasswordForm = True
      End If
   End If

  
End Function

