Attribute VB_Name = "Phrost32"
Option Explicit
''to be tested''



''confirmed''
'aol://1391:
'aol://9293:
'aol://1723:
'aol://3548:
'aol://4344:613.hoch4962.3505268.552257542
'aol://4401:
'aol://5862:144/members.aol.com:/stevecase
'aol://4344:773.HOTNIP1.6843825.521317437
'aol://4950:
'aol://8143:

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG As Long = &H80000005
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const OPEN_EXISTING = 3
Public Const OPEN_ALWAYS = 4
Public Const FILE_CURRENT = 1
Public Const FILE_BEGIN = 0
Public Const SRCCOPY = &HCC0020
Public Const CB_GETCOUNT = &H146
Public Const CB_GETITEMDATA = &H150
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const WM_MOUSEMOVE = &H200
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT
Public Const SW_Hide = 0
Public Const SW_SHOW = 5
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26
Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const RGN_DIFF = 4
Public Const SC_CLICKMOVE = &HF012&

Declare Function GetDriveType& Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String)
Declare Function GetLogicalDrives& Lib "kernel32" ()
Declare Function CreateFile& Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long)
Declare Function ReadFile& Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long)
Declare Function WriteFile& Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Long)
Declare Function CloseHandle& Lib "kernel32" (ByVal hObject As Long)
Declare Function SetFilePointer& Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByVal lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long)
Declare Function lwrite& Lib "kernel32" Alias "_lwrite" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal wBytes As Long)
Declare Function lread& Lib "kernel32" Alias "_lread" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal wBytes As Long)
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hWndCallback As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function GetCapture Lib "user32" () As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwReserved As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName$, ByVal lpdwReserved As Long, lpdwType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String)
Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String)
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)
Declare Function SendMessageByString& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam$)
Declare Function SetParent& Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long)
Declare Function GetParent& Lib "user32" (ByVal hwnd As Long)
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function HideCaret& Lib "user32" (ByVal hwnd As Long)
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Type RGB
    R As Integer
    G As Integer
    b As Integer
End Type

Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uId As Long
        uFlags As Long
        ucallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public Type POINTAPI
        x As Long
        y As Long
End Type
Private rp As Boolean
Private gp As Boolean
Private bp As Boolean
Private CurRgn, TempRgn As Long
Public Sub FormDrag(TheForm As Form)
ReleaseCapture
Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub
Function ReadINI(AppName, KeyName, FileName As String) As String
Dim ret As String
ret = String(255, Chr(0))
ReadINI = Left(ret, GetPrivateProfileString(AppName, ByVal KeyName, "", ret, Len(ret), FileName))
End Function
Function WriteINI(sAppname As String, sKeyName As String, sNewString, sFileName As String)
WriteINI = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)
End Function
Sub stayontop(myfrm As Form, SetOnTop As Boolean)
Dim lFlag As Long
If SetOnTop Then
    lFlag = HWND_TOPMOST
Else
    lFlag = HWND_NOTOPMOST
End If
SetWindowPos myfrm.hwnd, lFlag, myfrm.Left / Screen.TwipsPerPixelX, myfrm.Top / Screen.TwipsPerPixelY, myfrm.Width / Screen.TwipsPerPixelX, myfrm.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub
Function ReplaceString(text As String, what As String, WithWhat As String, Optional All As Boolean) As String
Dim pos As Integer
Dim temp As String, temp2 As String, temp3 As String
pos = InStr(text, what)
If pos = 0 Then ReplaceString = text: Exit Function
temp = text

If All = True Then
    Do: DoEvents
        pos = InStr(temp, what)
        If pos = 0 Then Exit Do
        temp2 = Left(temp, pos - 1)
        temp3 = Mid(temp, pos + Len(what))
        temp = temp2 & WithWhat & temp3
    Loop
Else
    pos = InStr(temp, what)
    temp2 = Left(temp, pos - 1)
    temp3 = Mid(temp, pos + Len(what))
    temp = temp2 & WithWhat & temp3
End If

ReplaceString = temp
End Function
Sub DisableACT()
Call SystemParametersInfo(97, True, 0&, 0)
End Sub
Sub EnableACD()
Call SystemParametersInfo(97, False, 0&, 0)
End Sub
Function RoomName()
RoomName = ReplaceString(GetCaption(FindRoom), " ", "", True)
End Function
Function aolwindow()
Dim aol%
aol% = FindWindow("AOL Frame25", vbNullString)
aolwindow = aol%
End Function
Function MDIWindow() As Long
Dim MDI%
MDI% = FindChildByClass(aolwindow, "MDIClient")
MDIWindow = MDI%
End Function
Sub SetFocus()
Dim x
x = GetCaption(aolwindow)
AppActivate x
End Sub
Function Upchat()
Dim die%, x
die% = FindWindow("_AOL_MODAL", vbNullString)
x = ShowWindow(die%, SW_Hide)
x = ShowWindow(die%, SW_MINIMIZE)
Call SetFocus
End Function
Sub UnUpchat()
Dim die%, x
die% = FindWindow("_AOL_MODAL", vbNullString)
x = ShowWindow(die%, SW_RESTORE)
Call SetFocus
End Sub
Function FindChildByClass(parentHwnd, childhand)
Dim ReturnString$, handles&, Parent, Copy
Copy = parentHwnd
Parent = GetWindow(parentHwnd, 5)
Top: ReturnString$ = String$(250, 0): handles& = GetClassName(Parent, ReturnString$, 250)
If Left$(ReturnString$, Len(childhand)) Like childhand Then GoTo ending:
Parent = GetWindow(Parent, 2)
If Parent > 0 Then GoTo Top

ending:
FindChildByClass = Parent
parentHwnd = Copy
End Function
Function Find2NdChildByClass(parentHwnd, childhand)
Dim ReturnString$, handles&, Parent, Copy, foundfirst As Boolean
foundfirst = False
Copy = parentHwnd
Parent = GetWindow(parentHwnd, 5)

Top:
ReturnString$ = String$(250, 0)
handles& = GetClassName(Parent, ReturnString$, 250)
If Left$(ReturnString$, Len(childhand)) Like childhand Then
    If foundfirst = True Then
        GoTo ending:
    Else
        foundfirst = True
    End If
End If
Parent = GetWindow(Parent, 2)
If Parent > 0 Then GoTo Top

ending:
Find2NdChildByClass = Parent
parentHwnd = Copy
End Function
Public Function FindChildByTitle(ParentWindow As Long, WindowTxt As String) As Long
FindChildByTitle& = FindWindowEx(ParentWindow&, 0&, vbNullString, WindowTxt$)
End Function
Sub Click(Button%)
Dim sendnow%
sendnow% = SendMessageByNum(Button%, WM_LBUTTONDOWN, &HD, 0)
sendnow% = SendMessageByNum(Button%, WM_LBUTTONUP, &HD, 0)
End Sub
Sub ChatClear()
Send ("<font face=" + "symbol" + ">ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ<font face=" + "arial" + ">")
End Sub
Public Function SendMail(Person As String, Subject As String, message As String) As String
Dim aol As Long, MDI As Long, tool As Long, Toolbar As Long
Dim ToolIcon As Long, OpenSend As Long, DoIt As Long
Dim Rich As Long, EditTo As Long, EditCC As Long
Dim EditSubject As Long, SendButton As Long
Dim Combo As Long, fCombo As Long, ErrorWindow As Long
Dim Button1 As Long, Button2 As Long, error As Long
Dim view As Long, text As String, btn As Long, modal As Long
Dim FullWindow As Long, NoButton As Long
aol& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
Call SendMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
DoEvents
Do
    DoEvents
    OpenSend& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
    EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
    EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
    EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
    Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
    Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
    fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
    Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
    Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
    SendButton& = FindChildByClassEx(OpenSend&, "_AOL_Icon", 16)
Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
Call SendMessageByString(EditTo&, WM_SETTEXT, 0, Person$)
DoEvents
Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, Subject$)
DoEvents
Call SendMessageByString(Rich&, WM_SETTEXT, 0, message$)
DoEvents
Do: DoEvents
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    OpenSend& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
    error& = FindChildByTitle(MDIWindow, "Error")
    If error& <> 0 Then GoTo error1
Loop While OpenSend& <> 0
Do: DoEvents
    modal& = FindWindow("_AOL_Modal", vbNullString)
    btn& = FindChildByClass(modal&, "_AOL_Icon")
    icon (btn&)
Loop While modal& <> 0 And btn& <> 0
SendMail = "Mail Sent": Exit Function
error1:
error& = FindChildByTitle(MDIWindow, "Error")
view& = FindChildByClass(error&, "_AOL_View")
text$ = GetText(view&)
text$ = Mid(text$, InStr(text, Chr(13) & Chr(10) & Chr(13) & Chr(10)) + 4)
Do: DoEvents
btn& = FindChildByClass(error&, "_AOL_Icon")
Loop While btn& = 0
Do: DoEvents
btn& = FindChildByClass(error&, "_AOL_Icon")
Call icon(btn&)
Loop While btn& <> 0
SendMail = text$
Send (SendMail)
End Function
Public Function SendMailEx(lst As ListBox, Subject As String, message As String) As String
Dim aol As Long, MDI As Long, tool As Long, Toolbar As Long
Dim ToolIcon As Long, OpenSend As Long, DoIt As Long
Dim Rich As Long, EditTo As Long, EditCC As Long
Dim EditSubject As Long, SendButton As Long
Dim Combo As Long, fCombo As Long, ErrorWindow As Long
Dim Button1 As Long, Button2 As Long, error As Long
Dim view As Long, text As String, btn As Long, modal As Long
Dim FullWindow As Long, NoButton As Long, i
aol& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
Call SendMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
DoEvents
Do
    DoEvents
    OpenSend& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
    EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
    EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
    EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
    Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
    Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
    fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
    Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
    Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
    SendButton& = FindChildByClassEx(OpenSend&, "_AOL_Icon", 16)
    If OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2& Then GoTo begin
Loop
begin:
Call SendMessageByString(EditTo&, WM_SETTEXT, 0, "")
Call SendMessageByString(EditTo&, WM_SETTEXT, 0, ListToString(lst, False))
DoEvents
Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, Subject$)
DoEvents
Call SendMessageByString(Rich&, WM_SETTEXT, 0, message$)
DoEvents
Do: DoEvents
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    OpenSend& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
    error& = FindChildByTitle(MDIWindow, "Error")
    If error& <> 0 Then GoTo error1
Loop While OpenSend& <> 0
Do: DoEvents
    modal& = FindWindow("_AOL_Modal", vbNullString)
    btn& = FindChildByClass(modal&, "_AOL_Icon")
    icon (btn&)
Loop While modal& <> 0 And btn& <> 0
SendMailEx = "Mail Sent": Exit Function
error1:

Do: DoEvents
error& = FindChildByTitle(MDIWindow, "Error")
view& = FindChildByClass(error&, "_AOL_View")
text$ = GetText(view&)
Loop Until Len(text) > 0
text$ = Mid(text$, InStr(text, Chr(13) & Chr(10) & Chr(13) & Chr(10)) + 4)
text$ = Mid(text$, 1, InStr(text$, " - ") - 1)
For i = 0 To lst.ListCount - 1
If LCase(lst.List(i)) = LCase(text$) Then lst.RemoveItem i
Next i
Do: DoEvents
error& = FindChildByTitle(MDIWindow, "Error")
btn& = FindChildByClass(error&, "_AOL_Icon")
Loop While btn& = 0
Do: DoEvents
error& = FindChildByTitle(MDIWindow, "Error")
btn& = FindChildByClass(error&, "_AOL_Icon")
Call icon(btn&)
Loop While btn& <> 0
GoTo begin:
End Function
Public Sub Send(Chat As String)
Dim temp As String
If Chat = "" Then Exit Sub
Dim Room As Long, AORich As Long, AORich2 As Long
Room& = FindRoom&
AORich& = FindWindowEx(Room, 0&, "RICHCNTL", vbNullString)
AORich2& = FindWindowEx(Room, AORich, "RICHCNTL", vbNullString)
temp = GetText(AORich2&)
Call SendMessageByString(AORich2, WM_SETTEXT, 0&, "")
Call SendMessageByString(AORich2, WM_SETTEXT, 0&, Chat$)
Do: DoEvents
    Call SendMessageLong(AORich2, WM_CHAR, ENTER_KEY, 0&)
Loop While Len(GetText(AORich2&)) > 0
Call SendMessageByString(AORich2, WM_SETTEXT, 0&, temp)
End Sub
Public Function FindRoom() As Long
Dim aol As Long, MDI As Long, child As Long
Dim Rich As Long, AOLList As Long
Dim aolicon As Long, AOLStatic As Long
aol& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
aolicon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
If Rich& <> 0& And AOLList& <> 0& And aolicon& <> 0& And AOLStatic& <> 0& Then
    FindRoom& = child&
    Exit Function
Else
    Do
        child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
        Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
        AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
        aolicon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
        AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
        If Rich& <> 0& And AOLList& <> 0& And aolicon& <> 0& And AOLStatic& <> 0& Then
            FindRoom& = child&
            Exit Function
        End If
    Loop Until child& = 0&
End If
FindRoom& = child&
End Function
Public Function RoomCount() As Long
Dim aol As Long, MDI As Long, rMail As Long, rList As Long
Dim Count As Long
aol& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
rMail& = FindRoom
rList& = FindWindowEx(rMail&, 0&, "_AOL_Listbox", vbNullString)
Count& = SendMessage(rList&, LB_GETCOUNT, 0&, 0&)
RoomCount& = Count&
End Function
Public Function ChatIgnoreByIndex(index As Long, x As Integer) As Boolean
Dim Room As Long, sList As Long, iWindow As Long
Dim iCheck As Long, a As Long, Count As Long, XX As Integer
1 XX = 0
Count& = RoomCount&
If index& > Count& - 1 Then Exit Function
Room& = FindRoom&
sList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
Call SendMessage(sList&, LB_SETCURSEL, index&, 0&)
Call PostMessage(sList&, WM_LBUTTONDBLCLK, 0&, 0&)
Do
    DoEvents
    iWindow& = FindInfoWindow
Loop Until iWindow& <> 0&
DoEvents
iCheck& = FindWindowEx(iWindow&, 0&, "_AOL_Checkbox", vbNullString)
DoEvents
Do
    DoEvents
    XX = XX + 1
    a& = SendMessage(iCheck&, BM_GETCHECK, 0&, 0&)
    If x <> a& Then
        Call SendMessage(iCheck&, WM_LBUTTONDOWN, 0&, 0&)
        DoEvents
        Call SendMessage(iCheck&, WM_LBUTTONUP, 0&, 0&)
        DoEvents
        ChatIgnoreByIndex = True
        Exit Do
    Else
        ChatIgnoreByIndex = False
        Exit Do
    End If
    If XX > 5000 Then GoTo 1
Loop Until a& <> 0&
DoEvents
Call PostMessage(iWindow&, WM_CLOSE, 0&, 0&)
End Function
Public Function Ignore(name As String) As Boolean
On Error Resume Next
Dim cProcess As Long, itmHold As Long, screenname As String
Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
Dim rList As Long, sThread As Long, mThread As Long
Dim lIndex As Long
Room& = FindRoom&
If Room& = 0& Then Exit Function
rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
sThread& = GetWindowThreadProcessId(rList, cProcess&)
mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
If mThread& Then
    For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
        screenname$ = String$(4, vbNullChar)
        itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
        itmHold& = itmHold& + 24
        Call ReadProcessMemory(mThread&, itmHold&, screenname$, 4, rBytes)
        Call CopyMemory(psnHold&, ByVal screenname$, 4)
        psnHold& = psnHold& + 6
        screenname$ = String$(16, vbNullChar)
        Call ReadProcessMemory(mThread&, psnHold&, screenname$, Len(screenname$), rBytes&)
        screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
        If screenname$ <> GetUser$ And LCase(screenname$) = LCase(ChatName(name$)) Then
            lIndex& = index&
            Ignore = ChatIgnoreByIndex(lIndex&, 1)
            DoEvents
            Exit Function
        End If
    Next index&
    Call CloseHandle(mThread)
End If
End Function
Public Function UnIgnore(name As String) As Boolean
On Error Resume Next
Dim cProcess As Long, itmHold As Long, screenname As String
Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
Dim rList As Long, sThread As Long, mThread As Long
Dim lIndex As Long
Room& = FindRoom&
If Room& = 0& Then Exit Function
rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
sThread& = GetWindowThreadProcessId(rList, cProcess&)
mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
If mThread& Then
    For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
        screenname$ = String$(4, vbNullChar)
        itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
        itmHold& = itmHold& + 24
        Call ReadProcessMemory(mThread&, itmHold&, screenname$, 4, rBytes)
        Call CopyMemory(psnHold&, ByVal screenname$, 4)
        psnHold& = psnHold& + 6
        screenname$ = String$(16, vbNullChar)
        Call ReadProcessMemory(mThread&, psnHold&, screenname$, Len(screenname$), rBytes&)
        screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
        If screenname$ <> GetUser$ And LCase(screenname$) = LCase(ChatName(name$)) Then
            lIndex& = index&
            UnIgnore = ChatIgnoreByIndex(lIndex&, 0)
            DoEvents
            Exit Function
        End If
    Next index&
    Call CloseHandle(mThread)
End If
End Function
Public Sub WaitForOKOrRoom(Room As String)
Dim RoomTitle As String, FullWindow As Long, FullButton As Long
Room$ = LCase(ReplaceString(Room$, " ", "", True))
Do
    DoEvents
    RoomTitle$ = GetCaption(FindRoom&)
    RoomTitle$ = LCase(ReplaceString(RoomTitle$, " ", "", True))
    FullWindow& = FindWindow("#32770", "America Online")
    FullButton& = FindWindowEx(FullWindow&, 0&, "Button", "OK")
Loop Until (FullWindow& <> 0& And FullButton& <> 0&) Or Room$ = RoomTitle$
DoEvents
If FullWindow& <> 0& Then
    Do
        DoEvents
        Call SendMessage(FullButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(FullButton&, WM_KEYUP, VK_SPACE, 0&)
        Call SendMessage(FullButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(FullButton&, WM_KEYUP, VK_SPACE, 0&)
        FullWindow& = FindWindow("#32770", "America Online")
        FullButton& = FindWindowEx(FullWindow&, 0&, "Button", "OK")
    Loop Until FullWindow& = 0& And FullButton& = 0&
End If
DoEvents
End Sub
Public Sub PR(RoomName As String, Optional Restricted As Boolean)
Dim Room As String, i As Integer
If Restricted = False Then
    Room$ = "aol://2719:2-2-" & RoomName
Else
    Room$ = "aol://2719:2-2-"
    For i = 1 To Len(RoomName$)
        Room$ = Room$ & Mid(RoomName$, i, 1) & "%20"
    Next i
    Room$ = Left(Room$, Len(Room$) - 3)
End If
Call Keyword(Room$)
End Sub
Public Sub IM(Person As String, message As String)
Dim aol As Long, MDI As Long, IM As Long, Rich As Long
Dim SendButton As Long, OK As Long, Button As Long
aol& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
Call Keyword("aol://9293:" & Person$)
Do
    DoEvents
    IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
    SendButton& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
    SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
    SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
    SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
    SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
    SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
    SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
    SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
    SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
Loop Until IM& <> 0& And Rich& <> 0& And SendButton& <> 0&
Call SendMessageByString(Rich&, WM_SETTEXT, 0&, message$)
Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
Do
    DoEvents
    OK& = FindWindow("#32770", "America Online")
    IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
Loop Until OK& <> 0& Or IM& = 0&
If OK& <> 0& Then
    Button& = FindWindowEx(OK&, 0&, "Button", vbNullString)
    Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
    Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
    Call PostMessage(IM&, WM_CLOSE, 0&, 0&)
End If
End Sub
Public Sub IMIgnore(Person As String)
Call IM("$IM_OFF, " & Person$, "phrost32.bas")
End Sub
Public Sub IMUnIgnore(Person As String)
Call IM("$IM_ON, " & Person$, "phrost32.bas")
End Sub
Public Sub IMsOff()
Call IM("$IM_OFF", "phrost32.bas")
End Sub
Public Sub IMsOn()
Call IM("$IM_ON", "phrost32.bas")
End Sub
Public Function IMSender() As String
Dim IM As Long, Caption As String
Caption$ = GetCaption(FindIM&)
If InStr(Caption$, ":") = 0& Then
    IMSender$ = ""
    Exit Function
Else
    IMSender$ = Right(Caption$, Len(Caption$) - InStr(Caption$, ":") - 1)
End If
End Function
Public Function IMText() As String
Dim Rich As Long
Rich& = FindWindowEx(FindIM&, 0&, "RICHCNTL", vbNullString)
IMText$ = GetText(Rich&)
End Function
Public Function IMLastMsg() As String
Dim Rich As Long, MsgString As String, Spot As Long
Dim NewSpot As Long
Rich& = FindWindowEx(FindIM&, 0&, "RICHCNTL", vbNullString)
MsgString$ = GetText(Rich&)
NewSpot& = InStr(MsgString$, Chr(9))
Do
    Spot& = NewSpot&
    NewSpot& = InStr(Spot& + 1, MsgString$, Chr(9))
Loop Until NewSpot& <= 0&
MsgString$ = Right(MsgString$, Len(MsgString$) - Spot& - 1)
IMLastMsg$ = Left(MsgString$, Len(MsgString$) - 1)
End Function
Public Sub Keyword(Keyword As String)
Dim aol As Long, tool As Long, Toolbar As Long
Dim Combo As Long, EditWin As Long
aol& = FindWindow("AOL Frame25", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
Combo& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
Call SendMessageByString(EditWin&, WM_SETTEXT, 0&, Keyword$)
Call SendMessageLong(EditWin&, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(EditWin&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Function GetCaption(WindowHandle As Long) As String
Dim Buffer As String, TextLength As Long
TextLength& = GetWindowTextLength(WindowHandle&)
Buffer$ = String(TextLength&, 0&)
Call GetWindowText(WindowHandle&, Buffer$, TextLength& + 1)
GetCaption$ = Buffer$
End Function
Public Function GetText(WindowHandle As Long) As String
Dim Buffer As String, TextLength As Long
TextLength& = SendMessage(WindowHandle&, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = String(TextLength&, 0&)
Call SendMessageByString(WindowHandle&, WM_GETTEXT, TextLength& + 1, Buffer$)
GetText$ = Buffer$
End Function
Public Sub Button(mButton As Long)
Call SendMessage(mButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(mButton&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub icon(aIcon As Long)
Call SendMessage(aIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(aIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub CloseWindow(Window As Long)
Call SendMessage(Window&, WM_CLOSE, 0&, 0&)
End Sub
Public Function GetUser() As String
Dim aol As Long, MDI As Long, welcome As Long
Dim child As Long, UserString As String
aol& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
UserString$ = GetCaption(child&)
If InStr(UserString$, "Welcome, ") = 1 Then
    UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
    GetUser$ = UserString$
    Exit Function
Else
    Do
        child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
        UserString$ = GetCaption(child&)
        If InStr(UserString$, "Welcome, ") = 1 Then
            UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
            GetUser$ = UserString$
            Exit Function
        End If
    Loop Until child& = 0&
End If
GetUser$ = ""
End Function
Public Sub Pause(duration As Long)
Dim Current As Long
Current = Timer
Do Until Timer - Current >= duration
    DoEvents
Loop
End Sub
Public Sub PlayMidi(MIDIFile As String)
Dim SafeFile As String
SafeFile$ = Dir(MIDIFile$)
If SafeFile$ <> "" Then
    Call mciSendString("play " & MIDIFile$, 0&, 0, 0)
End If
End Sub
Public Sub StopMIDI(MIDIFile As String)
Dim SafeFile As String
SafeFile$ = Dir(MIDIFile$)
If SafeFile$ <> "" Then
    Call mciSendString("stop " & MIDIFile$, 0&, 0, 0)
End If
End Sub
Public Sub PlayWav(WavFile As String)
Dim SafeFile As String
SafeFile$ = Dir(WavFile$)
If SafeFile$ <> "" Then
    Call SndPlaySound(WavFile$, SND_FLAG)
End If
End Sub
Public Sub SetText(Window As Long, text As String)
Call SendMessageByString(Window&, WM_SETTEXT, 0&, text$)
End Sub
Public Function ListToString(thelist As ListBox, OneName As Boolean) As String
Dim DoList As Long, MailString As String
If thelist.List(0) = "" Then Exit Function
For DoList& = 0 To thelist.ListCount - 1
    If OneName = True Then
        MailString$ = MailString$ & "(" & thelist.List(DoList&) & "), "
    Else
        MailString$ = MailString$ & thelist.List(DoList&) & ", "
    End If
Next DoList&
MailString$ = Mid(MailString$, 1, Len(MailString$) - 2)
ListToString$ = MailString$
End Function
Public Sub FormOnTop(FormName As Form)
Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub WindowHide(hwnd As Long)
Call ShowWindow(hwnd&, SW_Hide)
End Sub
Public Sub WindowShow(hwnd As Long)
Call ShowWindow(hwnd&, SW_SHOW)
End Sub
Public Sub runmenu(TopMenu As Long, SubMenu As Long)
Dim aol As Long, aMenu As Long, sMenu As Long, mnID As Long
Dim mVal As Long
aol& = FindWindow("AOL Frame25", vbNullString)
aMenu& = GetMenu(aol&)
sMenu& = GetSubMenu(aMenu&, TopMenu&)
mnID& = GetMenuItemID(sMenu&, SubMenu&)
Call SendMessageLong(aol&, WM_COMMAND, mnID&, 0&)
End Sub
Public Sub RunMenuByString(SearchString As String)
Dim aol As Long, aMenu As Long, mCount As Long
Dim LookFor As Long, sMenu As Long, sCount As Long
Dim LookSub As Long, sID As Long, sString As String
aol& = FindWindow("AOL Frame25", vbNullString)
aMenu& = GetMenu(aol&)
mCount& = GetMenuItemCount(aMenu&)
For LookFor& = 0& To mCount& - 1
    sMenu& = GetSubMenu(aMenu&, LookFor&)
    sCount& = GetMenuItemCount(sMenu&)
    For LookSub& = 0 To sCount& - 1
        sID& = GetMenuItemID(sMenu&, LookSub&)
        sString$ = String$(100, " ")
        Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
        If InStr(LCase(sString$), LCase(SearchString$)) Then
            Call SendMessageLong(aol&, WM_COMMAND, sID&, 0&)
            Exit Sub
        End If
    Next LookSub&
Next LookFor&
End Sub
Function ChatName(ByVal partial As String) As String
On Error Resume Next
Dim cProcess As Long, itmHold As Long, screenname As String
Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
Dim rList As Long, sThread As Long, mThread As Long
Dim lIndex As Long, i As Integer, j As Integer, k As String
Room& = FindRoom&
If Room& = 0& Then Exit Function
rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
sThread& = GetWindowThreadProcessId(rList, cProcess&)
mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
If mThread& Then
    For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
        screenname$ = String$(4, vbNullChar)
        itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
        itmHold& = itmHold& + 24
        Call ReadProcessMemory(mThread&, itmHold&, screenname$, 4, rBytes)
        Call CopyMemory(psnHold&, ByVal screenname$, 4)
        psnHold& = psnHold& + 6
        screenname$ = String$(16, vbNullChar)
        Call ReadProcessMemory(mThread&, psnHold&, screenname$, Len(screenname$), rBytes&)
        screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
        i = InStr(LCase(screenname$), LCase(partial$))
        If i Then
            ChatName$ = screenname$
            DoEvents
            Exit Function
        End If
1   Next index&
For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
        screenname$ = String$(4, vbNullChar)
        itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
        itmHold& = itmHold& + 24
        Call ReadProcessMemory(mThread&, itmHold&, screenname$, 4, rBytes)
        Call CopyMemory(psnHold&, ByVal screenname$, 4)
        psnHold& = psnHold& + 6
        screenname$ = String$(16, vbNullChar)
        Call ReadProcessMemory(mThread&, psnHold&, screenname$, Len(screenname$), rBytes&)
        screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
        For i = 1 To Len(partial$)
            k = LCase(Mid(partial, i, 1))
            j = InStr(LCase(screenname$), k)
            If j = 0 Then GoTo 2
        Next i
        ChatName = screenname$
2   Next index&
Call CloseHandle(mThread)
End If
End Function
Function RoomEdit() As Long
Dim Room&
Room& = FindRoom&
RoomEdit& = Find2NdChildByClass(Room&, "RICHCNTL")
End Function
Function rgb2hex(R As Integer, G As Integer, b As Integer)
Dim a As String, C As Integer, NewR As String, NewG As String, NewB As String
a$ = Hex(R)
If Len(a$) = 1 Then a$ = "0" & a$
NewR = a$
a$ = Hex(G)
If Len(a$) = 1 Then a$ = "0" & a$
NewG = a$
a$ = Hex(b)
If Len(a$) = 1 Then a$ = "0" & a$
NewB = a$
rgb2hex = "<font color=#" & NewR & NewG & NewB & ">"
End Function
Function Fade1(what As String, R As Integer, G As Integer, b As Integer, ro As Integer, bo As Integer, go As Integer, Wavy As Boolean, Optional Bold As Boolean, Optional italic As Boolean, Optional Underline As Boolean, Optional strikethru As Boolean) As String
Dim rd As Integer, gd As Integer, bd As Integer, temp As Integer
Dim no1 As Boolean, no2 As Boolean, up As Boolean, dn As Boolean, lw As String, rw As String
Dim i As Integer, tempo As String
no1 = True: up = False: no2 = False: dn = False
rp = True
bp = True
gp = True
For i = 1 To Len(what)
    If Wavy = True Then
        If no1 = True Then no1 = False: up = True: no2 = False: dn = False: lw = "": rw = "": GoTo 1
        If up = True Then no1 = False: up = False: no2 = True: dn = False: lw = "<sup>": rw = "</sup>": GoTo 1
        If no2 = True Then no1 = False: up = False: no2 = False: dn = True: lw = "": rw = "": GoTo 1
        If dn = True Then no1 = True: up = False: no2 = False: dn = False: lw = "<sub>": rw = "</sub>": GoTo 1
1       tempo = tempo & lw & rgb2hex(R, G, b) & Mid(what, i, 1) & rw
        GoSub add
    Else
        tempo = tempo & rgb2hex(R, G, b) & Mid(what, i, 1)
        GoSub add
    End If
Next i
If Bold = True Then tempo = "<B>" & tempo
If italic = True Then tempo = "<i>" & tempo
If Underline = True Then tempo = "<u>" & tempo
If strikethru = True Then tempo = "<s>" & tempo
Fade1 = tempo
Exit Function

add:
If rp = True Then
    If R + ro <= 255 Then
        R = R + ro
    ElseIf R + ro > 255 Then
        rd = (R + rd) - 255
        rp = False
    End If
ElseIf rp = False Then
    If R - ro >= 0 Then
        R = R - ro
    ElseIf R - ro < 0 Then
        temp = ro - R
        R = 0 + temp
        rp = True
    End If
End If

If gp = True Then
    If G + go <= 255 Then
        G = G + go
    ElseIf G + go > 255 Then
        gd = (G + gd) - 255
        gp = False
    End If
ElseIf gp = False Then
    If G - go >= 0 Then
        G = G - go
    ElseIf G - go < 0 Then
        temp = go - G
        G = 0 + temp
        gp = True
    End If
End If

If bp = True Then
    If b + bo <= 255 Then
        b = b + bo
    ElseIf b + bo > 255 Then
        bd = (b + bd) - 255
        bp = False
    End If
ElseIf bp = False Then
    If b - bo >= 0 Then
        b = b - bo
    ElseIf b - bo < 0 Then
        temp = bo - b
        b = 0 + temp
        bp = True
    End If
End If
Return

End Function
Function FindChildByClassEx(parentHwnd As Long, childhand As String, childnumber As Integer)
Dim ReturnString$, handles&, Parent, Copy, curr
curr = 1
Copy = parentHwnd
Parent = GetWindow(parentHwnd, 5)
Top:
ReturnString$ = String$(250, 0)
handles& = GetClassName(Parent, ReturnString$, 250)
If Left$(ReturnString$, Len(childhand)) Like childhand Then
    If curr = childnumber Then
        GoTo ending:
    Else
        curr = curr + 1
    End If
End If
Parent = GetWindow(Parent, 2)
If Parent > 0 Then GoTo Top

ending:
FindChildByClassEx = Parent
parentHwnd = Copy
End Function
Function ColorEx(what As String, color1 As RGB, color2 As RGB) As String
ColorEx = rgb2hex(color1.R, color1.G, color1.b) & what & rgb2hex(color2.R, color2.G, color2.b)
End Function
Function FontEX(what As String, font1 As String, font2 As String) As String
FontEX = "<font face=" & Chr(34) & font1 & Chr(34) & ">" & what & "<font face=" & Chr(34) & font2 & Chr(34) & ">"
End Function
Function BU(what As String) As String
BU = "<b>" & Left(what, 1) & "</b><u>" & Mid(what, 2) & "</u>"
End Function
Function LoadColor(Source As RGB) As RGB
LoadColor.R = Source.R
LoadColor.G = Source.G
LoadColor.b = Source.b
End Function
Function NumberCheck(numb As String) As Boolean
Dim i%
For i = 1 To Len(numb)
    If Mid(numb, i, 1) = "0" Or Mid(numb, i, 1) = "1" Or Mid(numb, i, 1) = "2" Or Mid(numb, i, 1) = "3" Or Mid(numb, i, 1) = "4" Or Mid(numb, i, 1) = "5" Or Mid(numb, i, 1) = "6" Or Mid(numb, i, 1) = "7" Or Mid(numb, i, 1) = "8" Or Mid(numb, i, 1) = "9" Or Mid(numb, i, 1) = "-" Then
        NumberCheck = True
    Else
        NumberCheck = False
        Exit For
    End If
Next i
End Function
Function RoomFont1(ByVal FontN As String) As String
On Error Resume Next
Dim cProcess As Long, itmHold As Long, FontName As String
Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
Dim rList As Long, sThread As Long, mThread As Long
Dim lIndex As Long
Room& = FindRoom&
If Room& = 0& Then Exit Function
rList& = FindWindowEx(Room&, 0&, "_AOL_Combobox", vbNullString)
sThread& = GetWindowThreadProcessId(rList, cProcess&)
mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
If mThread& Then
    For index& = 0 To SendMessage(rList, CB_GETCOUNT, 0, 0) - 1
        FontName$ = String$(4, vbNullChar)
        itmHold& = SendMessage(rList, CB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
        itmHold& = itmHold& + 24
        Call ReadProcessMemory(mThread&, itmHold&, FontName$, 4, rBytes)
        Call CopyMemory(psnHold&, ByVal FontName$, 4)
        psnHold& = psnHold& + 6
        FontName$ = String$(26, vbNullChar)
        Call ReadProcessMemory(mThread&, psnHold&, FontName$, Len(FontName$), rBytes&)
        FontName$ = Left$(FontName$, InStr(FontName$, vbNullChar) - 1)
        If InStr(LCase(FontName$), LCase(FontN)) Then
            lIndex& = index&
            DoEvents
            RoomFont1 = FontName$
            Exit Function
        End If
    Next index&
    Call CloseHandle(mThread)
End If

End Function
Function RoomFontEx(ByVal FontN As String) As String
On Error Resume Next
Dim cProcess As Long, itmHold As Long, FontName As String
Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
Dim rList As Long, sThread As Long, mThread As Long
Dim lIndex As Long
Room& = FindRoom&
If Room& = 0& Then Exit Function
rList& = FindWindowEx(Room&, 0&, "_AOL_Combobox", vbNullString)
sThread& = GetWindowThreadProcessId(rList, cProcess&)
mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
If mThread& Then
    For index& = 0 To SendMessage(rList, CB_GETCOUNT, 0, 0) - 1
        FontName$ = String$(4, vbNullChar)
        itmHold& = SendMessage(rList, CB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
        itmHold& = itmHold& + 24
        Call ReadProcessMemory(mThread&, itmHold&, FontName$, 4, rBytes)
        Call CopyMemory(psnHold&, ByVal FontName$, 4)
        psnHold& = psnHold& + 6
        FontName$ = String$(26, vbNullChar)
        Call ReadProcessMemory(mThread&, psnHold&, FontName$, Len(FontName$), rBytes&)
        FontName$ = Left$(FontName$, InStr(FontName$, vbNullChar) - 1)
        If LCase(FontName$) = LCase(FontN) Then
            lIndex& = index&
            DoEvents
            RoomFontEx = FontName$
            Exit Function
        End If
    Next index&
    Call CloseHandle(mThread)
End If
End Function
Sub AddRoomToListBox(thelist As ListBox, AddUser As Boolean)
On Error Resume Next
Dim cProcess As Long, itmHold As Long, screenname As String
Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
Dim rList As Long, sThread As Long, mThread As Long
Room& = FindRoom&
If Room& = 0& Then Exit Sub
rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
sThread& = GetWindowThreadProcessId(rList, cProcess&)
mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
If mThread& Then
    For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
        screenname$ = String$(4, vbNullChar)
        itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
        itmHold& = itmHold& + 24
        Call ReadProcessMemory(mThread&, itmHold&, screenname$, 4, rBytes)
        Call CopyMemory(psnHold&, ByVal screenname$, 4)
        psnHold& = psnHold& + 6
        screenname$ = String$(16, vbNullChar)
        Call ReadProcessMemory(mThread&, psnHold&, screenname$, Len(screenname$), rBytes&)
        screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
        If screenname$ <> GetUser$ Or AddUser = True Then
            thelist.AddItem screenname$
        End If
    Next index&
    Call CloseHandle(mThread)
End If
End Sub
Public Function FindIM() As Long
Dim aol As Long, MDI As Long, child As Long, Caption As String
aol& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
Caption$ = GetCaption(child&)
If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
    FindIM& = child&
    Exit Function
Else
    Do
        child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
        Caption$ = GetCaption(child&)
        If InStr(Caption$, "Instant Message") > 3 Then
            FindIM& = child&
            Exit Function
        End If
    Loop Until child& = 0&
End If
FindIM& = child&
End Function
Public Function FindInfoWindow() As Long
Dim aol As Long, MDI As Long, child As Long
Dim AOLCheck As Long, aolicon As Long, AOLStatic As Long
Dim AOLIcon2 As Long, AOLGlyph As Long
aol& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
AOLCheck& = FindWindowEx(child&, 0&, "_AOL_Checkbox", vbNullString)
AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
AOLGlyph& = FindWindowEx(child&, 0&, "_AOL_Glyph", vbNullString)
aolicon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(child&, aolicon&, "_AOL_Icon", vbNullString)
If AOLCheck& <> 0& And AOLStatic& <> 0& And AOLGlyph& <> 0& And aolicon& <> 0& And AOLIcon2& <> 0& Then
    FindInfoWindow& = child&
    Exit Function
Else
    Do
        child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
        AOLCheck& = FindWindowEx(child&, 0&, "_AOL_Checkbox", vbNullString)
        AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
        AOLGlyph& = FindWindowEx(child&, 0&, "_AOL_Glyph", vbNullString)
        aolicon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon2& = FindWindowEx(child&, aolicon&, "_AOL_Icon", vbNullString)
        If AOLCheck& <> 0& And AOLStatic& <> 0& And AOLGlyph& <> 0& And aolicon& <> 0& And AOLIcon2& <> 0& Then
            FindInfoWindow& = child&
            Exit Function
        End If
    Loop Until child& = 0&
End If
FindInfoWindow& = child&
End Function
Function FileSearch(FileName As String, lst As ListBox)
Dim z As String, a As Integer, i%, dd$
Dim lines$
lst.Clear
For i = 65 To 90
    If GetDriveType(Chr(i) & ":\") <> 3 Then GoTo 1
    Pause (1.5)
    Call LoadFiles(FileName, Chr(i) & ":\")
    Pause (1.5)
    Close 1
    Open "c:" & "\songs.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, lines
        lst.AddItem lines
    Loop
    Close #1
1 Next i
End Function
Sub LoadFiles(FileName As String, FilePath As String)
Dim FileShell As String, IsThere As Variant
On Error Resume Next
Kill "c:\songs.txt"
FileShell = "C:\command.com /C dir " & FilePath & FileName & " /s/b > " & "c:" & "\songs.txt"
IsThere = Shell(FileShell)
On Error Resume Next
Kill "c:\songs.txt"
End Sub
Public Function AutoFormShape(frm As Form, transColor)
Dim x, y As Integer, success As Long
CurRgn = CreateRectRgn(0, 0, frm.ScaleWidth, frm.ScaleHeight)
While y <= frm.ScaleHeight
    While x <= frm.ScaleWidth
        If GetPixel(frm.hdc, x, y) = transColor Then
            TempRgn = CreateRectRgn(x, y, x + 1, y + 1)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
        End If
        x = x + 1
    Wend
        y = y + 1
        x = 0
Wend
success = SetWindowRgn(frm.hwnd, CurRgn, True)
DeleteObject (CurRgn)
End Function
Public Sub TransparentForm(frm As Form, transparent As Boolean)
Const RGN_DIFF = 4
Const RGN_OR = 2
Dim outer_rgn As Long, inner_rgn As Long, wid As Single, hgt As Single
Dim border_width As Single, title_height As Single, ctl_left As Single
Dim ctl_top As Single, ctl_right As Single, ctl_bottom As Single
Dim control_rgn As Long, combined_rgn As Long, ctl As Control
If transparent = True Then
With frm
    If .WindowState = vbMinimized Then Exit Sub
    frm.ScaleMode = 3
    wid = .ScaleX(.Width, vbTwips, vbPixels)
    hgt = .ScaleY(.Height, vbTwips, vbPixels)
    outer_rgn = CreateRectRgn(0, 0, wid, hgt)
    border_width = (wid - .ScaleWidth) / 2
    title_height = hgt - border_width - .ScaleHeight
    inner_rgn = CreateRectRgn(border_width, title_height, wid - border_width, hgt - border_width)
    combined_rgn = CreateRectRgn(0, 0, 0, 0)
    CombineRgn combined_rgn, outer_rgn, inner_rgn, RGN_DIFF
    For Each ctl In .Controls
        If ctl.Container Is frm Then
            ctl_left = .ScaleX(ctl.Left, frm.ScaleMode, vbPixels) + border_width
            ctl_top = .ScaleX(ctl.Top, frm.ScaleMode, vbPixels) + title_height
            ctl_right = .ScaleX(ctl.Width, frm.ScaleMode, vbPixels) + ctl_left
            ctl_bottom = .ScaleX(ctl.Height, frm.ScaleMode, vbPixels) + ctl_top
            control_rgn = CreateRectRgn(ctl_left, ctl_top, ctl_right, ctl_bottom)
            CombineRgn combined_rgn, combined_rgn, control_rgn, RGN_OR
        End If
    Next ctl
    SetWindowRgn .hwnd, combined_rgn, True
End With
ElseIf transparent = False Then
    If frm.WindowState = vbMinimized Then Exit Sub
    frm.ScaleMode = 3
    wid = frm.ScaleX(frm.Width, vbTwips, vbPixels)
    hgt = frm.ScaleY(frm.Height, vbTwips, vbPixels)
    outer_rgn = CreateRectRgn(0, 0, wid, hgt)
    border_width = (wid - frm.ScaleWidth) / 2
    title_height = hgt - border_width - frm.ScaleHeight
    inner_rgn = CreateRectRgn(border_width, title_height, wid - border_width, hgt - border_width)
    For Each ctl In frm.Controls
        If ctl.Container Is frm Then
            ctl_left = frm.ScaleX(ctl.Left, frm.ScaleMode, vbPixels) + border_width
            ctl_top = frm.ScaleX(ctl.Top, frm.ScaleMode, vbPixels) + title_height
            ctl_right = frm.ScaleX(ctl.Width, frm.ScaleMode, vbPixels) + ctl_left
            ctl_bottom = frm.ScaleX(ctl.Height, frm.ScaleMode, vbPixels) + ctl_top
            control_rgn = CreateRectRgn(ctl_left, ctl_top, ctl_right, ctl_bottom)
            CombineRgn combined_rgn, combined_rgn, control_rgn, RGN_OR
        End If
    Next ctl
    SetWindowRgn frm.hwnd, combined_rgn, True
End If
End Sub
Function Underline(TheText As String, WhatToUnderline As String) As String
Dim x1 As Integer, x2 As Integer
x1 = InStr(LCase(TheText$), LCase(WhatToUnderline$))
If x1 = 0 Then Underline = TheText: Exit Function
x2 = x1 + Len(WhatToUnderline$)
Underline = Left(TheText, x1 - 1) & "<u>" & Mid(TheText, x1, Len(WhatToUnderline$)) & "</u>" & Mid(TheText, x2)
End Function
Public Sub BuddiesToListBox(thelist As ListBox)
On Error Resume Next
Dim cProcess As Long, itmHold As Long, screenname As String
Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
Dim rList As Long, sThread As Long, mThread As Long
thelist.Clear
Room& = FindChildByTitle(MDIWindow&, "Buddy List Window")
If Room& = 0 Then
    Call Keyword("buddyview")
    Do: DoEvents
        Room& = FindChildByTitle(MDIWindow&, "Buddy List Window")
    Loop While Room& = 0
    Pause 0.5
End If

rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
sThread& = GetWindowThreadProcessId(rList, cProcess&)
mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
If mThread& Then
    For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
        screenname$ = String$(4, vbNullChar)
        itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
        itmHold& = itmHold& + 24
        Call ReadProcessMemory(mThread&, itmHold&, screenname$, 4, rBytes)
        Call CopyMemory(psnHold&, ByVal screenname$, 4)
        psnHold& = psnHold& + 6
        screenname$ = String$(30, vbNullChar)
        Call ReadProcessMemory(mThread&, psnHold&, screenname$, Len(screenname$), rBytes&)
        screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
        If Right(screenname$, 1) = "*" Then screenname = Left(screenname, Len(screenname) - 1)
        If Right(screenname$, 1) <> ")" Then
            thelist.AddItem RTrim(LTrim(screenname$))
        End If
    Next index&
    Call CloseHandle(mThread)
End If
End Sub
Private Function GetProcessHandle(ByVal hwnd As Long) As Long
On Error Resume Next
Dim ThreadId As Long
Dim ProcessID As Long
ThreadId = GetWindowThreadProcessId(hwnd, ProcessID)
GetProcessHandle = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, ProcessID)
End Function
Function RegGetString$(hInKey As Long, ByVal subkey$, ByVal valname$)
Dim retval$, hSubKey As Long, dwType As Long, SZ As Long
Dim R As Long
Dim v$
retval$ = ""
Const KEY_ALL_ACCESS As Long = &HF0063
Const ERROR_SUCCESS As Long = 0
Const REG_SZ As Long = 1
R = RegOpenKeyEx(hInKey, subkey$, 0, KEY_ALL_ACCESS, hSubKey)
If R <> ERROR_SUCCESS Then GoTo Quit_Now
SZ = 256: v$ = String$(SZ, 0)
R = RegQueryValueEx(hSubKey, valname$, 0, dwType, ByVal v$, SZ)
If R = ERROR_SUCCESS And dwType = REG_SZ Then
retval$ = Left$(v$, SZ)
Else
retval$ = Left$(v$, SZ)
End If
If hInKey = 0 Then R = RegCloseKey(hSubKey)
Quit_Now:
RegGetString$ = retval$
End Function
Function GetBit() As String
GetBit = RegGetString$(HKEY_CURRENT_CONFIG, "Display\Settings", "BitsPerPixel")
End Function
Function GetCurrentPrinter() As String
GetCurrentPrinter = RegGetString$(HKEY_CURRENT_CONFIG, "System\CurrentControlSet\Control\Print\Printers", "Default")
End Function
Function GetWinUserName() As String
GetWinUserName = RegGetString$(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
End Function
Function GetWinVer() As String
GetWinVer = RegGetString$(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "VersionNumber")
End Function
Function GetWVer2() As String
GetWVer2 = RegGetString$(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProductName")
End Function
Function AOLShell()
AOLShell = RegGetString$(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Aol.exe", "Path")
End Function
Function AOLModal()
AOLModal = FindWindow("_AOL_Modal", vbNullString)
End Function
Sub GuestSetSNPW(screenname$, password$)
Dim aol&, x&
aol& = AOLModal()
aol& = FindChildByClass(aol&, "_AOL_Edit")
x& = SendMessageByString(aol&, WM_SETTEXT, 0, screenname$)
aol& = GetWindow(aol&, GW_HWNDNEXT)
aol& = GetWindow(aol&, GW_HWNDNEXT)
x = SendMessageByString(aol&, WM_SETTEXT, 0, password$)
End Sub
Function Bold(what As String) As String
Bold = "<b>" & what$ & "</b>"
End Function
Sub OpenCD()
Dim retvalue&
retvalue& = mciSendString("set CDAudio door open", vbNullString, 0, 0)
End Sub
Sub CloseCD()
Dim retvalue&
retvalue = mciSendString("set CDAudio door closed", vbNullString, 0, 0)
End Sub
