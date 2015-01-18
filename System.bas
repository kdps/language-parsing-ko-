Attribute VB_Name = "System"
Option Explicit
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Public Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, ByRef lpVolumeSerialNumber As Long, ByRef lpMaximumComponentLength As Long, ByRef lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpszCommand As String, ByVal lpszReturnString As String, ByVal cchReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Sub GetSystemInfo Lib "kernel32.dll" (ByRef lpSystemInfo As SYSTEM_INFO)

'시스템 정보
Public Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

Public Const LWA_ALPHA As Long = &H2
Public Const WS_EX_LAYERED As Long = &H80000
Public Const GWL_EXSTYLE As Long = -20
Public Const SW_SHOW As Long = 5
Public Const RDW_UPDATENOW As Long = &H100

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

'모니터
Public Const HWND_BROADCAST& = &HFFFF&
Public Const WM_SYSCOMMAND& = &H112&
Public Const SC_MONITORPOWER& = &HF170&
Public Const MONITOR_ON& = &HFFFFFFFF
Public Const MONITOR_OFF& = 2&

Public Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szexeFile As String * 6400
End Type

Function RoundFRM(frm As Form)
Dim Result As Long
Result = CreateRoundRectRgn(0, 0, (frm.Width + 1) / Screen.TwipsPerPixelX, (frm.Height + 1) / Screen.TwipsPerPixelY, 13, 13)
SetWindowRgn frm.hWnd, Result, True
End Function

Public Function GetMainSerialNumber() As String
Dim VolumeSerialNumber As Long
If GetVolumeInformation("C:\", vbNullString, 0&, VolumeSerialNumber, ByVal 0&, ByVal 0&, vbNullString, 0&) Then
    GetMainSerialNumber = Left$(Hex$(VolumeSerialNumber), 4) & "-" & Mid$(Hex$(VolumeSerialNumber), 5, 4)
Else
    GetMainSerialNumber = "0000-0000"
End If
End Function

Public Sub TurnOnMonitor()
SendMessage HWND_BROADCAST, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal MONITOR_ON
End Sub

Public Sub TurnOffMonitor()
SendMessage HWND_BROADCAST, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal MONITOR_OFF
End Sub

Public Sub ChangeVar(TheVar As String, NewVal As Variant, TheProfile As String)
Dim i As Integer
For i = var_name.Count To 1 Step -1
    If var_name(i) = TheVar And var_cont(i) = NewVal And var_profile(i) = TheProfile Then
        var_cont.Remove i
        var_name.Remove i
        var_profile.Remove i
        var_cont.Add NewVal
        var_name.Add TheVar
        var_profile.Add TheProfile
        Exit Sub
    Else
        var_cont.Add NewVal
        var_name.Add TheVar
        var_profile.Add TheProfile
        Exit Sub
    End If
Next i
End Sub

Public Function GetVar(TheVar As String) As Variant
Dim i As Integer
For i = 1 To var_name.Count
    If var_name(i) = TheVar And var_profile(i) = "" Then
        GetVar = var_cont(i)
        Exit Function
    End If
Next i
End Function

Public Function GetAddon(TheVar As String) As Variant
Dim i As Integer
For i = 1 To var_name.Count
    If var_name(i) = TheVar And var_addon(i) <> "" Then
        GetAddon = var_addon(i)
        Exit Function
    End If
Next i
End Function

Public Function GetProfile2(TheVar As String) As Variant
Dim i As Integer
For i = 1 To var_name.Count
    If var_name(i) = TheVar Then
        GetProfile2 = var_profile(i)
        Exit Function
    End If
Next i
End Function

Public Function GetProfile(TheVar As String, TheProfile As String) As Variant
Dim i As Integer
For i = 1 To var_name.Count
    If var_name(i) = TheVar And var_profile(i) = TheProfile Then
        GetProfile = var_cont(i)
        Exit Function
    End If
Next i
End Function

Public Function KillAppByName(MyName As String) As Boolean
    Const PROCESS_ALL_ACCESS = 0
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim AppKill As Boolean
    Dim appCount As Integer
    Dim i As Integer
    On Local Error GoTo Finish
    appCount = 0
    
    Const TH32CS_SNAPPROCESS As Long = 2&
    
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    
    Do While rProcessFound
    i = InStr(1, uProcess.szexeFile, Chr(0))
    szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
    If Right$(szExename, Len(MyName)) = LCase$(MyName) Then
    KillAppByName = True
    appCount = appCount + 1
    myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
    AppKill = TerminateProcess(myProcess, exitCode)
    Call CloseHandle(myProcess)
    End If
    rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop
    
    Call CloseHandle(hSnapshot)
Finish:
End Function

