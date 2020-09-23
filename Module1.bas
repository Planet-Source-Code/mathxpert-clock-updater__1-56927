Attribute VB_Name = "Module1"
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3
Private Const scUserAgent = "VB Project"
Private Const INTERNET_FLAG_RELOAD = &H80000000

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hOpen As Long, ByVal sUrl As String, ByVal sHeaders As String, ByVal lLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type
   
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Enum T_KeyClasses
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_EVENT = &H1
Private Const KEY_NOTIFY = &H10
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Private Const KEY_ALL_ACCESS = (STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE)

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Declare Sub GetLocalTime Lib "kernel32" (localTime As SYSTEMTIME)
Private Declare Sub SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Public RunTimer As Long
Public bUsed As Boolean
Public Cancelled As Boolean
Public TimeD As Single
Public ErrCode As Long
Public MSAdv As Integer
Public bUncheck As Boolean
Public bUnloading As Boolean
Public bFocus As Boolean

Public Function IsWinNT() As Boolean
Dim osvi As OSVERSIONINFO
osvi.dwOSVersionInfoSize = Len(osvi)
GetVersionEx osvi
IsWinNT = (osvi.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

Public Function GetStartupRegValue() As String
GetStartupRegValue = "Software\Microsoft\Windows\CurrentVersion\Run" & IIf(IsWinNT, "", "Services")
End Function

Public Function FormatAppPath(Path As String) As String
If InStr(Path, " ") > 0 Then
    FormatAppPath = """" & Path & """"
Else
    FormatAppPath = Path
End If
End Function

Public Function ValueExists(rClass As T_KeyClasses, Path As String, sKey As String) As Boolean
Dim res As Long
Dim hKey As Long
Dim tmpVal As Long
Dim KeyValSize As Long

On Error GoTo Catch

res = RegOpenKeyEx(rClass, Path, 0, KEY_ALL_ACCESS, hKey)
If res <> 0 Then Exit Function

res = RegQueryValueExNULL(hKey, sKey, 0, tmpVal, 0, KeyValSize)
If res <> 0 Then Exit Function

ValueExists = True

Catch:
End Function

Public Sub DeleteValue(rClass As T_KeyClasses, Path As String, sKey As String)
Dim hKey As Long
Dim res As Long

res = RegOpenKeyEx(rClass, Path, 0, KEY_ALL_ACCESS, hKey)
res = RegDeleteValue(hKey, sKey)
RegCloseKey hKey
End Sub

Public Function GetRegValue(KeyRoot As T_KeyClasses, Path As String, sKey As String) As String
Dim hKey As Long
Dim KeyValType As Long
Dim KeyValSize As Long
Dim KeyVal As String
Dim tmpVal As String
Dim res As Long
Dim i As Integer

res = RegOpenKeyEx(KeyRoot, Path, 0, KEY_ALL_ACCESS, hKey)
If res <> 0 Then GoTo Errore

tmpVal = String$(1024, 0)
KeyValSize = 1024
res = RegQueryValueEx(hKey, sKey, 0, KeyValType, tmpVal, KeyValSize)
If res <> 0 Then GoTo Errore

If Asc(Mid$(tmpVal, KeyValSize, 1)) = 0 Then
    tmpVal = Left$(tmpVal, KeyValSize - 1)
Else
    tmpVal = Left$(tmpVal, KeyValSize)
End If

Select Case KeyValType
    Case REG_SZ
        KeyVal = tmpVal
    Case REG_DWORD
        For i = Len(tmpVal) To 1 Step -1
            KeyVal = KeyVal + Hex(Asc(Mid$(tmpVal, i, 1)))
        Next
        KeyVal = Format$("&h" + KeyVal)
End Select

GetRegValue = KeyVal
RegCloseKey hKey
Exit Function

Errore:
    GetRegValue = ""
    RegCloseKey hKey
End Function

Public Function SetRegValue(KeyRoot As T_KeyClasses, Path As String, sKey As String, NewValue As String) As Boolean
Dim hKey As Long
Dim KeyValType As Long
Dim KeyValSize As Long
Dim KeyVal As String
Dim tmpVal As String
Dim res As Long
Dim i As Integer
Dim x As Long

res = RegOpenKeyEx(KeyRoot, Path, 0, KEY_ALL_ACCESS, hKey)
If res <> 0 Then GoTo Errore

tmpVal = String$(1024, 0)
KeyValSize = 1024
res = RegQueryValueEx(hKey, sKey, 0, KeyValType, tmpVal, KeyValSize)

Select Case res
    Case 2
        KeyValType = REG_SZ
    Case Is <> 0
        GoTo Errore
End Select

Select Case KeyValType
    Case REG_SZ
        tmpVal = NewValue
    Case REG_DWORD
        x = Val(NewValue)
        tmpVal = ""
        For i = 0 To 3
            tmpVal = tmpVal & Chr$(x Mod 256)
            x = x / 256
        Next
End Select

KeyValSize = Len(tmpVal)
res = RegSetValueEx(hKey, sKey, 0, KeyValType, tmpVal, KeyValSize)
If res <> 0 Then GoTo Errore

SetRegValue = True
RegCloseKey hKey
Exit Function

Errore:
    SetRegValue = False
    RegCloseKey hKey
End Function

Public Function TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long) As Long
Form2.Label1.Caption = "System time:  " & Format$(Now, "M/D/YY  H:MM:SS." & MS & " AM/PM")
End Function

Public Function MS() As String
Dim MySys As SYSTEMTIME, tmp As String
GetLocalTime MySys
tmp = CStr(MySys.wMilliseconds)
MS = Format$(tmp, "000")
End Function

Public Sub ChangeShapeColors()
With Form1
    .Shape1.BackColor = vbButtonFace
    .Shape2.BackColor = vbButtonFace
    .Shape3.BackColor = vbButtonFace
    .Shape4.BackColor = vbButtonFace
    .Shape5.BackColor = vbButtonFace
    .Shape6.BackColor = vbButtonFace
    .Shape7.BackColor = vbButtonFace
End With
End Sub

Public Sub RestoreEverything(ByVal sText As String)
ChangeShapeColors
With Form1
    .Label1.Caption = sText
    .MousePointer = vbNormal
    .Command1.Enabled = True
    .mPopSync.Enabled = True
    .Command2.Enabled = False
    .mPopStop.Enabled = False
    .ChangeIcon True
End With
bUsed = False
End Sub

Private Function GetTimeProps(ByVal TheDate As String) As Date
Dim sDate As String, sTime As String, tTheDate As String, ClnPl As Long

ErrCode = 0

tTheDate = TheDate
ClnPl = InStr(tTheDate, ":")

sDate = Mid$(tTheDate, ClnPl - 8, 5) & "-" & Mid$(tTheDate, ClnPl - 11, 2)
sTime = Mid$(tTheDate, ClnPl - 2, 8)

If IsDate(sDate) And IsDate(sTime) Then
    If Mid$(tTheDate, ClnPl + 12, 1) <> "0" Then
        ErrCode = 2
    Else
        MSAdv = Mid(tTheDate, ClnPl + 14, 3) & Mid(tTheDate, ClnPl + 18, 1)
        GetTimeProps = CDate(sDate & " " & sTime)
    End If
Else
    ErrCode = 1
End If
End Function

Public Sub OnSynchro()

If bUsed = False Then
    bUsed = True
    Cancelled = False

    Dim sHost As String
    Dim hInternet As Long
    Dim hHttp As Long
    Dim bRet As Boolean
    Dim sBuff As String * 2048
    Dim lNumberOfBytesRead As Long
    Dim sBuffer As String
    
    sHost = Form1.Combo1.Text
    Form1.ChangeIcon
    With Form1
        .Label1.Caption = "Synchronizing; please wait..."
        .Label4.Caption = ""
        .Label5.Caption = ""
        DoEvents
        .MousePointer = vbHourglass
        .Command1.Enabled = False
        .mPopSync.Enabled = False
        .Command2.Enabled = True
        .mPopStop.Enabled = True
    End With
    
    If Cancelled = False Then
        hInternet = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
        If hInternet = 0 Then
            RestoreEverything "Cannot preconfigure the connection!"
            Exit Sub
        End If
        DoEvents
        Form1.Shape1.BackColor = vbHighlight
        DoEvents
    Else
        RestoreEverything "Synchronization aborted."
        Exit Sub
    End If
    
    If Cancelled = False Then
        hHttp = InternetOpenUrl(hInternet, "http://" & sHost & ":13/", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
        If hHttp = 0 Then
            RestoreEverything "Cannot find host!"
            Exit Sub
        End If
        DoEvents
        Form1.Shape2.BackColor = vbHighlight
        DoEvents
    Else
        RestoreEverything "Synchronization aborted."
        InternetCloseHandle hHttp
        InternetCloseHandle hInternet
        Exit Sub
    End If
    
    If Cancelled = False Then
        sBuff = vbNullString
        bRet = InternetReadFile(hHttp, sBuff, Len(sBuff), lNumberOfBytesRead)
        sBuffer = sBuffer & Left$(sBuff, lNumberOfBytesRead)
        DoEvents
        Form1.Shape3.BackColor = vbHighlight
        DoEvents
    Else
        RestoreEverything "Synchronization aborted."
        InternetCloseHandle hHttp
        InternetCloseHandle hInternet
        Exit Sub
    End If
    
    If Cancelled = False Then
        If hHttp <> 0 Then InternetCloseHandle hHttp
        If hInternet <> 0 Then InternetCloseHandle hInternet
        DoEvents
        Form1.Shape4.BackColor = vbHighlight
        DoEvents
    Else
        RestoreEverything "Synchronization aborted."
        InternetCloseHandle hHttp
        InternetCloseHandle hInternet
        Exit Sub
    End If
    
    Dim dtOldTime As SYSTEMTIME
    Dim tmUniversal As SYSTEMTIME
    Dim tmLocal As SYSTEMTIME
    
    If Cancelled = False Then
        GetLocalTime dtOldTime
        DoEvents
        Form1.Shape5.BackColor = vbHighlight
        DoEvents
    Else
        RestoreEverything "Synchronization aborted."
        Exit Sub
    End If
    
    DoEvents
    Form1.Shape6.BackColor = vbHighlight
    DoEvents
    
    If Cancelled = False Then
        Dim AccDt As Date
        AccDt = GetTimeProps(sBuff)
        If ErrCode = 1 Then
            RestoreEverything "Time error!"
            Exit Sub
        ElseIf ErrCode = 2 Then
            RestoreEverything "Inaccurate time from server!"
            Exit Sub
        End If
        If MSAdv <> 0 Then AccDt = DateAdd("s", -1, AccDt)
        With tmUniversal
            .wYear = Year(AccDt)
            .wMonth = Month(AccDt)
            .wDay = Day(AccDt)
            .wHour = Hour(AccDt)
            .wMinute = Minute(AccDt)
            .wSecond = Second(AccDt)
            If MSAdv = 0 Then
                .wMilliseconds = 0
            Else
                .wMilliseconds = (10000 - MSAdv) / 10
            End If
        End With
        
        SetSystemTime tmUniversal
        GetLocalTime tmLocal
        
        Dim hr, mn, sec, msec, Nhr, Nmn, Nsec, Nmsec
        hr = Right(CStr(dtOldTime.wHour), 2)
        hr = IIf(hr < 10, "0" & Right(hr, 1), hr)
        mn = Right(CStr(dtOldTime.wMinute), 2)
        mn = IIf(mn < 10, "0" & Right(mn, 1), mn)
        sec = Right(CStr(dtOldTime.wSecond), 2)
        sec = IIf(sec < 10, "0" & Right(sec, 1), sec)
        msec = Right(CStr(dtOldTime.wMilliseconds), 3)
        msec = IIf(msec = 0, "000", IIf(msec < 10, "00" & Right(msec, 1), IIf(msec < 100, "0" & Right(msec, 2), msec)))
        Nhr = Right(CStr(tmLocal.wHour), 2)
        Nhr = IIf(Nhr < 10, "0" & Right(Nhr, 1), Nhr)
        Nmn = Right(CStr(tmLocal.wMinute), 2)
        Nmn = IIf(Nmn < 10, "0" & Right(Nmn, 1), Nmn)
        Nsec = Right(CStr(tmLocal.wSecond), 2)
        Nsec = IIf(Nsec < 10, "0" & Right(Nsec, 1), Nsec)
        Nmsec = Right(CStr(tmLocal.wMilliseconds), 3)
        Nmsec = IIf(Nmsec = 0, "000", IIf(Nmsec < 10, "00" & Right(Nmsec, 1), IIf(Nmsec < 100, "0" & Right(Nmsec, 2), Nmsec)))
        
        With Form1
            .MousePointer = vbNormal
            .Command1.Enabled = True
            .mPopSync.Enabled = True
            .Command2.Enabled = False
            .mPopStop.Enabled = False
            .Label1.Caption = "Successful synchronization!"
            .Label4.Caption = Format$(CStr(dtOldTime.wMonth) & "/" & CStr(dtOldTime.wDay) & "/" & CStr(dtOldTime.wYear) & " " & hr & ":" & mn & ":" & sec, "M/D/YY  H:MM:SS." & msec & " AM/PM")
            .Label5.Caption = Format$(CStr(tmLocal.wMonth) & "/" & CStr(tmLocal.wDay) & "/" & CStr(tmLocal.wYear) & " " & Nhr & ":" & Nmn & ":" & Nsec, "M/D/YY  H:MM:SS." & Nmsec & " AM/PM")
            .ChangeInterval 12 - .VScroll1.Value
            .ChangeIcon True
        End With
        
        DoEvents
        Form1.Shape7.BackColor = vbHighlight
        DoEvents
    Else
        RestoreEverything "Synchronization aborted."
        Exit Sub
    End If
    
    ChangeShapeColors
    
    bUsed = False

End If
End Sub
