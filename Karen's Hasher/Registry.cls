VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Provides convenient access to the Windows Registry."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
' Copyright © 1998-2005 Karen Kenworthy
' All Rights Reserved
' http://www.karenware.com/
' Version 1.17 7/21/2005

Public Enum REG_ERROR
    ERROR_SUCCESS = 0
    ERROR_ACCESS_DENIED = 5
    ERROR_MORE_DATA = 234
    ERROR_NO_MORE_ITEMS = 259
    ERROR_BADKEY = 1010
    ERROR_CANTOPEN = 1011
    ERROR_CANTREAD = 1012
    ERROR_REGISTRY_CORRUPT = 1015
End Enum

Public Enum REG_HKEY
    HKEY_DEFAULT = 0
    HKEY_DYN_DATA = -2147483642
    HKEY_CURRENT_CONFIG = -2147483643
    HKEY_PERFORMANCE_DATA = -2147483644
    HKEY_USERS = -2147483645
    HKEY_LOCAL_MACHINE = -2147483646
    HKEY_CURRENT_USER = -2147483647
    HKEY_CLASSES_ROOT = -2147483648#
End Enum

Public Enum REG_TYPE
    REG_NONE = 0
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD = 4
    REG_DWORD_LITTLE_ENDIAN = 4
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7
    REG_RESOURCE_LIST = 8
    REG_FULL_RESOURCE_DESCRIPTOR = 9
    REG_RESOURCE_REQUIREMENTS_LIST = 10
    REG_QWORD = 11     ' 64-bit number
    REG_QWORD_LITTLE_ENDIAN = 11
End Enum

Private Const PTSUBKEY = "Software\KarenWare\Power Tools\"
Private Const OLDPTSUBKEY = "Software\Windows Magazine\Power Tools\"
Private Const OLDPTSUBKEY2 = "Software\Windows Magazine\Power Tools"
Private Const OLDPTSUBKEY3 = "Software\Windows Magazine"
Private Const OLDPTROOT = "Software\Windows Magazine"

Private Const REG_LASTRUN = "LastRun"
Private Const REG_APPPATH = "AppPath"
Private Const REG_HOME = "Home"

Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_CREATED_NEW_KEY = &H1
Private Const REG_OPENED_EXISTING_KEY = &H2

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_CREATE_LINK = &H20
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

#If REG_FILE_ASSOC Then
Private Enum SHCN_EVENT
    SHCNE_ALLEVENTS = &H7FFFFFFF
    SHCNE_ASSOCCHANGED = &H8000000
    SHCNE_ATTRIBUTES = &H800&
    SHCNE_CREATE = &H2&
    SHCNE_DELETE = &H4&
    SHCNE_DISKEVENTS = &H2381F
    SHCNE_DRIVEADD = &H100&
    SHCNE_DRIVEADDGUI = &H10000
    SHCNE_DRIVEREMOVED = &H80&
    SHCNE_EXTENDED_EVENT = &H4000000
    SHCNE_FREESPACE = &H40000
    SHCNE_GLOBALEVENTS = &HC0581E0
    SHCNE_INTERRUPT = &H80000000
    SHCNE_MEDIAINSERTED = &H20&
    SHCNE_MEDIAREMOVED = &H40&
    SHCNE_MKDIR = &H8&
    SHCNE_NETSHARE = &H200&
    SHCNE_NETUNSHARE = &H400&
    SHCNE_RENAMEFOLDER = &H20000
    SHCNE_RENAMEITEM = &H1&
    SHCNE_RMDIR = &H10&
    SHCNE_SERVERDISCONNECT = &H4000&
    SHCNE_UPDATEDIR = &H1000&
    SHCNE_UPDATEIMAGE = &H8000&
    SHCNE_UPDATEITEM = &H2000&
End Enum

Private Enum SHCN_FLAG
    SHCNF_ACCEPT_INTERRUPTS = &H1
    SHCNF_ACCEPT_NON_INTERRUPTS = &H2
    SHCNF_DWORD = &H3
    SHCNF_FLUSH = &H1000
    SHCNF_FLUSHNOWAIT = &H2000
    SHCNF_IDLIST = &H0
    SHCNF_NO_PROXY = &H8000
    SHCNF_PATHA = &H1
    SHCNF_PATHW = &H5
    SHCNF_PRINTERA = &H2
    SHCNF_PRINTERW = &H6
    SHCNF_TYPE = &HFF&
End Enum
#End If ' REG_FILE_ASSOC

Private Const STARTF_USESHOWWINDOW = &H1
Private Const SW_NORMAL = 1
Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWMINNOACTIVE = 7

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

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

' Startup infomation used to position new window
Private Type STARTUPINFO
   cb As Long
   lpReserved As Long
   lpDesktop As Long
   lpTitle As Long
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

#If REG_FILE_ASSOC Then
Private Declare Function SHChangeNotify Lib "Shell32.dll" ( _
    ByVal wEventID As SHCN_EVENT, _
    ByVal uFlags As SHCN_FLAG, _
    ByVal dwItem1 As Long, _
    ByVal dwItem2 As Long) As Long
#End If ' REG_FILE_ASSOC

Private Declare Sub GetStartupInfoA Lib "kernel32" ( _
    lpStartupInfo As STARTUPINFO)

Private Declare Function RegOpenKeyExA Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As Long) As Long

Private Declare Function RegQueryValueExA Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Any, _
    dwSize As Long) As Long
    ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Private Declare Function RegQueryValueExABinary _
    Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    ByVal lpData As Long, _
    dwSize As Long) As Long

Private Declare Function RegSetValueExA Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal dwReserved As Long, _
    ByVal dwType As Long, _
    lpValue As Any, _
    ByVal dwSize As Long) As Long

Private Declare Function RegDeleteKeyA Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValueA Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" ( _
    ByVal hKey As Long) As Long

Private Declare Function RegQueryInfoKeyA Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal lpClass As String, _
    lpcbClass As Long, _
    ByVal lpReserved As Long, _
    lpcSubKeys As Long, _
    lpcbMaxSubKeyLen As Long, _
    lpcbMaxClassLen As Long, _
    lpcValues As Long, _
    lpcbMaxValueNameLen As Long, _
    lpcbMaxValueLen As Long, _
    lpcbSecurityDescriptor As Long, _
    lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegEnumValueA Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal dwIndex As Long, _
    ByVal lpValueName As String, _
    lpcbValueName As Long, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Byte, _
    lpcbData As Long) As Long

Private Declare Function RegConnectRegistryA Lib "advapi32.dll" ( _
    ByVal lpMachineName As String, _
    ByVal hKey As Long, _
    phkResult As Long) As Long

Private Declare Function RegCreateKeyExA Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal reserved As Long, _
    ByVal lpClass As String, _
    ByVal dwOptions As Long, _
    ByVal samDesired As Long, _
    lpSecurityAttributes As SECURITY_ATTRIBUTES, _
    phkResult As Long, _
    lpdwDisposition As Long) As Long

Private Declare Function RegEnumKeyA Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal dwIndex As Long, _
    ByVal lpName As String, _
    ByVal cbName As Long) As Long

Private Declare Function RegEnumKeyExA Lib "advapi32.dll" ( _
    ByVal hKey As Long, _
    ByVal dwIndex As Long, _
    ByVal lpName As String, _
    lpcbName As Long, _
    ByVal lpReserved As Long, _
    ByVal lpClass As String, _
    lpcbClass As Long, _
    lpftLastWriteTime As FILETIME) As Long

Private Declare Function FileTimeToSystemTime Lib "kernel32" ( _
    lpFileTime As FILETIME, _
    lpSystemTime As SYSTEMTIME) As Long

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" ( _
    lpFileTime As FILETIME, _
    lpLocalFileTime As FILETIME) As Long

'Private Declare Function RegLoadKeyA Lib "advapi32.dll" ( _
'    ByVal hKey As Long, _
'    ByVal lpSubKey As String, _
'    ByVal lpFile As String) As Long

'Private Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" ( _
'    ByVal hKey As Long, _
'    ByVal bWatchSubtree As Long, _
'    ByVal dwNotifyFilter As Long, _
'    ByVal hEvent As Long, _
'    ByVal fAsynchronus As Long) As Long

'Private Declare Function RegOpenKey Lib "advapi32.dll" ( _
'    ByVal hKey As Long, _
'    ByVal lpSubKey As String, _
'    phkResult As Long) As Long

'Private Declare Function OSRegQueryValueA Lib "advapi32.dll" Alias "RegQueryValueA" ( _
'    ByVal hKey As Long, _
'    ByVal lpSubKey As String, _
'    ByVal lpValue As String, _
'    lpcbValue As Long) As Long

'private declare Function RegReplaceKey Lib "advapi32.dll" _
'    Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
'    ByVal lpNewFile As String, ByVal lpOldFile As String) As Long

'private declare Function RegRestoreKey Lib "advapi32.dll" _
'    Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, _
'    ByVal dwFlags As Long) As Long

'local variable(s) to hold property value(s)
Private mvarDefaultKey As String
Private mvarDefaultHKEY As REG_HKEY
#If REG_FILE_ASSOC Then
Public Sub FileAssocAdd(Ext As String, Desc As String, Optional ExeFullName As String = "", Optional Cmd As String = "%1", Optional Action As String = "Open")
    Dim exepath As String
    Dim exe As String
    Dim FileClass As String
    Dim ProperExe As String
    Dim ProperCmd As String

    If Len(ExeFullName) > 0 Then
        exe = ExeFullName
    Else
        exepath = App.Path
        If Right(exepath, 1) <> "\" Then exepath = exepath & "\"
        exe = exepath & App.EXEName
        If LCase(Right(exe, 4)) <> ".exe" Then exe = exe & ".exe"
        If Not ApiFileExists(exe) Then
            exe = exepath & "PT" & App.EXEName
            If LCase(Right(exe, 4)) <> ".exe" Then exe = exe & ".exe"
        End If
    End If

    If Left(Ext, 1) <> "." Then Ext = "." & Ext

    FileClass = "KarenWare." & Replace(Ext, ".", "")

    If Left(Cmd, 1) <> vbQuote Then
        ProperCmd = vbQuote & Cmd & vbQuote
    Else
        ProperCmd = Cmd
    End If
        
    If Left(exe, 1) <> vbQuote Then
        ProperExe = vbQuote & exe & vbQuote
    Else
        ProperExe = exe
    End If

    WriteString HKEY_CLASSES_ROOT, FileClass, "", Desc
    WriteString HKEY_CLASSES_ROOT, FileClass & "\DefaultIcon", "", ProperExe
    WriteString HKEY_CLASSES_ROOT, FileClass & "\shell\" & Action & "\command", "", ProperExe & " " & ProperCmd
    WriteString HKEY_CLASSES_ROOT, Ext, "", FileClass
    WriteString HKEY_CLASSES_ROOT, Ext, "AppKey", APPKEY

    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST Or SHCNF_FLUSHNOWAIT, 0, 0
End Sub
Public Sub FileAssocDel(Ext As String)
    Dim FileClass As String
    
    If Left(Ext, 1) <> "." Then Ext = "." & Ext
    FileClass = "KarenWare." & Replace(Ext, ".", "")

    DeleteKeyAbs FileClass, HKEY_CLASSES_ROOT
    DeleteKeyAbs Ext, HKEY_CLASSES_ROOT

    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST Or SHCNF_FLUSHNOWAIT, 0, 0
End Sub
Public Function FileAssocOwned(Ext As String, Optional TestAppKey As String = APPKEY) As Boolean
    Dim OwnerAppKey As String

    If Left(Ext, 1) <> "." Then Ext = "." & Ext
    OwnerAppKey = Read(HKEY_CLASSES_ROOT, Ext, "AppKey", "")

    If StrComp(OwnerAppKey, TestAppKey, vbTextCompare) = 0 Then
        FileAssocOwned = True
    Else
        FileAssocOwned = False
    End If
End Function
Public Function FileAssocExists(Ext As String) As Boolean
    Dim FileClass As String

    If Left(Ext, 1) <> "." Then Ext = "." & Ext
    FileClass = Read(HKEY_CLASSES_ROOT, Ext, "", "")

    If Len(FileClass) > 0 Then
        FileAssocExists = True
    Else
        FileAssocExists = False
    End If
End Function
#End If ' REG_FILE_ASSOC
Public Function GetAppSubKey(Optional AppID As String = APPKEY)
    GetAppSubKey = PTSUBKEY & AppID
End Function
Public Function GetFormSubKey(Frm As Form, Optional AppID As String = APPKEY)
    Dim ID As String

    ID = Frm.Name
    If Left(ID, 3) = "frm" Then ID = Mid(ID, 4)

    GetFormSubKey = PTSUBKEY & AppID & "\" & ID
End Function
#If Not (USE_PTFORM = 1) Then
Public Sub SaveFormSize(Frm As Form, Optional AppID As String = APPKEY)
    Dim rkey As String

    rkey = GetFormSubKey(Frm, AppID)

    WriteLong HKEY_CURRENT_USER, rkey, "State", Frm.WindowState

    If Frm.WindowState = vbMinimized Then Exit Sub

    If Frm.WindowState = vbNormal Then
        WriteLong HKEY_CURRENT_USER, rkey, "Width", Frm.Width
        WriteLong HKEY_CURRENT_USER, rkey, "Height", Frm.Height
    End If
End Sub
Public Function GetFormSize(Frm As Form, Optional AppID As String = APPKEY) As Boolean
    Dim rkey As String
    Dim w As Long
    Dim h As Long
    Dim State As Long
    Dim si As STARTUPINFO

    GetFormSize = False

    rkey = GetFormSubKey(Frm, AppID)

    w = Read(HKEY_CURRENT_USER, rkey, "Width", Frm.Width)
    h = Read(HKEY_CURRENT_USER, rkey, "Height", Frm.Height)
    On Error Resume Next
    If w > 0 Then Frm.Width = w
    If h > 0 Then Frm.Height = h

    si.cb = Len(si)
    Call GetStartupInfoA(si)

    If Not (Frm Is Forms(0)) Then ' not startup form
        State = Read(HKEY_CURRENT_USER, rkey, "State", vbNormal)
    ElseIf si.dwFlags And STARTF_USESHOWWINDOW Then
        If si.wShowWindow = SW_SHOWMAXIMIZED Then
            State = vbMaximized
        ElseIf si.wShowWindow = SW_SHOWMINIMIZED Then
            State = vbMinimized
        ElseIf si.wShowWindow = SW_SHOWMINNOACTIVE Then
            State = vbMinimized
        Else
            State = Read(HKEY_CURRENT_USER, rkey, "State", vbNormal)
        End If
    End If
    Frm.WindowState = State
    If State = vbMinimized Then GetFormSize = True
    Err.Clear
End Function
Public Sub SaveFormPos(Frm As Form, Optional AppID As String = APPKEY)
    Dim rkey As String

    If Frm.WindowState <> vbNormal Then Exit Sub

    rkey = GetFormSubKey(Frm, AppID)

    WriteLong HKEY_CURRENT_USER, rkey, "Top", Frm.Top
    WriteLong HKEY_CURRENT_USER, rkey, "Left", Frm.Left
End Sub
Public Sub GetFormPos(Frm As Form, Optional AppID As String = APPKEY)
    Dim rkey As String

    rkey = GetFormSubKey(Frm, AppID)
    On Error Resume Next
    Frm.Top = Read(HKEY_CURRENT_USER, rkey, "Top", (Screen.Height - Frm.Height) / 2)
    Frm.Left = Read(HKEY_CURRENT_USER, rkey, "Left", (Screen.Width - Frm.Width) / 2)
    If (Frm.Top < 0) Or (Frm.Top > (Screen.Height - 1000)) Then
        Frm.Top = (Screen.Height - Frm.Height) / 2
    End If
    If (Frm.Left < 0) Or (Frm.Left > (Screen.Width - 1000)) Then
        Frm.Left = (Screen.Width - Frm.Width) / 2
    End If
    Err.Clear
End Sub
#End If ' Not USE_PTFORM
Public Sub SaveFormValue(Frm As Form, DataName As String, DataValue As Variant, Optional AppID As String = APPKEY)
    Dim rkey As String

    rkey = GetFormSubKey(Frm, AppID)

    WriteValue DataName, DataValue, rkey
End Sub
Public Function GetFormValue(Frm As Form, DataName As String, DefaultValue As Variant, Optional AppID As String = APPKEY) As Variant
    Dim rkey As String

    rkey = GetFormSubKey(Frm, AppID)

    GetFormValue = ReadValue(DataName, DefaultValue, rkey)
End Function
Public Function GetValuesApp(RelativeSubKey As String, Optional hKey As REG_HKEY = HKEY_DEFAULT) As REG_VAL_ARRAY
    Set GetValuesApp = GetValuesAbs(EffectiveKey(RelativeSubKey), hKey)
End Function
Public Function GetValuesAbs(AbsoluteSubKey As String, Optional hKey As REG_HKEY = HKEY_DEFAULT) As REG_VAL_ARRAY
    Dim hSubKey As Long
    Dim i As Long
    Dim s As String
    Dim Result As Long
    Dim ValueName As String
    Dim LastWriteTime As FILETIME
    Dim SubKeyCnt As Long
    Dim MaxSubKeyLen As Long
    Dim MaxClassLen As Long
    Dim ValueCnt As Long
    Dim MaxValueNameLen As Long
    Dim MaxValueLen As Long
    Dim SecDescLen As Long
    Dim DataType As REG_TYPE
    Dim DataSize As Long
    Dim ba() As Byte
    Dim v As Variant
    Dim inx As Long
    Dim va As REG_VAL_ARRAY
    Dim EffhKey As REG_HKEY

    EffhKey = EffectiveHKEY(hKey)
    Result = RegOpenKeyExA(EffhKey, AbsoluteSubKey, 0, KEY_ENUMERATE_SUB_KEYS Or KEY_READ, hSubKey)

    Set va = New REG_VAL_ARRAY
    If Result <> ERROR_SUCCESS Then
        Set GetValuesAbs = va
        Exit Function
    End If

    Result = RegQueryInfoKeyA(hSubKey, vbNullString, 0, 0, SubKeyCnt, MaxSubKeyLen, MaxClassLen, ValueCnt, MaxValueNameLen, MaxValueLen, SecDescLen, LastWriteTime)

    ValueName = String(MaxValueNameLen + 1, 0)
    ReDim ba(MaxValueLen + 1)
    DataSize = MaxValueLen
    Result = RegEnumValueA(hSubKey, va.Count, ValueName, Len(ValueName), _
        0, DataType, ba(0), DataSize)
    While Result = ERROR_SUCCESS
        inx = va.ValueAdd(ApiTextStrip(ValueName), DataType)
        va.Value(inx).Tag = ""
        Select Case DataType
            Case REG_SZ:
                s = ""
                i = 0
                While i < DataSize + 1
                    s = s & Chr(ba(i))
                    i = i + 1
                Wend
                va.Value(inx).Data = ApiTextStrip(s)
                va.Value(inx).Size = Len(s)
            Case REG_EXPAND_SZ:
                s = ""
                i = 0
                While i < (DataSize * 2) + 1
                    s = s & Chr(ba(i))
                    i = i + 1
                Wend
                va.Value(inx).Data = s
                va.Value(inx).Size = (DataSize * 2)
            Case REG_MULTI_SZ:
                s = ""
                i = 0
                While i < (DataSize * 2) + 1
                    s = s & Chr(ba(i))
                    i = i + 1
                Wend
                va.Value(inx).Data = s
                va.Value(inx).Size = (DataSize * 2)
            Case REG_BINARY:
                ReDim Preserve ba(DataSize - 1)
                va.Value(inx).Data = ba
                va.Value(inx).Size = DataSize
            Case REG_DWORD:
                i = ba(0)
                i = i + (CLng(ba(1)) * 256)
                i = i + (CLng(ba(2)) * 256 * 256)
                s = Hex(i)
                If Len(s) < 6 Then s = String(6 - Len(s), "0") & s
                'i = (CLng(ba(3)) * 256 * 256 * 256)
                s = "&h" & Hex(ba(3)) & s
                va.Value(inx).Data = Val(s)
                va.Value(inx).Size = 4
            Case Else:
                va.Value(inx).Data = ba
                va.Value(inx).Size = DataSize
        End Select

        ValueName = String(MaxValueNameLen + 1, 0)
        ReDim ba(MaxValueLen + 1)
        DataSize = MaxValueLen
        Result = RegEnumValueA(hSubKey, va.Count, ValueName, Len(ValueName), _
            0, DataType, ba(0), DataSize)
    Wend

    RegCloseKey hSubKey
    Set GetValuesAbs = va
End Function
Public Function GetSubKeysApp(RelativeSubKey As String, Optional hKey As REG_HKEY = HKEY_DEFAULT) As REG_VAL_ARRAY
    Set GetSubKeysApp = GetSubKeysAbs(EffectiveKey(RelativeSubKey), hKey)
End Function
Public Function GetSubKeysAbs(AbsoluteSubKey As String, Optional hKey As REG_HKEY = HKEY_DEFAULT) As REG_KEY_ARRAY
    Dim s As String
    Dim hSubKey As Long
    Dim i As Long
    Dim Result As Long
    Dim ft As FILETIME
    Dim SubKeyCnt As Long
    Dim MaxSubKeyLen As Long
    Dim MaxClassLen As Long
    Dim ValueCnt As Long
    Dim MaxValueNameLen As Long
    Dim MaxValueLen As Long
    Dim SecDescLen As Long
    Dim inx As Long
    Dim ka As REG_KEY_ARRAY
    Dim EffhKey As REG_HKEY

    EffhKey = EffectiveHKEY(hKey)
    Result = RegOpenKeyExA(EffhKey, AbsoluteSubKey, 0, KEY_ENUMERATE_SUB_KEYS Or KEY_READ, hSubKey)

    Set ka = New REG_KEY_ARRAY
    If Result <> ERROR_SUCCESS Then
        Set GetSubKeysAbs = ka
        Exit Function
    End If

    Result = RegQueryInfoKeyA(hSubKey, vbNullString, 0, 0, SubKeyCnt, MaxSubKeyLen, MaxClassLen, ValueCnt, MaxValueNameLen, MaxValueLen, SecDescLen, ft)
    If (Result <> ERROR_SUCCESS) Or (SubKeyCnt <= 0) Then
        Set GetSubKeysAbs = ka
        Exit Function
    End If

    For i = 1 To SubKeyCnt
        inx = ka.KeyAdd()
        s = String(MaxSubKeyLen + 1, 0)
        Result = RegEnumKeyExA(hSubKey, i - 1, s, Len(s), 0, vbNullString, 0, ft)
        ka.Key(inx).Name = ApiTextStrip(s)
        ka.Key(inx).LastWriteTime = FileTimeToDate(ft)
        ka.Key(inx).Tag = ""
    Next i

    RegCloseKey hSubKey
    Set GetSubKeysAbs = ka
End Function
Private Function FileTimeToDate(ft As FILETIME) As Date
    Dim d As Date
    Dim lft As FILETIME
    Dim st As SYSTEMTIME

    FileTimeToLocalFileTime ft, lft
    FileTimeToSystemTime lft, st
    d = DateSerial(st.wYear, st.wMonth, st.wDay) + TimeSerial(st.wHour, st.wMinute, st.wSecond)
    FileTimeToDate = d
End Function
Private Function Read(hKey As REG_HKEY, SubKey As String, DataName As String, DefaultData As Variant) As Variant
    Dim hKeyResult As Long
    Dim lData As Long
    Dim sData As String
    Dim DataType As Long
    Dim DataSize As Long
    Dim Result As Long
    Dim buf() As Byte
    Dim s As String

    Read = DefaultData
    Result = RegOpenKeyExA(hKey, SubKey, 0, KEY_QUERY_VALUE, hKeyResult)
    If Result <> ERROR_SUCCESS Then Exit Function

    Result = RegQueryValueExA(hKeyResult, DataName, 0&, DataType, ByVal 0, DataSize)
    If (Result <> ERROR_SUCCESS) Or (DataSize <= 0) Then
        RegCloseKey hKeyResult
        Exit Function
    End If

    Select Case DataType
        Case REG_SZ, REG_EXPAND_SZ:
            sData = Space(DataSize)
            Result = RegQueryValueExA(hKeyResult, DataName, 0&, DataType, ByVal sData, Len(sData))
            If Result = ERROR_SUCCESS Then s = ApiTextStrip(RTrim(sData))
            If VarType(DefaultData) = vbSingle Then
                Read = CVar(CSng(s))
            ElseIf VarType(DefaultData) = vbDouble Then
                Read = CVar(CDbl(s))
            ElseIf VarType(DefaultData) = vbCurrency Then
                Read = CVar(CCur(s))
            ElseIf VarType(DefaultData) = vbDecimal Then
                Read = CVar(CDec(s))
            Else
                Read = CVar(s)
            End If
        Case REG_DWORD
            Result = RegQueryValueExA(hKeyResult, DataName, 0&, DataType, lData, 4)
            If Result = ERROR_SUCCESS Then Read = CVar(lData)
        Case REG_MULTI_SZ:
            sData = Space(DataSize)
            Result = RegQueryValueExA(hKeyResult, DataName, 0&, DataType, ByVal sData, Len(sData))
            If Result = ERROR_SUCCESS Then Read = CVar(sData)
        Case REG_BINARY:
            ReDim buf(DataSize - 1)
            Result = RegQueryValueExABinary(hKeyResult, DataName, 0, DataType, VarPtr(buf(0)), DataSize)
            If Result = ERROR_SUCCESS Then Read = buf
        End Select

    RegCloseKey hKeyResult
End Function
Private Function EffectiveHKEY(hKey As REG_HKEY) As REG_HKEY
    If hKey = HKEY_DEFAULT Then
        EffectiveHKEY = DefaultHKEY
    Else
        EffectiveHKEY = hKey
    End If
End Function
Private Function EffectiveKey(Key As String) As String
    If Len(Key) > 0 Then
        EffectiveKey = Key
    Else
        EffectiveKey = mvarDefaultKey
    End If
End Function
Public Function ReadValue(ValueName As String, DefaultValue As Variant, Optional Key As String = "", Optional hKey As REG_HKEY = HKEY_DEFAULT) As Variant
    ReadValue = Read(EffectiveHKEY(hKey), EffectiveKey(Key), ValueName, DefaultValue)
End Function
Public Function WriteValue(Name As String, Data As Variant, Optional Key As String = "", Optional hKey As REG_HKEY = HKEY_DEFAULT) As REG_ERROR
    Select Case VarType(Data)
        Case vbArray Or vbByte:
            ' not supported
            WriteBinary EffectiveHKEY(hKey), EffectiveKey(Key), Name, Data

        Case vbUserDefinedType:
            WriteUdt EffectiveHKEY(hKey), EffectiveKey(Key), Name, Data

        Case vbString:
            WriteString EffectiveHKEY(hKey), EffectiveKey(Key), Name, CStr(Data)

        Case vbLong:
            WriteLong EffectiveHKEY(hKey), EffectiveKey(Key), Name, CLng(Data)

        Case vbInteger:
            WriteLong EffectiveHKEY(hKey), EffectiveKey(Key), Name, CLng(Data)

        Case vbSingle, vbDouble, vbCurrency, vbDecimal:
            WriteString EffectiveHKEY(hKey), EffectiveKey(Key), Name, CStr(Data)
        
        Case vbByte:
            WriteLong EffectiveHKEY(hKey), EffectiveKey(Key), Name, CLng(Data)

        Case vbBoolean:
            WriteLong EffectiveHKEY(hKey), EffectiveKey(Key), Name, CLng(Data)

        Case Else:
            ' not supported

    End Select

End Function
Private Function WriteString(hKey As REG_HKEY, Key As String, Name As String, Data As String) As REG_ERROR
    Dim sa As SECURITY_ATTRIBUTES
    Dim hKeyResult As Long
    Dim lDisposition As Long
    Dim Result As Long

    sa.nLength = Len(sa)
    sa.lpSecurityDescriptor = 0
    sa.bInheritHandle = False
    
    Result = RegCreateKeyExA(hKey, _
        Key, 0, vbNullString, _
        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
        sa, hKeyResult, lDisposition)
    
    If Data <= "" Then Data = ""  ' bug fix?
    If (Result = ERROR_SUCCESS) Or (Result = REG_CREATED_NEW_KEY) Or (Result = REG_OPENED_EXISTING_KEY) Then
        Result = RegSetValueExA(hKeyResult, Name, 0&, _
            REG_SZ, ByVal Data, Len(Data))
        RegCloseKey hKeyResult
    End If
    WriteString = Result
End Function
Private Function WriteLong(hKey As REG_HKEY, SubKey As String, DataName As String, DataValue As Long) As Long
    Dim sa As SECURITY_ATTRIBUTES
    Dim hKeyResult As Long
    Dim lDisposition As Long
    Dim Result As Long

    sa.nLength = Len(sa)
    sa.lpSecurityDescriptor = 0
    sa.bInheritHandle = False

    Result = RegCreateKeyExA(hKey, _
        SubKey, 0, "", _
        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
        sa, hKeyResult, lDisposition)

    If (Result = ERROR_SUCCESS) Or (Result = REG_CREATED_NEW_KEY) Or (Result = REG_OPENED_EXISTING_KEY) Then
        Result = RegSetValueExA(hKeyResult, DataName, 0&, _
            REG_DWORD, DataValue, 4)
        RegCloseKey hKeyResult
    End If
    WriteLong = Result
End Function
Private Function WriteBinary(hKey As REG_HKEY, SubKey As String, DataName As String, DataValue As Variant) As Long
    Dim sa As SECURITY_ATTRIBUTES
    Dim hKeyResult As Long
    Dim lDisposition As Long
    Dim Result As Long
    Dim buf() As Byte

    sa.nLength = Len(sa)
    sa.lpSecurityDescriptor = 0
    sa.bInheritHandle = False

    Result = RegCreateKeyExA(hKey, _
        SubKey, 0, "", _
        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
        sa, hKeyResult, lDisposition)

    If (Result = ERROR_SUCCESS) Or (Result = REG_CREATED_NEW_KEY) Or (Result = REG_OPENED_EXISTING_KEY) Then
        buf = DataValue
        Result = RegSetValueExA(hKeyResult, DataName, 0&, _
            REG_BINARY, buf(0), UBound(buf) + 1)
        RegCloseKey hKeyResult
    End If
    WriteBinary = Result
End Function
Private Function WriteUdt(hKey As REG_HKEY, SubKey As String, DataName As String, DataValue As Variant) As Long
    Dim sa As SECURITY_ATTRIBUTES
    Dim hKeyResult As Long
    Dim lDisposition As Long
    Dim Result As Long

    sa.nLength = Len(sa)
    sa.lpSecurityDescriptor = 0
    sa.bInheritHandle = False

    Result = RegCreateKeyExA(hKey, _
        SubKey, 0, "", _
        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
        sa, hKeyResult, lDisposition)

    If (Result = ERROR_SUCCESS) Or (Result = REG_CREATED_NEW_KEY) Or (Result = REG_OPENED_EXISTING_KEY) Then
        Result = RegSetValueExA(hKeyResult, DataName, 0&, _
            REG_BINARY, DataValue, LenB(DataValue))
        RegCloseKey hKeyResult
    End If
    WriteUdt = Result
End Function
Public Function DeleteValue(ValueName As String, Optional Key As String = "", Optional hKey As REG_HKEY = HKEY_DEFAULT) As Long
    Dim Result As Long
    Dim hKeyResult As Long

    Result = RegOpenKeyExA(EffectiveHKEY(hKey), EffectiveKey(Key), 0, KEY_WRITE, hKeyResult)
    If Result <> ERROR_SUCCESS Then Exit Function

    Result = RegDeleteValueA(hKeyResult, ValueName)
    DeleteValue = Result
    RegCloseKey hKeyResult
End Function
Public Function AddKey(SubKey As String, Optional hKey As REG_HKEY = HKEY_DEFAULT) As REG_ERROR
    Dim Result As Long
    Dim hnd As Long
    Dim sa As SECURITY_ATTRIBUTES
    Dim EffhKey As REG_HKEY

    sa.nLength = Len(sa)
    sa.lpSecurityDescriptor = 0
    sa.bInheritHandle = False

    EffhKey = EffectiveHKEY(hKey)
    RegCreateKeyExA EffhKey, SubKey, 0, vbNull, REG_OPTION_NON_VOLATILE, KEY_CREATE_SUB_KEY, sa, hnd, vbNull
    RegCloseKey hnd
End Function
Public Function DeleteKeyApp(RelativeSubKey As String, Optional hKey As REG_HKEY = HKEY_DEFAULT) As REG_ERROR
    DeleteKeyApp = DeleteKeyAbs(EffectiveKey(RelativeSubKey), hKey)
End Function
Public Function DeleteKeyAbs(AbsoluteSubKey As String, Optional hKey As REG_HKEY = HKEY_DEFAULT) As REG_ERROR
    Dim Result As Long
    Dim ka As REG_KEY_ARRAY
    Dim i As Long
    Dim EffhKey As REG_HKEY

    EffhKey = EffectiveHKEY(hKey)
    Set ka = GetSubKeysAbs(AbsoluteSubKey, EffhKey)
    For i = ka.LowerBound To ka.UpperBound
        Result = DeleteKeyAbs(AbsoluteSubKey & "\" & ka.Key(i).Name, EffhKey)
        If Result <> ERROR_SUCCESS Then
            DeleteKeyAbs = Result
            Exit Function
        End If
    Next i

    Result = RegDeleteKeyA(EffhKey, AbsoluteSubKey)
    DeleteKeyAbs = Result
End Function
Public Function GetAppPath() As String
    Dim AppPath As String

    AppPath = ReadValue(REG_APPPATH, "")
    GetAppPath = AppPath
End Function
Public Function GetHome() As String
    Dim Home As String

    Home = ReadValue(REG_HOME, "")
    GetHome = Home
End Function
Public Sub Home()
    Dim Home As String
    Dim d As Date

    Home = App.Path
    If Right(Home, 1) <> "\" Then Home = Home & "\"
    WriteValue REG_APPPATH, Home

    Home = Home & App.EXEName
    If LCase(Right(Home, 4)) <> ".exe" Then Home = Home & ".exe"
    WriteValue REG_HOME, Home

    d = Now()
    WriteValue REG_LASTRUN, CStr(CDbl(d))
End Sub
Public Function GetLastRun() As Date
    Dim s As String
    Dim d As Date

    s = ReadValue(REG_LASTRUN, "")
    If IsNumeric(s) Then
        d = CDate(CDbl(s))
    Else
        d = 0
    End If
    GetLastRun = d
End Function
Public Sub Migrate(Optional OldKey As String = "", Optional NewKey As String = "")
    Dim va As REG_VAL_ARRAY
    Dim ka As REG_KEY_ARRAY
    Dim i As Long

    OldKey = Trim(OldKey)
    If Len(OldKey) <= 0 Then OldKey = OLDPTSUBKEY & APPKEY
    NewKey = Trim(NewKey)
    If Len(NewKey) <= 0 Then NewKey = PTSUBKEY & APPKEY

    If Right(OldKey, 1) = "\" Then OldKey = Left(OldKey, Len(OldKey) - 1)
    If Right(NewKey, 1) = "\" Then NewKey = Left(NewKey, Len(NewKey) - 1)

    Set va = GetValuesAbs(OldKey)
    For i = va.LowerBound To va.UpperBound
        ' copy and delete values
        Select Case va.Value(i).DataType
            Case REG_DWORD:
                WriteLong HKEY_CURRENT_USER, NewKey, va.Value(i).Name, CLng(va.Value(i).Data)
            Case Else
                WriteString HKEY_CURRENT_USER, NewKey, va.Value(i).Name, CStr(va.Value(i).Data)
        End Select
        DeleteValue va.Value(i).Name, NewKey
    Next i

    Set ka = GetSubKeysAbs(OldKey)
    For i = ka.LowerBound To ka.UpperBound
        Migrate OldKey & "\" & ka.Key(i).Name, NewKey & "\" & ka.Key(i).Name
        RegDeleteKeyA HKEY_CURRENT_USER, OldKey & "\" & ka.Key(i).Name
    Next i
    RegDeleteKeyA HKEY_CURRENT_USER, OldKey

    DeleteKeyAbs OLDPTROOT

'    Set ka = GetSubKeys(OLDPTSUBKEY2)
'    If ka.Cnt > 0 Then ' remove Software\Windows Magazine\Power Tools
'        DeleteKey OLDPTSUBKEY2
'    Else
'        Exit Sub
'    End If
'
'    Set ka = GetSubKeys(OLDPTSUBKEY3)
'    If ka.Cnt > 0 Then  ' remove Software\Windows Magazine
'        DeleteKey OLDPTSUBKEY3
'    End If
End Sub
Public Sub SortKeysByTag(ByRef ka As REG_KEY_ARRAY)
    Dim swap As REG_KEY_ENTRY
    Dim i As Long
    Dim lim As Long
    Dim Sorted As Boolean

    If ka.Count <= 1 Then Exit Sub

    lim = ka.UpperBound
    Do
        Sorted = True
        For i = ka.LowerBound + 1 To lim
            If StrComp(ka.Key(i).Tag, ka.Key(i - 1).Tag, vbTextCompare) < 0 Then
                Set swap = ka.Key(i)
                Set ka.Key(i) = ka.Key(i - 1)
                Set ka.Key(i - 1) = swap
                Sorted = False
            End If
        Next i
    Loop While Sorted = False
End Sub
Public Sub SortDatumByTag(ByRef va As REG_VAL_ARRAY)
    Dim swap As REG_VAL_ENTRY
    Dim i As Long
    Dim lim As Long
    Dim Sorted As Boolean

    If va.Count <= 1 Then Exit Sub

    lim = va.UpperBound
    Do
        Sorted = True
        For i = va.LowerBound + 1 To lim
            If StrComp(va.Value(i).Tag, va.Value(i - 1).Tag, vbTextCompare) < 0 Then
                Set swap = va.Value(i)
                Set va.Value(i) = va.Value(i - 1)
                Set va.Value(i - 1) = swap
                Sorted = False
            End If
        Next i
    Loop While Sorted = False
End Sub
Public Sub SortKeysByName(ByRef ka As REG_KEY_ARRAY)
    Dim swap As REG_KEY_ENTRY
    Dim i As Long
    Dim lim As Long
    Dim Sorted As Boolean

    If ka.Count <= 1 Then Exit Sub

    lim = ka.UpperBound
    Do
        Sorted = True
        For i = ka.LowerBound + 1 To lim
            If StrComp(ka.Key(i).Name, ka.Key(i - 1).Name, vbTextCompare) < 0 Then
                Set swap = ka.Key(i)
                Set ka.Key(i) = ka.Key(i - 1)
                Set ka.Key(i - 1) = swap
                Sorted = False
            End If
        Next i
    Loop While Sorted = False
End Sub
Public Sub SortDatumByName(ByRef va As REG_VAL_ARRAY)
    Dim swap As REG_VAL_ENTRY
    Dim i As Long
    Dim lim As Long
    Dim Sorted As Boolean

    If va.Count <= 1 Then Exit Sub

    lim = va.UpperBound
    Do
        Sorted = True
        For i = va.LowerBound + 1 To lim
            If StrComp(va.Value(i).Name, va.Value(i - 1).Name, vbTextCompare) < 0 Then
                Set swap = va.Value(i)
                Set va.Value(i) = va.Value(i - 1)
                Set va.Value(i - 1) = swap
                Sorted = False
            End If
        Next i
    Loop While Sorted = False
End Sub
Public Property Let DefaultHKEY(ByVal NewValue As REG_HKEY)
    mvarDefaultHKEY = NewValue
End Property
Public Property Get DefaultHKEY() As REG_HKEY
    DefaultHKEY = mvarDefaultHKEY
End Property
Public Property Let DefaultKey(ByVal NewValue As String)
    mvarDefaultKey = NewValue
End Property
Public Property Get DefaultKey() As String
    DefaultKey = mvarDefaultKey
End Property
Private Sub Class_Initialize()
    DefaultHKEY = HKEY_CURRENT_USER
    DefaultKey = PTSUBKEY & APPKEY
End Sub
