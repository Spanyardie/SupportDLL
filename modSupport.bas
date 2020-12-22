Attribute VB_Name = "modSupport"
    Option Explicit


    Public Const REG_SZ As Long = 1
    Public Const REG_DWORD As Long = 4
    
    Public Const HKEY_CLASSES_ROOT = &H80000000
    Public Const HKEY_CURRENT_USER = &H80000001
    Public Const HKEY_LOCAL_MACHINE = &H80000002
    Public Const HKEY_USERS = &H80000003
    
    Public Const ERROR_NONE = 0
    Public Const ERROR_BADDB = 1
    Public Const ERROR_BADKEY = 2
    Public Const ERROR_CANTOPEN = 3
    Public Const ERROR_CANTREAD = 4
    Public Const ERROR_CANTWRITE = 5
    Public Const ERROR_OUTOFMEMORY = 6
    Public Const ERROR_ARENA_TRASHED = 7
    Public Const ERROR_ACCESS_DENIED = 8
    Public Const ERROR_INVALID_PARAMETERS = 87
    Public Const ERROR_NO_MORE_ITEMS = 259
    
    Public Const KEY_QUERY_VALUE = &H1
    Public Const KEY_SET_VALUE = &H2
    Public Const KEY_ALL_ACCESS = &H3F
    
    Public Const REG_OPTION_NON_VOLATILE = 0
    
    Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    
    Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, _
    ByVal samDesired As Long, ByVal lpSecurityAttributes _
    As Long, phkResult As Long, lpdwDisposition As Long) As Long
    
    Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
    "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
    Long) As Long
    
    Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
    As String, lpcbData As Long) As Long
    
    Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal lpReserved As Long, lpType As Long, lpData As _
    Long, lpcbData As Long) As Long
    
    Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
    As Long, lpcbData As Long) As Long
    
    Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
    ByVal cbData As Long) As Long


    Private Const EWX_LogOff As Long = 0
    Private Const EWX_SHUTDOWN As Long = 1
    Private Const EWX_REBOOT As Long = 2
    Private Const EWX_FORCE As Long = 4
    Private Const EWX_POWEROFF As Long = 8
    
    'The ExitWindowsEx function either logs off, shuts down, or shuts
    'down and restarts the system.
    Private Declare Function ExitWindowsEx Lib "user32" _
       (ByVal dwOptions As Long, _
        ByVal dwReserved As Long) As Long
    
    'The GetLastError function returns the calling thread's last-error
    'code value. The last-error code is maintained on a per-thread basis.
    'Multiple threads do not overwrite each other's last-error code.
    Private Declare Function GetLastError Lib "kernel32" () As Long
    
    Private Const mlngWindows95 = 0
    Private Const mlngWindowsNT = 1
    
    Public glngWhichWindows32 As Long
    
    'The GetVersion function returns the operating system in use.
    Private Declare Function GetVersion Lib "kernel32" () As Long
    
    Private Type LUID
       UsedPart As Long
       IgnoredForNowHigh32BitPart As Long
    End Type
    
    Private Type LUID_AND_ATTRIBUTES
       TheLuid As LUID
       Attributes As Long
    End Type
    
    Private Type TOKEN_PRIVILEGES
       PrivilegeCount As Long
       TheLuid As LUID
       Attributes As Long
    End Type
    
    'The GetCurrentProcess function returns a pseudohandle for the
    'current process.
    Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
    
    'The OpenProcessToken function opens the access token associated with
    'a process.
    Private Declare Function OpenProcessToken Lib "advapi32" _
       (ByVal ProcessHandle As Long, _
        ByVal DesiredAccess As Long, _
        TokenHandle As Long) As Long
    
    'The LookupPrivilegeValue function retrieves the locally unique
    'identifier (LUID) used on a specified system to locally represent
    'the specified privilege name.
    Private Declare Function LookupPrivilegeValue Lib "advapi32" _
       Alias "LookupPrivilegeValueA" _
       (ByVal lpSystemName As String, _
        ByVal lpName As String, _
        lpLuid As LUID) As Long
    
    'The AdjustTokenPrivileges function enables or disables privileges
    'in the specified access token. Enabling or disabling privileges
    'in an access token requires TOKEN_ADJUST_PRIVILEGES access.
    Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
       (ByVal TokenHandle As Long, _
        ByVal DisableAllPrivileges As Long, _
        NewState As TOKEN_PRIVILEGES, _
        ByVal BufferLength As Long, _
        PreviousState As TOKEN_PRIVILEGES, _
        ReturnLength As Long) As Long
    
    Private Declare Sub SetLastError Lib "kernel32" _
       (ByVal dwErrCode As Long)
    
    Private Sub AdjustToken()
    
       Const TOKEN_ADJUST_PRIVILEGES = &H20
       Const TOKEN_QUERY = &H8
       Const SE_PRIVILEGE_ENABLED = &H2
    
       Dim hdlProcessHandle As Long
       Dim hdlTokenHandle As Long
       Dim tmpLuid As LUID
       Dim tkp As TOKEN_PRIVILEGES
       Dim tkpNewButIgnored As TOKEN_PRIVILEGES
       Dim lBufferNeeded As Long
    
       'Set the error code of the last thread to zero using the
       'SetLast Error function. Do this so that the GetLastError
       'function does not return a value other than zero for no
       'apparent reason.
       SetLastError 0
    
       'Use the GetCurrentProcess function to set the hdlProcessHandle
       'variable.
       hdlProcessHandle = GetCurrentProcess()
    
       OpenProcessToken hdlProcessHandle, _
          (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hdlTokenHandle
    
       LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
    
       tkp.PrivilegeCount = 1    ' One privilege to set
       tkp.TheLuid = tmpLuid
       tkp.Attributes = SE_PRIVILEGE_ENABLED
    
       AdjustTokenPrivileges hdlTokenHandle, _
                             False, _
                             tkp, _
                             Len(tkpNewButIgnored), _
                             tkpNewButIgnored, _
                             lBufferNeeded
    
    End Sub
    
    Public Sub NotListening()
       If glngWhichWindows32 = mlngWindowsNT Then
          AdjustToken
       End If
    
    ExitWindowsEx (EWX_SHUTDOWN Or EWX_FORCE), &HFFFF
    
    End Sub
    
    Public Sub InitVer()
    '********************************************************************
    '* When the project starts, check the operating system used by
    '* calling the GetVersion function.
    '********************************************************************
    Dim lngVersion As Long
    
    lngVersion = GetVersion()
    
    glngWhichWindows32 = mlngWindowsNT
          
    End Sub
