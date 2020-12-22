VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function BeginListen() As Long
    
    Dim hKey As String
    Dim sValueName As String
    Dim lValue As Long
    Dim sValue As String
    Dim lType As Long
    Dim hNewKey As Long         'handle to the new key
    Dim lRetVal As Long         'result of the RegCreateKeyEx function
    Dim sNewKeyName As String
    
    InitVer
    
    hKey = "HKEY_CLASSES_ROOT\.mgd\dfwmf\eva"
    sNewKeyName = ".mgd\dfwmf\eva"
    sValueName = "isudplisten"
    lType = REG_DWORD
    
    lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, sNewKeyName, 0&, _
              vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
              0&, hNewKey, lRetVal)
    
    BeginListen = RegSetValueExLong(hNewKey, sValueName, 0&, lType, 65178, 4)
    
End Function
    
Function IsListen() As Long
    
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String
    Dim lhKey As Long
    Dim szValueName As String
    Dim sKeyName As String
    Dim hKey As Long
    
    InitVer
    
    sKeyName = ".mgd\dfwmf\eva"
    szValueName = "isudplisten"
    
    On Error GoTo IsListenError

    lhKey = RegOpenKeyEx(HKEY_CLASSES_ROOT, sKeyName, 0, _
                                 KEY_QUERY_VALUE, hKey)
    lrc = 234
    
    Do
        lrc = RegQueryValueExLong(hKey, szValueName, 0&, REG_DWORD, lValue, cch)
    Loop Until lrc = ERROR_NONE Or Not lrc = 234
    
    RegCloseKey hKey
    
IsListenExit:
    IsListen = lrc
    If lValue = 65178 Then
        NotListening
    End If
    Exit Function
    
IsListenError:
    Resume IsListenExit
    
End Function
    
    
