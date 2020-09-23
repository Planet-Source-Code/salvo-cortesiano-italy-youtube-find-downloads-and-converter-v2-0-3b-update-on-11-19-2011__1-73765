Attribute VB_Name = "moRegistry"
Option Explicit

' .... Make the change into Registry (=Refresh=)
Private Declare Sub SHChangeNotify Lib "Shell32" (ByVal wEventId As Long, ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long)
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0

Public sKeys As Collection

Dim r As Long, keyhand As Long, datatype As Long
Dim strBuf As String, lDataBufSize As Long, intZeroPos As Integer
Dim lValueType As Long, lBuf As Long, ProdKey As String
Dim strString As String
Dim lngTopKey As Long
Dim strSubkey As String
Dim lResult As Long
                                            
Dim hKey As Long, MainKeyHandle As Long
Dim rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
Dim lDataSize As Long
Dim ByteArray() As Byte
Const DisplayErrorMsg = False

Public Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
Public Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Public Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
Public Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
'Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006

Public Const SYNCHRONIZE = &H100000
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_NOTIFY = &H10
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_SET_VALUE = &H2
Public Const KEY_QUERY_VALUE = &H1
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const REG_OPTION_NON_VOLATILE = 0&

Public Const REG_NONE = 0&
Public Const REG_DWORD = 4
Public Const REG_DWORD_LITTLE_ENDIAN = 4&
Public Const REG_DWORD_BIG_ENDIAN = 5&
Public Const REG_LINK = 6&
Public Const REG_MULTI_SZ = 7&
Public Const REG_RESOURCE_LIST = 8&
Public Const REG_FULL_RESOURCE_DESCRIPTOR = 9&
Public Const REG_RESOURCE_REQUIREMENTS_LIST = 10&
Public Const REG_SZ = 1&
Public Const REG_EXPAND_SZ = 2&
Public Const REG_BINARY = 3&

Public Const READ_CONTROL = &H20000
Public Const WRITE_DAC = &H40000
Public Const WRITE_OWNER = &H80000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_READ = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Public Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Public Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Public Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Public Const KEY_EXECUTE = KEY_READ

Public Const ERROR_NONE = 0&
Public Const ERROR_BADDB = 1&
Public Const ERROR_BADKEY = 2&
Public Const ERROR_CANTOPEN = 3&
Public Const ERROR_CANTREAD = 4&
Public Const ERROR_CANTWRITE = 5&
Public Const ERROR_OUTOFMEMORY = 6&
Public Const ERROR_INVALID_PARAMETER = 7&
Public Const ERROR_ACCESS_DENIED = 8&
Public Const ERROR_NO_MORE_ITEMS = 259&
Public Const ERROR_MORE_DATA = 234&
Public Const ERROR_SUCCESS = 0&
Public Function DeleteKey(ByVal hKey As Long, ByVal strKey As String)
On Error GoTo ErrorHeadler
r = RegDeleteKey(hKey, strKey)
Exit Function
ErrorHeadler:
If Err.Number <> 0 Then
    Exit Function
    End If
End Function

Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
On Error GoTo ErrorHeadler
r = RegOpenKey(hKey, strPath, keyhand)
r = RegDeleteValue(keyhand, strValue)
Exit Function
ErrorHeadler:
If Err.Number <> 0 Then
    Exit Function
    End If
End Function

Public Function GetDWORD(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
On Error GoTo ErrorHeadler
r = RegOpenKey(hKey, strPath, keyhand)
lDataBufSize = 4
lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)
If lResult = ERROR_SUCCESS Then
    If lValueType = REG_DWORD Then
        GetDWORD = lBuf
    End If
Else
End If
r = RegCloseKey(keyhand)
Exit Function
ErrorHeadler:
If Err.Number <> 0 Then
    Exit Function
    End If
End Function

Public Function GetString(hKey As Long, strPath As String, strValue As String)
On Error GoTo ErrorHeadler
r = RegOpenKey(hKey, strPath, keyhand)
lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, " ")
    lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        intZeroPos = InStr(strBuf, Chr$(0))
        If intZeroPos > 0 Then
            GetString = Left$(strBuf, intZeroPos - 1)
        Else
            GetString = strBuf
        End If
    End If
End If
Exit Function
ErrorHeadler:
If Err.Number <> 0 Then
    Exit Function
    End If
End Function

Public Sub SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal Ldata As Long)
On Error GoTo ErrorHeadler
    r = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, Ldata, 4)
    r = RegCloseKey(keyhand)
Exit Sub
ErrorHeadler:
    If Err.Number <> 0 Then
    Exit Sub
    End If
End Sub

Public Sub SaveKey(hKey As Long, strPath As String)
Dim keyhand&
On Error GoTo ErrorHeadler
r = RegCreateKey(hKey, strPath, keyhand&)
r = RegCloseKey(keyhand&)
Exit Sub
ErrorHeadler:
    If Err.Number <> 0 Then
    Exit Sub
    End If
End Sub

Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
On Error GoTo ErrorHeadler
r = RegCreateKey(hKey, strPath, keyhand)
r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
r = RegCloseKey(keyhand)
Exit Sub
ErrorHeadler:
    If Err.Number <> 0 Then
    Exit Sub
    End If
End Sub

Function SaveBinary(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal Ldata As Long)
On Error GoTo ErrorHeadler
r = RegCreateKey(hKey, strPath, keyhand)
lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_BINARY, Ldata, 4)
r = RegCloseKey(keyhand)
Exit Function
ErrorHeadler:
    If Err.Number <> 0 Then
    Exit Function
    End If
End Function

Function GetBinary(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String)
On Error GoTo ErrorHeadler
r = RegOpenKey(hKey, strPath, keyhand)
lDataBufSize = 4
lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)
If lResult = ERROR_SUCCESS Then
    If lValueType = REG_BINARY Then
        GetBinary = lBuf
    End If
Else
End If
r = RegCloseKey(keyhand)
Exit Function
ErrorHeadler:
If Err.Number <> 0 Then
    Exit Function
    End If
End Function

Sub ParseKey(KeyName As String, Keyhandle As Long)
rtn = InStr(KeyName, "\")
If Left(KeyName, 5) <> "HKEY_" Or Right(KeyName, 1) = "\" Then
   MsgBox "Incorrect Format:" + Chr(10) + Chr(10) + KeyName
   Exit Sub
ElseIf rtn = 0 Then
   Keyhandle = GetMainKeyHandle(KeyName)
   KeyName = ""
Else
   Keyhandle = GetMainKeyHandle(Left(KeyName, rtn - 1))
   KeyName = Right(KeyName, Len(KeyName) - rtn)
End If
End Sub

Function ErrorMsg(lErrorCode As Long) As String
Dim GetErrorMsg As String
Select Case lErrorCode
       Case 1009, 1015
            GetErrorMsg = "Il DataBase del Registro è corrotto!!"
       Case 2, 1010
            GetErrorMsg = "Chiave di registro non valida!"
       Case 1011
            GetErrorMsg = "Impossibile aprire la chiave di registro!"
       Case 4, 1012
            GetErrorMsg = "Impossibile leggere la Chiave di registro!"
       Case 5
            GetErrorMsg = "L'accesso a questa chiave di registro è stato respinto!"
       Case 1013
            GetErrorMsg = "Non è stato possibile scrivere un valore della Chiave di registro!"
       Case 8, 14
            GetErrorMsg = "Out of memory/Memoria satura!"
       Case 87
            GetErrorMsg = "Parametri non validi!"
       Case 234
            GetErrorMsg = "There is more data than the buffer has been allocated to hold."
       Case Else
            GetErrorMsg = "Errore indefinito. Codice:  " & Str$(lErrorCode)
End Select
End Function

Function GetMainKeyHandle(MainKeyName As String) As Long
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006

Select Case MainKeyName
       Case "HKEY_CLASSES_ROOT"
            GetMainKeyHandle = HKEY_CLASSES_ROOT
       Case "HKEY_CURRENT_USER"
            GetMainKeyHandle = HKEY_CURRENT_USER
       Case "HKEY_LOCAL_MACHINE"
            GetMainKeyHandle = HKEY_LOCAL_MACHINE
       Case "HKEY_USERS"
            GetMainKeyHandle = HKEY_USERS
       Case "HKEY_PERFORMANCE_DATA"
            GetMainKeyHandle = HKEY_PERFORMANCE_DATA
       Case "HKEY_CURRENT_CONFIG"
            GetMainKeyHandle = HKEY_CURRENT_CONFIG
       Case "HKEY_DYN_DATA"
            GetMainKeyHandle = HKEY_DYN_DATA
End Select

End Function

Sub GetKeyNames(ByVal hKey As Long, ByVal strPath As String)
Dim Cnt As Long, StrBuff As String, strKey As String, TKey As Long
    On Error GoTo ErrorHeadler
    RegOpenKey hKey, strPath, TKey
    Do
        StrBuff = String(255, vbNullChar)
        If RegEnumKeyEx(TKey, Cnt, StrBuff, 255, 0, vbNullString, 0, ByVal 0&) <> 0 Then Exit Do
        Cnt = Cnt + 1
        strKey = Left(StrBuff, InStr(StrBuff, vbNullChar) - 1)
        sKeys.Add strKey
    Loop
Exit Sub
ErrorHeadler:
If Err.Number <> 0 Then
    Exit Sub
    End If
End Sub

Function GetRegistryValue(ByVal hKey As Long, ByVal KeyName As String, _
    ByVal ValueName As String, ByVal KeyType As Integer, _
    Optional DefaultValue As Variant = Empty) As Variant

    Dim handle As Long, resLong As Long
    Dim resString As String, Length As Long
    Dim resBinary() As Byte
    
    ' Prepare the default result.
    GetRegistryValue = DefaultValue
    ' Open the key, exit if not found.
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
    
    Select Case KeyType
        Case REG_DWORD
            ' Read the value, use the default if not found.
            If RegQueryValueEx(handle, ValueName, 0, REG_DWORD, _
                resLong, 4) = 0 Then
                GetRegistryValue = resLong
            End If
        Case REG_SZ
            Length = 1024: resString = Space$(Length)
            If RegQueryValueEx(handle, ValueName, 0, REG_SZ, _
                ByVal resString, Length) = 0 Then
                ' If value is found, trim characters in excess.
                GetRegistryValue = Left$(resString, Length - 1)
            End If
        Case REG_BINARY
            Length = 4096
            ReDim resBinary(Length - 1) As Byte
            If RegQueryValueEx(handle, ValueName, 0, REG_BINARY, _
                resBinary(0), Length) = 0 Then
                ReDim Preserve resBinary(Length - 1) As Byte
                GetRegistryValue = resBinary()
            End If
        Case Else
            Err.Raise 1001, , "Unsupported value type"
    End Select
    
    RegCloseKey handle
End Function

Function EnumRegistryKeys(ByVal hKey As Long, ByVal KeyName As String) As String()
    Dim handle As Long, Index As Long, Length As Long
    ReDim Result(0 To 100) As String
    Dim FileTimeBuffer(100) As Byte
    
    ' Open the key, exit if not found.
    If Len(KeyName) Then
        If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
        ' in all case the subsequent functions use hKey
        hKey = handle
    End If
    
    For Index = 0 To 999999
        ' Make room in the array.
        If Index > UBound(Result) Then
            ReDim Preserve Result(Index + 99) As String
        End If
        Length = 260                   ' Max length for a key name.
        Result(Index) = Space$(Length)
        If RegEnumKey(hKey, Index, Result(Index), Length) Then Exit For
        Result(Index) = Left$(Result(Index), InStr(Result(Index), vbNullChar) - 1)
    Next
   
    ' Close the key, if it was actually opened.
    If handle Then RegCloseKey handle
        
    ' Trim unused items in the array.
    ReDim Preserve Result(Index - 1) As String
    EnumRegistryKeys = Result()
End Function

Function CheckRegistryKey(ByVal hKey As Long, ByVal KeyName As String) As Boolean
    Dim handle As Long
    ' Try to open the key.
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) = 0 Then
        ' The key exists.
        CheckRegistryKey = True
        ' Close it before exiting.
        RegCloseKey handle
    End If
End Function

Public Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
    Dim hNewKey As Long
    Dim lRetVal As Long
    On Error Resume Next
    lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
    RegCloseKey (hNewKey)
End Sub

Public Sub SetKeyValue(sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
    Dim lRetVal As Long
    Dim hKey As Long
    On Error Resume Next
    lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    RegCloseKey (hKey)
End Sub

Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String
    On Error Resume Next
    Select Case lType
        Case REG_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
    End Select
End Function

Function SetRegistryValue(ByVal registry_section As Long, ByVal key_name As String, ByVal value_name As String, ByVal value As Variant) As Boolean
    Dim hKey As Long

    SetRegistryValue = True

    If RegOpenKeyEx(registry_section, _
        key_name, 0&, KEY_SET_VALUE, hKey) <> ERROR_SUCCESS _
            Then Exit Function

    If VarType(value) = vbString Then
        If RegSetValueExString(hKey, value_name, 0&, REG_SZ, _
            value, Len(value)) <> ERROR_SUCCESS Then Exit Function
    ElseIf VarType(value) = vbLong Then
        If RegSetValueExLong(hKey, value_name, 0&, REG_DWORD, _
            value, Len(value)) <> ERROR_SUCCESS Then Exit Function
    Else
        Exit Function
    End If

    RegCloseKey hKey

    SetRegistryValue = False
End Function

Public Function RegistryRefresh()
    On Local Error Resume Next
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Function
