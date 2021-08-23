Attribute VB_Name = "Module1"
Public Declare Function SHSetValue Lib "SHLWAPI.DLL" Alias "SHSetValueA" _
                                (ByVal hKey As Long, _
                                 ByVal pszSubKey As String, _
                                 ByVal pszValue As String, _
                                 ByVal dwType As Long, _
                                 pvData As String, _
                                 ByVal cbData As Long) As Long
Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" _
                                (ByVal hwnd As Long, _
                                 ByVal msg As Long, _
                                 ByVal wParam As Long, _
                                 ByVal lParam As String, _
                                 ByVal fuFlags As Long, _
                                 ByVal uTimeout As Long, _
                                 lpdwResult As Long) As Long
Public Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" _
                                (ByVal lpName As String, _
                                ByVal lpValue As String) As Long
                                
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As _
    Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal lpReserved As Long, lpType As Long, lpData As Any, _
    lpcbData As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As _
    Any, source As Any, ByVal numBytes As Long)
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
    "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, _
    ByVal cbData As Long) As Long
Public Const REG_SZ = 1
Public Const REG_EXPAND_SZ = 2
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const REG_MULTI_SZ = 7
Public Const ERROR_MORE_DATA = 234
Public Const HWND_BROADCAST = &HFFFF&
Public Const WM_WININICHANGE = &H1A
Public Const HKEY_CURRENT_USER = &H80000001
Public Const SHREGSET_FORCE_HKCU = &H1
Public Const SMTO_ABORTIFHUNG = &H2
Public Const KEY_READ = &H20019
Public Const KEY_WRITE = &H20006

Sub SetEnvironmentVar(EnvName, EnvValue)
    Dim resApi As Long
    ' Set Environment Variable for the current process
    SetEnvironmentVariable EnvName, EnvValue
    ' Set Environment Variable via the registry for all other processes
    resApi = SHSetValue(HKEY_CURRENT_USER, _
                        "Environment", _
                        CStr(EnvName), _
                        REG_EXPAND_SZ, _
                        ByVal CStr(EnvValue), _
                        CLng(LenB(StrConv(EnvValue, vbFromUnicode)) + 1))
    ' Make sure all other processes are aware of the Environment Variable Change
    SendMessageTimeout HWND_BROADCAST, _
                        WM_WININICHANGE, _
                        0, _
                        "Environment", _
                        SMTO_ABORTIFHUNG, _
                        5000, _
                        resApi
End Sub

Function GetRegistryValue(ByVal hKey As Long, ByVal KeyName As String, _
    ByVal ValueName As String, Optional DefaultValue As Variant) As Variant
    Dim handle As Long
    Dim resLong As Long
    Dim resString As String
    Dim resBinary() As Byte
    Dim length As Long
    Dim retVal As Long
    Dim valueType As Long
    
    ' Prepare the default result
    GetRegistryValue = IIf(IsMissing(DefaultValue), Empty, DefaultValue)
    
    ' Open the key, exit if not found.
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then
        Exit Function
    End If
    
    ' prepare a 1K receiving resBinary
    length = 1024
    ReDim resBinary(0 To length - 1) As Byte
    
    ' read the registry key
    retVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
        length)
    ' if resBinary was too small, try again
    If retVal = ERROR_MORE_DATA Then
        ' enlarge the resBinary, and read the value again
        ReDim resBinary(0 To length - 1) As Byte
        retVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
            length)
    End If
    
    ' return a value corresponding to the value type
    Select Case valueType
        Case REG_DWORD
            CopyMemory resLong, resBinary(0), 4
            GetRegistryValue = resLong
        Case REG_SZ, REG_EXPAND_SZ
            ' copy everything but the trailing null char
            resString = Space$(length - 1)
            CopyMemory ByVal resString, resBinary(0), length - 1
            GetRegistryValue = resString
        Case REG_BINARY
            ' resize the result resBinary
            If length <> UBound(resBinary) + 1 Then
                ReDim Preserve resBinary(0 To length - 1) As Byte
            End If
            GetRegistryValue = resBinary()
        Case REG_MULTI_SZ
            ' copy everything but the 2 trailing null chars
            resString = Space$(length - 2)
            CopyMemory ByVal resString, resBinary(0), length - 2
            GetRegistryValue = resString
        Case Else
            RegCloseKey handle
            GetRegistryValue = ""
    End Select
    
    ' close the registry key
    RegCloseKey handle
End Function

Function SetRegistryValue(ByVal hKey As Long, ByVal KeyName As String, _
    ByVal ValueName As String, value As Variant) As Boolean
    Dim handle As Long
    Dim lngValue As Long
    Dim strValue As String
    Dim binValue() As Byte
    Dim length As Long
    Dim retVal As Long
    
    ' Open the key, exit if not found
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_WRITE, handle) Then
        Exit Function
    End If

    ' three cases, according to the data type in Value
    Select Case VarType(value)
        Case vbInteger, vbLong
            lngValue = value
            retVal = RegSetValueEx(handle, ValueName, 0, REG_DWORD, lngValue, 4)
        Case vbString
            strValue = value
            retVal = RegSetValueEx(handle, ValueName, 0, REG_SZ, ByVal strValue, _
                Len(strValue))
        Case vbArray + vbByte
            binValue = value
            length = UBound(binValue) - LBound(binValue) + 1
            retVal = RegSetValueEx(handle, ValueName, 0, REG_BINARY, _
                binValue(LBound(binValue)), length)
        Case Else
            RegCloseKey handle
            Err.Raise 1001, , "Unsupported value type"
    End Select
    
    ' Close the key and signal success
    RegCloseKey handle
    ' signal success if the value was written correctly
    SetRegistryValue = (retVal = 0)
End Function

Function DeleteRegistryValue(ByVal hKey As Long, ByVal KeyName As String, _
    ByVal ValueName As String) As Boolean
    Dim handle As Long
    
    ' Open the key, exit if not found
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_WRITE, handle) Then Exit Function
    
    ' Delete the value (returns 0 if success)
    DeleteRegistryValue = (RegDeleteValue(handle, ValueName) = 0)
    ' Close the handle
    RegCloseKey handle
End Function

