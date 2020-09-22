Attribute VB_Name = "modReg"
Option Explicit

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long

Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_INVALID_PARAMETER = 7
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259
Public Const ERROR_NONE = 0
Public Const ERROR_OUTOFMEMORY = 6

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003

Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const STANDARD_RIGHTS_READ = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Public Const SYNCHRONIZE = &H100000

Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Public Const REG_BINARY As Long = 3
Public Const REG_EXPAND_SZ As Long = 2
Public Const REG_DWORD As Long = 4
Public Const REG_LINK As Long = 6
Public Const REG_MULTI_SZ As Long = 7
Public Const REG_NONE As Long = 0
Public Const REG_RESOURCE_LIST As Long = 8
Public Const REG_SZ As Long = 1

Public Enum InTypes
    ValNull = REG_NONE
    ValString = REG_SZ
    ValXString = REG_EXPAND_SZ
    ValBinary = REG_BINARY
    ValDWord = REG_DWORD
    ValLink = REG_LINK
    ValMultiString = REG_MULTI_SZ
    ValResList = REG_RESOURCE_LIST
End Enum

Public Function ReadRegistry(ByVal lngKey As Long, ByVal strSubkey As String, ByVal strValueName As String) As String
    ' This function allows you to get values from anywhere in the Registry,
    ' it currently only handles string and double word values.
    ' ----------
    ' Example:
    ' Text1.Text = ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal")
    
    On Error Resume Next
    
    Dim lngResult As Long
    Dim lngKeyValue As Long
    Dim lngDataTypeValue As Long
    Dim lngValueLen As Long
    Dim strValue As String
    Dim dblDVal As Double
    
    lngResult = RegOpenKey(lngKey, strSubkey, lngKeyValue)
    strValue = Space(2048)
    lngValueLen = Len(strValue)
    lngResult = RegQueryValueEx(lngKeyValue, strValueName, 0&, lngDataTypeValue, strValue, lngValueLen)
    
    If (lngResult = 0) And (Err.Number = 0) Then
       If lngDataTypeValue = REG_DWORD Then
          dblDVal = Asc(Mid(strValue, 1, 1)) + &H100& * Asc(Mid(strValue, 2, 1)) + &H10000 * Asc(Mid(strValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid(strValue, 4, 1)))
          strValue = Format(dblDVal, "000")
       End If
       
       strValue = Left(strValue, lngValueLen - 1)
    Else
       strValue = "Not Found"
    End If
    
    lngResult = RegCloseKey(lngKeyValue)
    
    ReadRegistry = strValue
End Function

Public Sub WriteRegistry(ByVal lngKey As Long, ByVal strSubkey As String, ByVal strValueName As String, ByVal ValType As InTypes, ByVal varValue As Variant)
    ' This sub allows you to write values into the entire Registry,
    ' it currently only handles string and double word values.
    ' ----------
    ' Example:
    ' Call WriteRegistry(HKEY_CURRENT_USER, "SOFTWARE\Company Name\Program Name", "NewValueName", ValString, "NewValue")
    ' Call WriteRegistry(HKEY_CURRENT_USER, "SOFTWARE\Company Name\Program Name", "NewValueName", ValDWord, "31")
    
    On Error Resume Next
    
    Dim lngResult As Long
    Dim lngKeyValue As Long
    Dim lngInLen As Long
    Dim lngNewVal As Long
    Dim strNewVal As String
    
    lngResult = RegCreateKey(lngKey, strSubkey, lngKeyValue)
    
    If ValType = ValDWord Then
        lngNewVal = CLng(varValue)
        lngInLen = 4
        lngResult = RegSetValueExLong(lngKeyValue, strValueName, 0&, ValType, lngNewVal, lngInLen)
    Else
        strNewVal = varValue
        lngInLen = Len(strNewVal)
        lngResult = RegSetValueExString(lngKeyValue, strValueName, 0&, ValType, strNewVal, lngInLen)
    End If
    
    lngResult = RegFlushKey(lngKeyValue)
    lngResult = RegCloseKey(lngKeyValue)
End Sub

Public Function ReadRegistryGetSubkey(ByVal lngKey As Long, ByVal strSubkey As String, ByVal lngID As Long) As String
    ' This function enumerates the subkeys under any given key. Call
    ' repeatedly until "Not Found" is returned - store values in array or something.
    ' ----------
    ' Example - this example just adds all the subkeys to a textbox - you will
    ' probably want to save then into an array or something:
    '
    ' Dim strValue As String
    ' Dim i
    ' strValue = ReadRegistryGetSubkey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", i)
    ' Do Until strValue = "Not Found"
    '     Text1.Text = Text1.Text & " " & strValue
    '     i = i + 1
    '     strValue = ReadRegistryGetSubkey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", i)
    ' Loop
    
    On Error Resume Next
    
    Dim lngResult As Long
    Dim lngKeyValue As Long
    Dim lngDataTypeValue As Long
    Dim lngValueLength As Long
    Dim strValue As String
    Dim dblDVal As Double
    
    lngResult = RegOpenKey(lngKey, strSubkey, lngKeyValue)
    strValue = Space(2048)
    lngValueLength = Len(strValue)
    lngResult = RegEnumKey(lngKeyValue, lngID, strValue, lngValueLength)
    
    If (lngResult = 0) And (Err.Number = 0) Then
        If lngDataTypeValue = REG_DWORD Then
            dblDVal = Asc(Mid(strValue, 1, 1)) + &H100& * Asc(Mid(strValue, 2, 1)) + &H10000 * Asc(Mid(strValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid(strValue, 4, 1)))
            strValue = Format(dblDVal, "000")
        End If
        
        strValue = Left(strValue, lngValueLength - 1)
    Else
        strValue = "Not Found"
    End If
    
    lngResult = RegCloseKey(lngKeyValue)
    
    ReadRegistryGetSubkey = strValue
End Function

Public Function ReadRegistryGetAll(ByVal lngKey As Long, ByVal strSubkey As String, ByVal lngID As Long) As Variant
    ' This function allows you to get all the values from anywhere in the
    ' Registry under any given subkey, it currently only returns string and
    ' double word values.
    ' ----------
    ' Example - returns list of names/values to multiline textbox:
    '
    ' Dim varValue As Variant
    ' Dim i
    ' varValue = ReadRegistryGetAll(HKEY_CURRENT_USER, "Software\Microsoft\Notepad", i)
    ' Do Until varValue(2) = "Not Found"
    '    Text1.Text = Text1.Text & varValue(1) & " " & varValue(2) & vbCrLf
    '    i = i + 1
    '    varValue = ReadRegistryGetAll(HKEY_CURRENT_USER, "Software\Microsoft\Notepad", i)
    ' Loop
    
    On Error Resume Next
    
    Dim lngResult As Long
    Dim lngKeyValue As Long
    Dim lngDataTypeValue As Long
    Dim lngValueLength As Long
    Dim lngValueNameLength As Long
    Dim strValueName As String
    Dim strValue As String
    Dim dblDVal As Double
    
    lngResult = RegOpenKey(lngKey, strSubkey, lngKeyValue)
    strValue = Space(2048)
    strValueName = Space(2048)
    lngValueLength = Len(strValue)
    lngValueNameLength = Len(strValueName)
    lngResult = RegEnumValue(lngKeyValue, lngID, strValueName, lngValueNameLength, 0&, lngDataTypeValue, strValue, lngValueLength)
    
    If (lngResult = 0) And (Err.Number = 0) Then
        If lngDataTypeValue = REG_DWORD Then
            dblDVal = Asc(Mid(strValue, 1, 1)) + &H100& * Asc(Mid(strValue, 2, 1)) + &H10000 * Asc(Mid(strValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid(strValue, 4, 1)))
            strValue = Format(dblDVal, "000")
        End If
        
        strValue = Left(strValue, lngValueLength - 1)
        strValueName = Left(strValueName, lngValueNameLength)
    Else
        strValue = "Not Found"
    End If
    
    lngResult = RegCloseKey(lngKeyValue)
    
    ' Return the datatype, value name and value as an array
    ReadRegistryGetAll = Array(lngDataTypeValue, strValueName, strValue)
End Function

Public Sub DeleteSubkey(ByVal lngKey As Long, ByVal strSubkey As String)
    ' This sub deletes a specified subkey (and all its subkeys and values)
    ' from the registry. Be VERY careful using this sub.
    ' ----------
    ' Example:
    ' Call DeleteSubkey(HKEY_CURRENT_USER, "Software\Company Name\Program Name")
    
    On Error Resume Next
    
    Dim lngResult As Long
    Dim lngKeyValue As Long
    
    lngResult = RegOpenKeyEx(lngKey, vbNullChar, 0&, KEY_ALL_ACCESS, lngKeyValue)
    lngResult = RegDeleteKey(lngKeyValue, strSubkey)
    lngResult = RegCloseKey(lngKeyValue)
End Sub

Public Function DeleteValue(ByVal lngKey As Long, ByVal strSubkey As String, ByVal strValue As String) As String
    ' This sub deletes a value from below a specified subkey.
    ' Be VERY careful using this sub.
    ' ----------
    ' Example:
    ' Call DeleteValue(HKEY_CURRENT_USER, "Software\Company Name\Program Name", "ValueToDelete")
    
    On Error Resume Next
    
    Dim lngResult As Long
    Dim lngKeyValue As Long
    
    lngResult = RegOpenKey(lngKey, strSubkey, lngKeyValue)
    lngResult = RegDeleteValue(lngKeyValue, strValue)
    lngResult = RegCloseKey(lngKeyValue)
End Function
