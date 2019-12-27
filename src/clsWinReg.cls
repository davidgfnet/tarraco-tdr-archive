VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWinReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eHKEY
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
    HKEY_FIRST = HKEY_CLASSES_ROOT
    HKEY_LAST = HKEY_DYN_DATA
End Enum
Public Enum eHKEYError
    ERROR_SUCCESS = 0
    ERROR_NONE = 0
    ERROR_FILE_NOT_FOUND = 2&
    ERROR_ACCESS_DENIED = 5&
    ERROR_OUTOFMEMORY = 6&
    ERROR_INVALID_PARAMETER = 7&
    ERROR_INVALID_PARAMETERS = 87&
    ERROR_MORE_DATA = 234&
    ERROR_NO_MORE_ITEMS = 259&
    ERROR_BADKEY = 1010&
End Enum

Public Enum eHKEYDataType
    REG_NONE = 0&
    REG_SZ = 1&
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
End Enum

Const SYNCHRONIZE = &H100000
Const READ_CONTROL = &H20000
Const STANDARD_RIGHTS_ALL = &H1F0000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)

Public Enum eREGSAM
    KEY_QUERY_VALUE = &H1
    KEY_SET_VALUE = &H2
    KEY_CREATE_SUB_KEY = &H4
    KEY_ENUMERATE_SUB_KEYS = &H8
    KEY_NOTIFY = &H10
    KEY_CREATE_LINK = &H20
    KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
    KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
    KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
End Enum

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" _
    (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, _
    ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, _
    lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, _
    lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, _
    lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, _
    phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal lpValueName As String, lpcbValueName As Long, _
    ByVal lpReserved As Long, lpType As Long, lpData As Any, _
    lpcbData As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
    (ByVal hKey As Long, ByVal lpszSubKey As String, _
    phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
    (ByVal hKey As Long, ByVal lpszSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
    (ByVal hKey As Long, ByVal szValueName As String) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" _
    (ByVal hKey As Long, ByVal iSubKey As Long, _
    ByVal lpszName As String, ByVal cchName As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, _
    ByVal lpName As String, lpcbName As Long, _
    ByVal lpReserved As Long, ByVal lpClass As String, _
    lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal lpValue As String, lpcbValue As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpszValueName As String, _
    ByVal dwReserved As Long, lpdwType As Long, _
    lpbData As Any, cbData As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal dwType As Long, ByVal lpData As String, _
    ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, ByVal lpszValueName As String, _
    ByVal dwReserved As Long, ByVal fdwType As Long, _
    lpbData As Any, ByVal cbData As Long) As Long
Private Declare Function RegSaveKeyA Lib "advapi32.dll" _
    (ByVal hKey As Long, ByVal lpFile As String, _
    lpSecurityAttributes As Long) As Long

Public Function CreateKey(ByVal sKey As String) As eHKEYError
    Dim lRet As eHKEYError
    Dim hKey2 As Long
    Dim hKey As Long
    hKey = ParseKey(sKey, hKey)
    lRet = RegOpenKeyEx(hKey, sKey, 0&, KEY_WRITE, hKey2)
    If lRet <> ERROR_SUCCESS Then
        lRet = RegCreateKey(hKey, sKey, hKey2)
    End If
    Call RegCloseKey(hKey2)
    CreateKey = lRet
End Function

Public Function ExistKey(ByVal sKey As String) As Boolean
    Dim ret As eHKEYError
    Dim hKey2 As Long
    Dim hKey As eHKEY
    hKey = HKEY_LOCAL_MACHINE
    hKey = ParseKey(sKey, hKey)
    ret = RegOpenKeyEx(hKey, sKey, 0&, KEY_READ, hKey2)
    If ret = ERROR_SUCCESS Then
        ExistKey = True
        Call RegCloseKey(hKey2)
    Else
        ExistKey = False
    End If
End Function

Private Function DeleteKeyNT(hParentKey As Long, szKey As String) As Long
    Dim hKey As Long, lRet As eHKEYError, cSubKeys As Long, cbMaxSubKeyLen As Long
    Dim cbSubKeyLen As Long, dwIndex As Long, ft As FILETIME, szTempSubKey As String
    Dim szSubKey As String
    
    lRet = RegOpenKeyEx(hParentKey, szKey, 0, KEY_ALL_ACCESS, hKey)
    If Not lRet = ERROR_SUCCESS Then
        DeleteKeyNT = lRet
        Exit Function
    End If
    
    lRet = RegQueryInfoKey(hKey, vbNullString, 0&, 0, _
                           cSubKeys, cbMaxSubKeyLen, _
                           0&, 0&, 0&, 0&, 0&, ft)
    If Not lRet = ERROR_SUCCESS Then
        DeleteKeyNT = lRet
        Call RegCloseKey(hKey)
        Exit Function
    End If
    
    If cSubKeys > 0 Then
        dwIndex = cSubKeys - 1                       ' start at the end
        cbMaxSubKeyLen = cbMaxSubKeyLen + 1          ' +1 for the null terminator
        szTempSubKey = String(cbMaxSubKeyLen, "*")   ' buffer to get name back in
        Do
            cbSubKeyLen = cbMaxSubKeyLen * 2
            
            lRet = RegEnumKeyEx(hKey, dwIndex, szTempSubKey, cbSubKeyLen, 0&, vbNullString, 0&, ft)
            If lRet = ERROR_SUCCESS Then
                szSubKey = left(szTempSubKey, cbSubKeyLen)
                Call DeleteKeyNT(hKey, szSubKey)
            End If
            dwIndex = dwIndex - 1                     ' enumerate backwards
        Loop While dwIndex >= 0
    End If
    
    Call RegCloseKey(hKey)
    
    lRet = RegDeleteKey(hParentKey, szKey)
    DeleteKeyNT = lRet
End Function

Public Function GetReg(ByVal sKey As String, Optional ByVal sValue As String = "", Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER, Optional ByVal bAsString As Boolean = False) As Variant
    Dim lRet As Long, hKey2 As Long, rDT As eHKEYDataType
    Dim retDT As eHKEYDataType, lSize As Long, sData As String
    Dim aData() As Byte, lDWord As Long, i As Long, sTmp As String
    hKey = ParseKey(sKey, hKey)
    lRet = RegOpenKeyEx(hKey, sKey, 0&, KEY_READ, hKey2)
    ReDim aData(0)
    lDWord = 0
    sData = ""
    
    If lRet = ERROR_SUCCESS Then
        lRet = RegQueryValueEx(hKey2, sValue, 0&, retDT, 0&, lSize)
        Select Case retDT
        Case REG_DWORD
            lRet = RegQueryValueEx(hKey2, sValue, 0&, rDT, lDWord, lSize)
        Case REG_EXPAND_SZ, REG_SZ, REG_MULTI_SZ
            If lSize Then
                sData = String$(lSize - 1, Chr$(0))
                lRet = RegQueryValueEx(hKey2, sValue, 0&, rDT, ByVal sData, lSize)
            End If
        Case Else
            If lSize Then
                ReDim aData(lSize)
                lRet = RegQueryValueEx(hKey2, sValue, 0&, rDT, aData(0), lSize)
            End If
        End Select
        RegCloseKey hKey2
    End If
    Select Case retDT
    Case REG_DWORD
        If bAsString Then
            GetReg = "0x" & right$("00000000" & Hex$(lDWord), 8) & " (" & lDWord & ")"
        Else
            GetReg = lDWord
        End If
    Case REG_EXPAND_SZ, REG_SZ
        GetReg = sData
    Case REG_MULTI_SZ
        GetReg = RTrimZero(sData, True)
    Case REG_BINARY
        If bAsString Then
            For i = 0 To UBound(aData) - 1
                sTmp = sTmp & right$("00" & Hex$(aData(i)), 2) & " "
            Next
            GetReg = sTmp
        Else
            GetReg = aData
        End If
    End Select
End Function

Private Function ParseKey(sKey As String, _
                          Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER _
                          ) As eHKEY
    Dim i As Long
    Dim sRootKey As String
    sKey = Trim$(sKey)
    If right$(sKey, 1) = "\" Then
        sKey = left$(sKey, Len(sKey) - 1)
    End If
    i = InStr(sKey, "HKEY_")
    If i Then
        i = InStr(sKey, "\")
        If i Then
            sRootKey = left$(sKey, i - 1)
            sKey = mid$(sKey, i + 1)
        Else
            sRootKey = sKey
            sKey = ""
        End If
    ElseIf left$(sKey, 5) = "HKCR\" Then
        sRootKey = "HKEY_CLASSES_ROOT"
        sKey = mid$(sKey, 6)
    ElseIf left$(sKey, 5) = "HKCU\" Then
        sRootKey = "HKEY_CURRENT_USER"
        sKey = mid$(sKey, 6)
    ElseIf left$(sKey, 5) = "HKLM\" Then
        sRootKey = "HKEY_LOCAL_MACHINE"
        sKey = mid$(sKey, 6)
    ElseIf left$(sKey, 4) = "HKU\" Then
        sRootKey = "HKEY_USERS"
        sKey = mid$(sKey, 5)
    ElseIf left$(sKey, 5) = "HKCC\" Then
        sRootKey = "HKEY_CURRENT_CONFIG"
        sKey = mid$(sKey, 6)
    ElseIf left$(sKey, 5) = "HKDD\" Then
        sRootKey = "HKEY_DYN_DATA"
        sKey = mid$(sKey, 6)
    ElseIf left$(sKey, 5) = "HKPD\" Then
        sRootKey = "HKEY_PERFORMANCE_DATA"
        sKey = mid$(sKey, 6)
    Else
        Select Case hKey
        Case HKEY_FIRST To HKEY_LAST
        Case Else
            hKey = HKEY_CLASSES_ROOT
        End Select
    End If
    If Len(sRootKey) Then
        Select Case sRootKey
        Case "HKEY_CLASSES_ROOT"
            hKey = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_USER"
            hKey = HKEY_CURRENT_USER
        Case "HKEY_LOCAL_MACHINE"
            hKey = HKEY_LOCAL_MACHINE
        Case "HKEY_USERS"
            hKey = HKEY_USERS
        Case "HKEY_CURRENT_CONFIG"
            hKey = HKEY_CURRENT_CONFIG
        Case "HKEY_DYN_DATA"
            hKey = HKEY_DYN_DATA
        Case "HKEY_PERFORMANCE_DATA"
            hKey = HKEY_PERFORMANCE_DATA
        Case Else
            hKey = HKEY_CLASSES_ROOT
        End Select
    End If
    
    ParseKey = hKey
End Function

Public Function OpenKeyEx(ByVal hKey As Long, ByVal lpSubKey As String, _
                        ByVal ulOptions As Long, _
                        ByVal samDesired As eREGSAM, phkResult As Long) As Long
Attribute OpenKeyEx.VB_Description = "Abre una clave del registro, en phkResult devuelve el handle de la clave abierta y se usará para los siguientes accesos."
Attribute OpenKeyEx.VB_HelpID = 21090
    OpenKeyEx = RegOpenKeyEx(hKey, lpSubKey, 0&, samDesired, phkResult)
End Function

Public Function OpenKeyQuery(ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As eREGSAM, phkResult As Long) As Long
Attribute OpenKeyQuery.VB_Description = "Llama a la función del API RegOpenKeyEx"
Attribute OpenKeyQuery.VB_HelpID = 21090
    OpenKeyQuery = RegOpenKeyEx(hKey, lpSubKey, 0&, KEY_QUERY_VALUE, phkResult)
End Function

Public Function EnumValue(ByVal hKey As Long, ByVal dwIndex As Long, _
    lpValueName As String, lpcbValueName As Long, _
    lpReserved As Long, lpType As Long, lpData As Byte, _
    lpcbData As Long) As Long
Attribute EnumValue.VB_Description = "Llama a la función del API RegEnumValue"
Attribute EnumValue.VB_HelpID = 21090
    EnumValue = RegEnumValue(hKey, dwIndex, _
                            lpValueName, lpcbValueName, _
                            lpReserved, lpType, lpData, _
                            lpcbData)

End Function

Public Function CloseKey(ByVal hKey As Long) As Long
Attribute CloseKey.VB_Description = "Cierra la clave abierta anteriormente"
Attribute CloseKey.VB_HelpID = 21090
    CloseKey = RegCloseKey(hKey)
End Function

Public Function QueryInfoKey(ByVal hKey As Long, lpcbMaxValueNameLen As Long) As Long
Attribute QueryInfoKey.VB_Description = "Llama a la función del API RegQueryInfoKey"
Attribute QueryInfoKey.VB_HelpID = 21090
    Dim lpftLastWriteTime As FILETIME
    
    QueryInfoKey = RegQueryInfoKey(hKey, 0&, 0&, 0&, 0&, 0&, 0&, 0&, _
                    lpcbMaxValueNameLen, 0&, 0&, lpftLastWriteTime)
End Function

Public Function RegSetValue2(ByVal hKey As Long, ByVal lpSubKey As String, _
                            ByVal dwType As eHKEYDataType, lpData As String, _
                            ByVal cbData As Long) As Long
Attribute RegSetValue2.VB_Description = "Llama a la función del API RegSetValueEx"
Attribute RegSetValue2.VB_HelpID = 21090
    cbData = Len(lpData)
    RegSetValue2 = RegSetValueEx(hKey, lpSubKey, 0&, REG_SZ, ByVal lpData, cbData)
End Function

Public Function DeleteKeyWin95(ByVal hKey As Long, ByVal szKey As String) As Long
Attribute DeleteKeyWin95.VB_Description = "Borra claves del registro, sólo para Windows 95, no funciona en NT ni en Windows 98"
Attribute DeleteKeyWin95.VB_HelpID = 21090
    
    DeleteKeyWin95 = RegDeleteKey(hKey, szKey)
    
End Function

Public Function SetReg(ByVal sKey As String, ByVal sName As String, _
                       Optional ByVal vValue As Variant, _
                       Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER, _
                       Optional ByVal RegDataType As eHKEYDataType = REG_SZ, _
                       Optional ByVal bCreateKey As Boolean = True) As eHKEYError
Attribute SetReg.VB_Description = "Asignar un valor de cualquier tipo"
Attribute SetReg.VB_HelpID = 21090
    Dim lRet As Long
    Dim hKey2 As Long
    Dim cbData As Long
    Dim aData() As Byte
    Dim sData As String
    Dim lData As Long

    hKey = ParseKey(sKey, hKey)
    
    lRet = RegOpenKeyEx(hKey, sKey, 0&, KEY_WRITE, hKey2)
    
    If lRet <> ERROR_SUCCESS Then
        If bCreateKey Then
            lRet = RegCreateKey(hKey, sKey, hKey2)
        End If
    End If
    If lRet <> ERROR_SUCCESS Then
        SetReg = lRet
        Exit Function
    End If
    
    Select Case RegDataType
    Case REG_BINARY
        aData = vValue
        cbData = UBound(aData)
        lRet = RegSetValueEx(hKey2, sName, 0&, RegDataType, aData(0), cbData)
    Case REG_DWORD
        cbData = 4
        lData = CLng(vValue)
        lRet = RegSetValueEx(hKey2, sName, 0&, RegDataType, lData, cbData)
    Case REG_SZ, REG_EXPAND_SZ
        sData = CStr(vValue)
        If Len(sData) = 0 Then
            sData = ""
        End If
        cbData = Len(sData) + 1
        lRet = RegSetValueEx(hKey2, sName, 0&, RegDataType, ByVal sData, cbData)
    Case Else
    End Select
    lRet = RegCloseKey(hKey2)
    
    SetReg = lRet
End Function

Public Function DeleteKey(ByVal sKey As String, _
                          Optional ByVal sValue As String = "", _
                          Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER _
                          ) As eHKEYError
    Dim lRet As eHKEYError
    Dim hKey2 As Long
    
    Select Case hKey
    Case HKEY_FIRST To HKEY_LAST
    Case Else
        hKey = HKEY_CLASSES_ROOT
    End Select
    
    hKey = ParseKey(sKey)
    
    If Len(sValue) = 0 Then
        DeleteKey = DeleteKeyNT(hKey, sKey)
        Exit Function
    End If
    lRet = RegOpenKeyEx(hKey, sKey, 0&, KEY_WRITE, hKey2)
    If lRet = ERROR_SUCCESS Then
        lRet = RegDeleteValue(hKey2, sValue)
        Call RegCloseKey(hKey2)
    End If
    
    DeleteKey = lRet
End Function

Public Function DeleteValue(ByVal sKey As String, _
                            ByVal sValue As String, _
                            Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER _
                            ) As eHKEYError
    Dim lRet As eHKEYError
    Dim hKey2 As Long
    
    Select Case hKey
    Case HKEY_FIRST To HKEY_LAST
    Case Else
        hKey = HKEY_CLASSES_ROOT
    End Select
    
    hKey = ParseKey(sKey)
    lRet = ERROR_NONE
    If Len(sValue) Then
        lRet = RegOpenKeyEx(hKey, sKey, 0&, KEY_WRITE, hKey2)
        If lRet = ERROR_SUCCESS Then
            lRet = RegDeleteValue(hKey2, sValue)
            Call RegCloseKey(hKey2)
        End If
    End If
    DeleteValue = lRet
End Function

Private Function RTrimZero(ByVal sString As String, _
                           Optional ByVal PorElFinal As Boolean = False) As String
    Dim i As Long
    
    If PorElFinal Then
        For i = Len(sString) To 1 Step -1
            If mid$(sString, i, 1) = Chr$(0) Then
                sString = left$(sString, i - 1)
                Exit For
            End If
        Next
        For i = 1 To Len(sString)
            If mid$(sString, i, 1) = Chr$(0) Then
                Mid$(sString, i, 1) = " "
            End If
        Next
    
    Else
        i = InStr(sString, Chr$(0))
        If i Then
            sString = left$(sString, i - 1)
        End If
    End If
    RTrimZero = sString
End Function

Public Function RegSaveKey(ByVal sKey As String, ByVal lpFile As String) As Long
Attribute RegSaveKey.VB_Description = "Guarda en un fichero el contenido de una clave con sus subclaves y datos, el formato no es el habitual de RegEdit"
Attribute RegSaveKey.VB_HelpID = 21090
    Const stmpFic As String = "\tmp.reg"
    Dim hKey As Long
    Dim hKey2 As Long
    Dim ret As eHKEYError
    
    hKey = ParseKey(sKey)
    ret = RegOpenKeyEx(hKey, sKey, 0&, 0&, hKey2)
    On Local Error Resume Next
    If Len(Dir$(stmpFic, vbHidden + vbReadOnly + vbSystem)) Then
        SetAttr stmpFic, vbNormal
        Kill stmpFic
    End If
    ret = RegSaveKeyA(hKey2, stmpFic, 0&)
    If ret = ERROR_SUCCESS Then
        SetAttr stmpFic, vbNormal
        FileCopy stmpFic, lpFile
        Kill stmpFic
    End If
    Err = 0
    RegCloseKey hKey2
End Function

Public Function GetRegString(ByVal sKey As String, Optional ByVal sValue As String = "", Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER) As String
    Dim ret As Long, hKey2 As Long, rDT As eHKEYDataType, sData As String, lSize As Long
    
    hKey = ParseKey(sKey, hKey)
    ret = RegOpenKeyEx(hKey, sKey, 0&, KEY_READ, hKey2)
    
    If ret = ERROR_SUCCESS Then
        ret = RegQueryValueEx(hKey2, sValue, 0&, rDT, 0&, lSize)
        Select Case rDT
        Case REG_SZ, REG_EXPAND_SZ
            If lSize Then
                sData = String$(lSize - 1, Chr$(0))
                ret = RegQueryValueEx(hKey2, sValue, 0&, rDT, ByVal sData, lSize)
            End If
        End Select
        RegCloseKey hKey2
    End If
    GetRegString = sData
End Function

