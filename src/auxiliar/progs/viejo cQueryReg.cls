VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cQueryReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colShellFolders As Collection
Private colShellFoldersKey As Collection

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Enum eHKEY
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    '
    HKEY_PERFORMANCE_DATA = &H80000004  ' Sólo para NT
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
    '
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
    REG_NONE = 0&               'No value type
    REG_SZ = 1&                 'Unicode null terminated string
    REG_EXPAND_SZ = 2           'Unicode null terminated string
                                '(with environment variable references)
    REG_BINARY = 3              'Free form binary
    REG_DWORD = 4               '32-bit number
    REG_DWORD_LITTLE_ENDIAN = 4 '32-bit number (same as REG_DWORD)
    REG_DWORD_BIG_ENDIAN = 5    '32-bit number
    REG_LINK = 6                'Symbolic Link (unicode)
    REG_MULTI_SZ = 7            'Multiple Unicode strings
    REG_RESOURCE_LIST = 8       'Resource list in the resource map
    REG_FULL_RESOURCE_DESCRIPTOR = 9    'Resource list in the hardware description
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
                                    'Permission to:
    KEY_QUERY_VALUE = &H1           '   query subkey data
    KEY_SET_VALUE = &H2             '   set subkey data
    KEY_CREATE_SUB_KEY = &H4        '   create subkeys
    KEY_ENUMERATE_SUB_KEYS = &H8    '   enumerate subkeys
    KEY_NOTIFY = &H10               '   for change notification
    KEY_CREATE_LINK = &H20          '   create a symbolic link

    KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
                                    
                                    'Permission for read access
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

Private Declare Function RegLoadKeyA Lib "advapi32.dll" _
    (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal lpFile As String) As Long

Private sUser As String
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long


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

Public Function ClassCLSID(ByVal sClass As String) As String
    Dim sClave As String
    Dim sClsid As String
    Const sRootKey As String = "HKEY_LOCAL_MACHINE\Software\Classes\"
    
    sClave = sRootKey & sClass & "\clsid"
    sClsid = GetRegString(sClave)
    
    ClassCLSID = sClsid
End Function

Public Function ClassTypeLib(ByVal sClass As String) As String
    Dim sClave As String
    Dim sClsid As String
    Const sRootKey As String = "HKEY_LOCAL_MACHINE\Software\Classes\"
    Dim sTypeLib As String
    
    sClave = sRootKey & sClass & "\clsid"
    sClsid = GetRegString(sClave)
    
    If Len(sClsid) Then
        sClave = sRootKey & "CLSID\" & sClsid & "\TypeLib"
        sTypeLib = GetRegString(sClave)
    End If
    ClassTypeLib = sTypeLib
End Function

Private Function DeleteKeyNT(hParentKey As Long, szKey As String) As Long
    Dim hKey As Long
    Dim lRet As eHKEYError
    Dim cSubKeys As Long
    Dim cbMaxSubKeyLen As Long
    Dim cbSubKeyLen As Long
    Dim dwIndex As Long
    Dim ft As FILETIME
    
    Dim szTempSubKey As String
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

Public Function GetRegDWord(ByVal sKey As String, Optional ByVal sValue As String = "", Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER, Optional ByVal bAsString As Boolean = False) As Variant
Attribute GetRegDWord.VB_Description = "Obtener un valor DWORD de una entrada del registro"
Attribute GetRegDWord.VB_HelpID = 21090
    Dim ret As Long
    Dim hKey2 As Long
    Dim rDT As eHKEYDataType
    Dim lSize As Long
    Dim lDWord As Long
    
    hKey = ParseKey(sKey, hKey)
    
    ret = RegOpenKeyEx(hKey, sKey, 0&, KEY_READ, hKey2)
    
    If ret = ERROR_SUCCESS Then
        ret = RegQueryValueEx(hKey2, sValue, 0&, rDT, 0&, lSize)
        If rDT = REG_DWORD Then
            ret = RegQueryValueEx(hKey2, sValue, 0&, rDT, lDWord, lSize)
        End If
        RegCloseKey hKey2
    End If
    If bAsString Then
        GetRegDWord = "0x" & right$("00000000" & Hex$(lDWord), 8) & " (" & lDWord & ")"
    Else
        GetRegDWord = lDWord
    End If
End Function

Public Function GetReg(ByVal sKey As String, Optional ByVal sValue As String = "", Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER, Optional ByVal bAsString As Boolean = False) As Variant
    Dim lRet As Long
    Dim hKey2 As Long
    Dim rDT As eHKEYDataType
    Dim retDT As eHKEYDataType
    Dim lSize As Long
    Dim sData As String
    Dim aData() As Byte
    Dim lDWord As Long
    Dim i As Long
    Dim sTmp As String
    
    hKey = ParseKey(sKey, hKey)
    
    
    lRet = RegOpenKeyEx(hKey, sKey, 0&, KEY_READ, hKey2)
    
        ' Valores por defecto
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
                ' Leer la cadena
                '(el ByVal es porque está declarada como Any)---v
                lRet = RegQueryValueEx(hKey2, sValue, 0&, rDT, ByVal sData, lSize)
            End If
        Case Else ' Tratarlos como REG_BINARY
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

Public Function GetRegType(ByVal sKey As String, _
                           ByVal sValue As String, _
                           Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER) As eHKEYDataType
    Dim lRet As Long
    Dim hKey2 As Long
    Dim retDT As eHKEYDataType
    Dim lSize As Long
    hKey = ParseKey(sKey, hKey)
    retDT = REG_NONE
    lRet = RegOpenKeyEx(hKey, sKey, 0&, KEY_READ, hKey2)
    If lRet = ERROR_SUCCESS Then
        lRet = RegQueryValueEx(hKey2, sValue, 0&, retDT, 0&, lSize)
        RegCloseKey hKey2
    End If
    GetRegType = retDT
End Function

Public Function GetRegBinary(ByVal sKey As String, Optional ByVal sValue As String = "", Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER, Optional ByVal bAsString As Boolean = False) As Variant
Attribute GetRegBinary.VB_Description = "Obtener un valor BINARY de una entrada del registro"
Attribute GetRegBinary.VB_HelpID = 21090
    Dim ret As Long
    Dim hKey2 As Long
    Dim rDT As eHKEYDataType
    Dim lSize As Long
    Dim aData() As Byte
    Dim i As Long
    Dim sTmp As String
    
    hKey = ParseKey(sKey, hKey)
    
    ReDim aData(0)
    
    ret = RegOpenKeyEx(hKey, sKey, 0&, KEY_READ, hKey2)
    
    If ret = ERROR_SUCCESS Then
        ret = RegQueryValueEx(hKey2, sValue, 0&, rDT, 0&, lSize)
        If rDT = REG_BINARY Then
            If lSize Then
                ReDim aData(lSize)
                ret = RegQueryValueEx(hKey2, sValue, 0&, rDT, aData(0), lSize)
            End If
        End If
        RegCloseKey hKey2
    End If
    If bAsString Then
        For i = 0 To UBound(aData) - 1
            sTmp = sTmp & Hex$(aData(i)) & " "
        Next
        GetRegBinary = sTmp
    Else
        GetRegBinary = aData
    End If
End Function

Public Function GetRegString(ByVal sKey As String, Optional ByVal sValue As String = "", Optional ByVal hKey As eHKEY = HKEY_CURRENT_USER) As String
Attribute GetRegString.VB_Description = "Obtener un valor STRING de una entrada del registro"
Attribute GetRegString.VB_HelpID = 21090
    Dim ret As Long
    Dim hKey2 As Long
    Dim rDT As eHKEYDataType
    Dim sData As String
    Dim lSize As Long
    
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

Public Function QueryRegBase(ByVal sValue As String, _
                            Optional ByVal hKey As eHKEY = HKEY_CLASSES_ROOT _
                            ) As String
    Dim sBuf As String
    Dim buflen As Long
    
    Select Case hKey
    Case HKEY_FIRST To HKEY_LAST
    Case Else
        hKey = HKEY_CLASSES_ROOT
    End Select
    
    sBuf = String$(300, Chr$(0))
    buflen = Len(sBuf)
    If RegQueryValue(hKey, sValue, sBuf, buflen) = ERROR_SUCCESS Then
        If buflen > 1 Then
            QueryRegBase = left$(sBuf, buflen - 1)
        Else
            QueryRegBase = ""
        End If
    Else
        QueryRegBase = ""
    End If
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
            sKey = Mid$(sKey, i + 1)
        Else
            sRootKey = sKey
            sKey = ""
        End If
    ElseIf left$(sKey, 5) = "HKCR\" Then
        sRootKey = "HKEY_CLASSES_ROOT"
        sKey = Mid$(sKey, 6)
    ElseIf left$(sKey, 5) = "HKCU\" Then
        sRootKey = "HKEY_CURRENT_USER"
        sKey = Mid$(sKey, 6)
    ElseIf left$(sKey, 5) = "HKLM\" Then
        sRootKey = "HKEY_LOCAL_MACHINE"
        sKey = Mid$(sKey, 6)
    ElseIf left$(sKey, 4) = "HKU\" Then
        sRootKey = "HKEY_USERS"
        sKey = Mid$(sKey, 5)
    ElseIf left$(sKey, 5) = "HKCC\" Then
        sRootKey = "HKEY_CURRENT_CONFIG"
        sKey = Mid$(sKey, 6)
    ElseIf left$(sKey, 5) = "HKDD\" Then
        sRootKey = "HKEY_DYN_DATA"
        sKey = Mid$(sKey, 6)
    ElseIf left$(sKey, 5) = "HKPD\" Then
        sRootKey = "HKEY_PERFORMANCE_DATA"
        sKey = Mid$(sKey, 6)
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

Public Function EnumValueString(ByVal hKey As Long, ByVal dwIndex As Long, _
    lpValueName As String, lpcbValueName As Long, _
    lpReserved As Long, lpType As Long, lpData As String, _
    lpcbData As Long) As Long
Attribute EnumValueString.VB_Description = "Llama a la función del API RegEnumValue"
Attribute EnumValueString.VB_HelpID = 21090
    EnumValueString = RegEnumValue(hKey, dwIndex, _
                            lpValueName, lpcbValueName, _
                            lpReserved, lpType, ByVal lpData, _
                            lpcbData)

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

Public Function EnumKeyEx(ByVal hKey As Long, ByVal dwIndex As Long, lpName As String, lpcbName As Long) As Long
Attribute EnumKeyEx.VB_Description = "Llama a la función del API RegEnumKeyEx"
Attribute EnumKeyEx.VB_HelpID = 21090
    Dim lpftLastWriteTime As FILETIME
    
    EnumKeyEx = RegEnumKeyEx(hKey, dwIndex, lpName, lpcbName, _
                             0&, 0&, 0&, lpftLastWriteTime)

End Function

Public Function ShellFolders(Optional bSoloClaves As Boolean = False, Optional Usuario As Boolean = True) As Variant
Attribute ShellFolders.VB_Description = "Devolverá las claves de la clave Shell Folders"
Attribute ShellFolders.VB_HelpID = 21090
    Dim sKey As String
    Dim buf As String
    Dim i As Long
    Dim sValue As String
    Dim iCount As Long
    Dim colKeys() As String
    Dim colShellFoldersKey As Collection
    Set colShellFolders = Nothing
    Set colShellFolders = New Collection
    Set colShellFoldersKey = New Collection
    
    For iCount = 0 To 1
        If iCount = 0 Then
            If Usuario And sUser <> "" Then
                sKey = "HKEY_USERS\" & sUser & "\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
            Else
                sKey = "HKEY_USERS\.Default\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
            End If
        Else
            sKey = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion"
        End If
        
        If EnumValues(colKeys(), sKey) Then
            For i = 1 To UBound(colKeys) Step 2
                If iCount = 0 Then
                    colShellFoldersKey.Add colKeys(i), colKeys(i)
                    colShellFolders.Add colKeys(i + 1), colKeys(i)
                Else
                    If InStr(colKeys(i + 1), ":\") Then
                        colShellFoldersKey.Add colKeys(i), colKeys(i)
                        colShellFolders.Add colKeys(i + 1), colKeys(i)
                    End If
                End If
            Next
        End If
    Next
    buf = String$(300, Chr$(0))
    i = GetWindowsDirectory(buf, Len(buf))
    sValue = left$(buf, i)
    buf = "WindowsDir"
    colShellFoldersKey.Add buf, buf
    colShellFolders.Add sValue, buf
    
    buf = String$(300, Chr$(0))
    i = GetSystemDirectory(buf, Len(buf))
    sValue = left$(buf, i)
    buf = "SystemDir"
    colShellFoldersKey.Add buf, buf
    colShellFolders.Add sValue, buf
    
    If bSoloClaves Then
        Set ShellFolders = colShellFoldersKey
    Else
        Set ShellFolders = colShellFolders
    End If
        
    Set colShellFoldersKey = Nothing
End Function

Private Sub Class_Initialize()
    Set colShellFolders = New Collection
    Dim n As Long
    n = 255
    sUser = Space(n)
    If GetUserName(sUser, n) <> 0 Then sUser = left$(sUser, n - 1)
End Sub

Private Sub Class_Terminate()
    Set colShellFolders = Nothing
End Sub

Public Function GetFolder(ByVal vIndex As Variant) As String
Attribute GetFolder.VB_Description = "Devuelve el directorio de la clave indicada"
Attribute GetFolder.VB_HelpID = 21090
    If colShellFolders.Count = 0 Then
        Call ShellFolders
    End If
    
    On Local Error Resume Next
    GetFolder = colShellFolders(vIndex)
    
    If Err Then
        GetFolder = ""
    End If
    
    Err = 0
    On Local Error GoTo 0
End Function

Public Function RegSetValue2(ByVal hKey As Long, ByVal lpSubKey As String, _
                            ByVal dwType As eHKEYDataType, lpData As String, _
                            ByVal cbData As Long) As Long
Attribute RegSetValue2.VB_Description = "Llama a la función del API RegSetValueEx"
Attribute RegSetValue2.VB_HelpID = 21090
    cbData = Len(lpData)
    RegSetValue2 = RegSetValueEx(hKey, lpSubKey, 0&, REG_SZ, ByVal lpData, cbData)
End Function

Public Sub AsociarExt(ByVal sExt As String, _
                      Optional ByVal sExe As String = "", _
                      Optional ByVal sCommand As String = "open", _
                      Optional ByVal bDefault As Boolean = True, _
                      Optional ByVal sProgId As String = "", _
                      Optional ByVal sDescription As String = "")
    Dim sDef As String
    Dim hKey As Long
    Dim phkResult As Long
    Dim lRet As eHKEYError
    Dim sValue As String
    Dim sKey As String
    Dim sAccess As String
    
    sExt = Trim$(sExt)
    sExe = Trim$(sExe)
    sCommand = Trim$(sCommand)
    sProgId = Trim$(sProgId)
    sDescription = Trim$(sDescription)
    
    If InStr(sExt, ".") = 0 Then
        sExt = "." & sExt
    End If
    If Len(sExe) Then
        If InStr(sExe, ".") = 0 Then
            sExe = sExe & ".exe"
        End If
        sExe = sExe & " "
    End If
    
    If Len(sProgId) = 0 Then
        sProgId = "progID" & sExt
    End If
    
    If Len(sDescription) = 0 Then
        sDescription = "Descripción de " & sProgId
    End If
    
    sAccess = sCommand
    lRet = InStr(sAccess, "&")
    If lRet Then
        sCommand = left$(sAccess, lRet - 1) & Mid$(sAccess, lRet + 1)
    End If
    
    sValue = sProgId
    sProgId = QueryRegBase(sExt)
    If Len(sProgId) = 0 Then
        sKey = sExt
        sProgId = sValue
        lRet = RegSetValue(HKEY_CLASSES_ROOT, sKey, REG_SZ, sValue, Len(sValue))
        '
        sKey = sProgId
        sValue = sDescription
        lRet = RegSetValue(HKEY_CLASSES_ROOT, sKey, REG_SZ, sValue, Len(sValue))
    End If
    sProgId = QueryRegBase(sExt)
    If Len(sProgId) Then
        sDef = "Software\Classes\" & sProgId & "\shell"
        
        hKey = HKEY_LOCAL_MACHINE
        lRet = RegCreateKey(hKey, sDef, phkResult)
        If lRet = ERROR_SUCCESS Then
            If sCommand <> "open" Then
                sKey = sCommand
                sValue = sAccess
                lRet = RegSetValue(phkResult, sKey, REG_SZ, sValue, Len(sValue))
                If Len(sExe) Then
                    sKey = sCommand & "\command"
                    sValue = sExe & Chr$(34) & "%1" & Chr$(34)
                    lRet = RegSetValue(phkResult, sKey, REG_SZ, sValue, Len(sValue))
                End If
            Else
                If Len(sExe) Then
                    sKey = "open\command"
                    sValue = sExe & Chr$(34) & "%1" & Chr$(34)
                    lRet = RegSetValue(phkResult, sKey, REG_SZ, sValue, Len(sValue))
                End If
            End If
            If bDefault Then
                sKey = ""
                sValue = sCommand 'sProgId
                lRet = RegSetValue(phkResult, sKey, REG_SZ, sValue, Len(sValue))
            End If
            lRet = RegCloseKey(phkResult)
        End If
    End If
End Sub

Public Sub AsociarIcono(ByVal sExt As String, _
                        Optional ByVal sExe As String = "", _
                        Optional ByVal Icono As Long = 0&)
    Dim sDef As String
    Dim hKey As Long
    Dim phkResult As Long
    Dim lRet As eHKEYError
    Dim sValue As String
    Dim sKey As String
    Dim sProgId As String
    sExt = Trim$(sExt)
    sExe = Trim$(sExe)
    If InStr(sExt, ".") = 0 Then
        sExt = "." & sExt
    End If
    If Len(sExe) Then
        If InStr(sExe, ".") = 0 Then
            sExe = sExe & ".exe"
        End If
    End If
    sProgId = QueryRegBase(sExt)
    If Len(sProgId) Then
        sDef = "Software\Classes\" & sProgId & "\DefaultIcon"
        hKey = HKEY_LOCAL_MACHINE
        If Len(sExe) = 0 Then
            Call DeleteKey(sDef, "", HKEY_LOCAL_MACHINE)
        Else
            lRet = RegCreateKey(hKey, sDef, phkResult)
            If lRet = ERROR_SUCCESS Then
                sKey = ""
                sValue = sExe & "," & CStr(Icono)
                lRet = RegSetValue(phkResult, sKey, REG_SZ, sValue, Len(sValue))
                lRet = RegCloseKey(phkResult)
            End If
        End If
    End If
End Sub

Public Function DeleteKeyWin95(ByVal hKey As Long, ByVal szKey As String) As Long
Attribute DeleteKeyWin95.VB_Description = "Borra claves del registro, sólo para Windows 95, no funciona en NT ni en Windows 98"
Attribute DeleteKeyWin95.VB_HelpID = 21090
    
    DeleteKeyWin95 = RegDeleteKey(hKey, szKey)
    
End Function

Public Sub DesasociarExt(ByVal sExt As String)
Attribute DesasociarExt.VB_Description = "Desasociar una extensión, borrar las entradas que haya en el registro"
Attribute DesasociarExt.VB_HelpID = 21090
    Dim sProgId As String
    Dim lRet As eHKEYError
    
    If InStr(sExt, ".") = 0 Then
        sExt = "." & sExt
    End If
    
    sProgId = QueryRegBase(sExt)
    If Len(sProgId) Then
        Call DeleteKeyNT(HKEY_CLASSES_ROOT, sExt)
        Call DeleteKeyNT(HKEY_CLASSES_ROOT, sProgId)
    End If
End Sub

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
'
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
    '
    DeleteValue = lRet
End Function

Public Function EnumKeys(ByRef colKeys() As String, ByVal sKey As String) As Boolean
Attribute EnumKeys.VB_Description = "Enumera todas las subclaves de la clave indicada"
Attribute EnumKeys.VB_HelpID = 21090
    Dim dwIndex             As Long
    Dim ret                 As Long
    Dim hKey2               As Long
    Dim hKey                As Long
    Dim lpName              As String
    Dim lpftLastWriteTime   As FILETIME
    Dim colItems            As Long
    Dim lSize               As Long
    Dim SubKeysNum          As Long
    Dim MaxSubKeyLen        As Long
    Dim numValues           As Long
    Dim MaxValueNameLen     As Long
    Dim MaxDataLen          As Long
    
    colItems = 0
    ReDim colKeys(0)
    
    hKey = ParseKey(sKey, hKey)
    
    ret = RegOpenKeyEx(hKey, sKey, 0&, KEY_READ, hKey2)
    
    EnumKeys = True
    If ret = ERROR_SUCCESS Then
        ret = RegQueryInfoKey(hKey2, vbNullString, 0&, 0&, _
                            SubKeysNum, MaxSubKeyLen, _
                            0&, 0&, 0&, _
                            0&, 0&, lpftLastWriteTime)
        
        For dwIndex = 0 To SubKeysNum
            lSize = MaxSubKeyLen * 2
            lpName = String$(lSize + 1, Chr$(0))
            ret = RegEnumKeyEx(hKey2, dwIndex, lpName, lSize, _
                            0&, vbNullString, 0&, lpftLastWriteTime)
            If ret = ERROR_MORE_DATA Or ret = ERROR_SUCCESS Then
                colItems = colItems + 1
                ReDim Preserve colKeys(colItems)
                colKeys(colItems) = left$(lpName, lSize)
            End If
        Next
    Else
        EnumKeys = False
    End If
    Call CloseKey(hKey2)
End Function

Private Function RTrimZero(ByVal sString As String, _
                           Optional ByVal PorElFinal As Boolean = False) As String
    Dim i As Long
    
    If PorElFinal Then
        For i = Len(sString) To 1 Step -1
            If Mid$(sString, i, 1) = Chr$(0) Then
                sString = left$(sString, i - 1)
                Exit For
            End If
        Next
        For i = 1 To Len(sString)
            If Mid$(sString, i, 1) = Chr$(0) Then
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

Public Function EnumValues(ByRef colKeys() As String, ByVal sKey As String) As Boolean
    Dim dwIndex             As Long
    Dim ret                 As Long
    Dim hKey2               As Long
    Dim hKey                As Long
    Dim lpName              As String
    Dim lpftLastWriteTime   As FILETIME
    Dim retDT               As eHKEYDataType
    Dim lSize               As Long
    Dim sData               As String
    Dim aData()             As Byte
    Dim lDWord              As Long
    Dim i                   As Long
    Dim colItems            As Long
    Dim SubKeysNum          As Long
    Dim MaxSubKeyLen        As Long
    Dim numValues           As Long
    Dim MaxValueNameLen     As Long
    Dim MaxDataLen          As Long
    
    hKey = ParseKey(sKey, hKey)
    
    ret = RegOpenKeyEx(hKey, sKey, 0&, KEY_READ, hKey2)
    
    EnumValues = True
    ReDim aData(0)
    lDWord = 0
    sData = ""
    
    colItems = 2
    ReDim colKeys(colItems)
    colKeys(1) = "(Default)"
    colKeys(2) = QueryRegBase(sKey, hKey)
    '
    If ret = ERROR_SUCCESS Then
        ret = RegQueryInfoKey(hKey2, vbNullString, 0&, 0&, _
                            SubKeysNum, MaxSubKeyLen, _
                            0&, numValues, MaxValueNameLen, _
                            MaxDataLen, 0&, lpftLastWriteTime)
        
        If ret = ERROR_INVALID_PARAMETERS Then
            Debug.Print "ERROR_INVALID_PARAMETERS"
            EnumValues = False
            GoTo SalirEnumValues
        End If
        
        For dwIndex = 0 To numValues
            lpName = String$(MaxValueNameLen + 1, Chr$(0))
            ret = RegEnumValue(hKey2, dwIndex, 0&, 0&, 0&, retDT, 0&, 0&)
            If ret = ERROR_MORE_DATA Or ret = ERROR_SUCCESS Then
                lSize = MaxDataLen
                Select Case retDT
                Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
                    sData = String$(lSize, Chr$(0))
                    ret = RegEnumValue(hKey2, dwIndex, lpName, Len(lpName), 0&, retDT, ByVal sData, lSize)
                    If retDT = REG_MULTI_SZ Then
                        sData = RTrimZero(sData, True)
                    Else
                        sData = RTrimZero(sData)
                    End If
                    lpName = RTrimZero(lpName)
                    ReDim Preserve colKeys(colItems + 2)
                    colKeys(colItems + 1) = lpName
                    colKeys(colItems + 2) = sData
                    colItems = colItems + 2
                Case REG_DWORD
                    ' Datos numéricos (long)
                    ret = RegEnumValue(hKey2, dwIndex, lpName, Len(lpName), 0&, retDT, lDWord, lSize)
                    sData = CStr(lDWord)
                    lpName = RTrimZero(lpName)
                    ReDim Preserve colKeys(colItems + 2)
                    colKeys(colItems + 1) = lpName
                    colKeys(colItems + 2) = sData
                    colItems = colItems + 2
                Case Else
                    If lSize Then
                        ReDim aData(lSize)
                        ret = RegEnumValue(hKey2, dwIndex, lpName, Len(lpName), 0&, retDT, aData(0), lSize)
                        lpName = RTrimZero(lpName)
                        sData = ""
                        For i = 0 To lSize - 1 'UBound(aData) - 1
                            sData = sData & right$("00" & Hex$(aData(i)), 2) & " "
                        Next
                        ReDim Preserve colKeys(colItems + 2)
                        colKeys(colItems + 1) = lpName
                        colKeys(colItems + 2) = sData
                        colItems = colItems + 2
                    End If
                End Select
            End If
        Next
    Else
        EnumValues = False
    End If
SalirEnumValues:
    ret = CloseKey(hKey2)
End Function
'
Public Function EnumValuesByType(ByRef colKeys() As String, ByVal sKey As String, _
                                 Optional ByVal tDT As eHKEYDataType = REG_SZ) As Boolean
    Dim dwIndex             As Long
    Dim ret                 As Long
    Dim hKey2               As Long
    Dim hKey                As Long
    Dim lpName              As String
    Dim lpftLastWriteTime   As FILETIME
    Dim retDT               As eHKEYDataType
    Dim lSize               As Long
    Dim sData               As String
    Dim aData()             As Byte
    Dim lDWord              As Long
    Dim i                   As Long
    Dim colItems            As Long
    Dim SubKeysNum          As Long
    Dim MaxSubKeyLen        As Long
    Dim numValues           As Long
    Dim MaxValueNameLen     As Long
    Dim MaxDataLen          As Long
    
    hKey = ParseKey(sKey, hKey)
    
    ret = RegOpenKeyEx(hKey, sKey, 0&, KEY_READ, hKey2)
    
    EnumValuesByType = True
    ReDim aData(0)
    lDWord = 0
    sData = ""
    
    colItems = 2
    ReDim colKeys(colItems)
    colKeys(1) = "(Default)"
    colKeys(2) = QueryRegBase(sKey, hKey)
    If ret = ERROR_SUCCESS Then
        ret = RegQueryInfoKey(hKey2, vbNullString, 0&, 0&, _
                            SubKeysNum, MaxSubKeyLen, _
                            0&, numValues, MaxValueNameLen, _
                            MaxDataLen, 0&, lpftLastWriteTime)
        
        If ret = ERROR_INVALID_PARAMETERS Then
            Debug.Print "ERROR_INVALID_PARAMETERS"
            EnumValuesByType = False
            GoTo SalirEnumValues
        End If
        
        For dwIndex = 0 To numValues
            lpName = String$(MaxValueNameLen + 1, Chr$(0))
            ret = RegEnumValue(hKey2, dwIndex, 0&, 0&, 0&, retDT, 0&, 0&)
            If ret = ERROR_MORE_DATA Or ret = ERROR_SUCCESS Then
                lSize = MaxDataLen
                If (tDT = retDT) Or (tDT = REG_SZ And (retDT = REG_EXPAND_SZ Or retDT = REG_MULTI_SZ)) Then
                    Select Case retDT
                    Case REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ
                        ' Datos de cadena
                        sData = String$(lSize, Chr$(0))
                        ret = RegEnumValue(hKey2, dwIndex, lpName, Len(lpName), 0&, retDT, ByVal sData, lSize)
                        If retDT = REG_MULTI_SZ Then
                            sData = RTrimZero(sData, True)
                        Else
                            sData = RTrimZero(sData)
                        End If
                        lpName = RTrimZero(lpName)
                        ReDim Preserve colKeys(colItems + 2)
                        colKeys(colItems + 1) = lpName
                        colKeys(colItems + 2) = sData
                        colItems = colItems + 2
                    Case REG_DWORD
                        ' Datos numéricos (long)
                        ret = RegEnumValue(hKey2, dwIndex, lpName, Len(lpName), 0&, retDT, lDWord, lSize)
                        sData = CStr(lDWord)
                        lpName = RTrimZero(lpName)
                        ReDim Preserve colKeys(colItems + 2)
                        colKeys(colItems + 1) = lpName
                        colKeys(colItems + 2) = sData
                        colItems = colItems + 2
                    'Case REG_BINARY
                    '    'Datos binarios
                    Case Else
                        ' Tratarlo como Binary
                        If lSize Then
                            ReDim aData(lSize)
                            ret = RegEnumValue(hKey2, dwIndex, lpName, Len(lpName), 0&, retDT, aData(0), lSize)
                            lpName = RTrimZero(lpName)
                            sData = ""
                            For i = 0 To lSize - 1 'UBound(aData) - 1
                                sData = sData & right$("00" & Hex$(aData(i)), 2) & " "
                            Next
                            ReDim Preserve colKeys(colItems + 2)
                            colKeys(colItems + 1) = lpName
                            colKeys(colItems + 2) = sData
                            colItems = colItems + 2
                        End If
                    End Select
                End If
            End If
        Next
    Else
        EnumValuesByType = False
    End If
SalirEnumValues:
    ret = CloseKey(hKey2)
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



Public Function UnRegister(ByVal sClass As String) As eHKEYError
    Dim sClave As String
    Dim sClsid As String
    Const sRootKey As String = "HKEY_LOCAL_MACHINE\Software\Classes\"
    Dim tKeyError As eHKEYError
    Dim sTypeLib As String
    
    UnRegister = ERROR_NONE
    
    sClave = sRootKey & sClass & "\clsid"
    sClsid = GetRegString(sClave)
    
    If Len(sClsid) Then
        sClave = sRootKey & sClass
        tKeyError = DeleteKey(sClave)
        
        If tKeyError = ERROR_NONE Then
            sClave = sRootKey & "CLSID\" & sClsid & "\TypeLib"
            sTypeLib = GetRegString(sClave)
            
            
            If Len(sTypeLib) Then
                sClave = sRootKey & "CLSID\" & sClsid
                tKeyError = DeleteKey(sClave)
                
                If tKeyError = ERROR_NONE Then
                    sClave = sRootKey & "TypeLib\" & sTypeLib
                    tKeyError = DeleteKey(sClave)
                    
                End If
                UnRegister = tKeyError
            Else
                UnRegister = ERROR_FILE_NOT_FOUND
            End If
        End If
    Else
        UnRegister = ERROR_FILE_NOT_FOUND
    End If
End Function

Public Property Get UserName() As String
    UserName = sUser
End Property

