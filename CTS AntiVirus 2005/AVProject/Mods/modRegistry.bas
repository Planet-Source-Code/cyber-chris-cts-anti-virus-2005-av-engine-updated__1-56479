Attribute VB_Name = "mdlRegistry"
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Private Const REG_SZ               As Long = 1
''Private Const REG_EXPAND_SZ       As Long = 2
Private Const REG_DWORD            As Long = 4
Private Const HKEY_CLASSES_ROOT    As Long = &H80000000
Private Const HKEY_CURRENT_USER    As Long = &H80000001
Private Const KEY_ALL_ACCESS       As Long = &H3F
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, _
                                                                                        ByVal lpValueName As String, _
                                                                                        ByVal Reserved As Long, _
                                                                                        ByVal dwType As Long, _
                                                                                        ByVal lpValue As String, _
                                                                                        ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, _
                                                                                      ByVal lpValueName As String, _
                                                                                      ByVal Reserved As Long, _
                                                                                      ByVal dwType As Long, _
                                                                                      lpValue As Long, _
                                                                                      ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, _
                                                                                    ByVal lpValueName As String) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, _
                                                                               ByVal lpSubKey As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, _
                                                                            ByVal lpSubKey As String, _
                                                                            phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                     ByVal lpValueName As String, _
                                                                                     ByVal lpReserved As Long, _
                                                                                     lpType As Long, _
                                                                                     lpData As Any, _
                                                                                     lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, _
                                                                                 ByVal lpValueName As String, _
                                                                                 ByVal Reserved As Long, _
                                                                                 ByVal dwType As Long, _
                                                                                 lpData As Any, _
                                                                                 ByVal cbData As Long) As Long
Public Sub DeleteValue(lPredefinedKey As Long, _
                       sKeyName As String, _
                       sValueName As String)
'Dim lRetVal As Long
Dim hKey    As Long   'handle of open key
'open the specified key
    RegOpenKeyEx lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey
    RegDeleteValue hKey, sValueName
    RegCloseKey (hKey)
End Sub
Public Sub SetKeyValue(lPredefinedKey As Long, _
                       sKeyName As String, _
                       sValueName As String, _
                       vValueSetting As Variant, _
                       lValueType As Long)
'Dim lRetVal As Long
Dim hKey    As Long   'handle of open key
    RegOpenKeyEx lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey
    SetValueEx hKey, sValueName, lValueType, vValueSetting
    RegCloseKey (hKey)
End Sub
Public Function SetValueEx(ByVal hKey As Long, _
                           sValueName As String, _
                           lType As Long, _
                           vValue As Variant) As Long
Dim lValue As Long
Dim sValue As String
    Select Case lType
    Case REG_SZ
        sValue = vValue
        SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
    Case REG_DWORD
        lValue = vValue
        SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
    End Select
End Function
''
''Public Sub CreateNewKey(ByVal lPredefinedKey As Long, ByVal sNewKeyName As String)
''
''
''
''Dim hNewKey As Long
'''lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
''RegCloseKey (hNewKey)
''End Sub
''
''
'''Public Function DeleteValue(ByVal Hkey As HKeyTypes, ByVal strPath As String, ByVal strValue As String)
'''
'''    'EXAMPLE:
'''    '
'''    'Call DeleteValue(HKEY_CURRENT_USER, "So
'''
'''    ' ftware\VBW\Registry", "Dword")
'''    '
'''    Dim keyhand As Long
'''    r = RegOpenKey(Hkey, strPath, keyhand)
'''    r = RegDeleteValue(keyhand, strValue)
'''    r = RegCloseKey(keyhand)
'''End Function
''Public Sub SaveString(hKey As Long, StrPath As String, strValue As String, strdata As String)
''
''
'''EXAMPLE:
'''
'''Call savestring(HKEY_CURRENT_USER, "Sof
''' tware\VBW\Registry", "String", text1.tex
''' t)
'''
''Dim keyhand As Long
'''Dim r       As Long
''
''
''
''RegCreateKey hKey, StrPath, keyhand
''
''
''RegSetValueEx keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata)
''
''
''RegCloseKey keyhand
''
''End Sub
''


