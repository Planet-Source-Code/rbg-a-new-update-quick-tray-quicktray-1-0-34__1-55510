Attribute VB_Name = "modAutoBoot"
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_SZ = 1
Private Const KEY_WRITE = 131078

Public Function DoStartUp(FileName As String, Discription As String)
Dim hKey As Long
 RegOpenKeyEx HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run-", 0, KEY_WRITE, hKey
 RegDeleteValue hKey, Discription
 RegCloseKey hKey
 RegOpenKeyEx HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", 0, KEY_WRITE, hKey
 RegSetValueEx hKey, Discription, 0, REG_SZ, FileName, Len(FileName)
 RegCloseKey hKey
End Function

Public Function DoNotStartUp(FileName As String, Discription As String)
Dim hKey As Long
 RegOpenKeyEx HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", 0, KEY_WRITE, hKey
 RegDeleteValue hKey, Discription
 RegCloseKey hKey
 RegOpenKeyEx HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run-", 0, KEY_WRITE, hKey
 RegSetValueEx hKey, Discription, 0, REG_SZ, FileName, Len(FileName)
 RegCloseKey hKey
End Function

