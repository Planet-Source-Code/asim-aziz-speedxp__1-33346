Attribute VB_Name = "modMain"
Option Explicit
'Following Are the functions to access registry
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

'Initialize common controls to use the new windows XP themes
Public Declare Sub InitCommonControls Lib "comctl32" ()

'To show the system settings cahnged restart dialogue
Public Declare Function SHRestartSystemMB Lib "shell32" Alias "#59" (ByVal hOwner As Long, ByVal sExtraPrompt As String, ByVal uFlags As Long) As Long

'Registry constants
Public Const REG_SZ As Long = 1
Public Const REG_DWORD As Long = 4
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
