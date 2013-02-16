Attribute VB_Name = "Registry"
Option Explicit

Private Const HKEY_CLASSES_ROOT = -2147483648#
Private Const HKEY_CURRENT_USER = -2147483647#
Private Const HKEY_LOCAL_MACHINE = -2147483646#
Private Const HKEY_USERS = -2147483645#

Private Const REG_NONE = 0                       ' No value type
Private Const REG_SZ = 1&                        '字符串值
Private Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Private Const REG_BINARY = 3&                    '二进制值
Private Const REG_DWORD = 4&                     'DWORD 值
Private Const REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
Private Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Private Const REG_LINK = 6                       ' Symbolic Link (unicode)
Private Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings

Private Const ERROR_NONE = 0
Private Const ERROR_BADKEY = 2
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_SUCCESS = 0

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.（已修改为Byval）
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.（已修改为Byval）
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Sub CreateUninstallInformation()
Dim Handle As Long
Debug.Print RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\QuickClip2", Handle)
Debug.Print Handle
Debug.Print RegSetValueEx(Handle, "My Value", 0, REG_SZ, "Test", 255)
Debug.Print RegCloseKey(Handle)
End Sub


'例程
'    Sub Main()
'      Dim nKeyHandle As Long, nValueType As Long, nLength As Long
'      Dim sValue As String
'      sValue = "I am a winner!"
'      Call RegCreateKey(HKEY_CURRENT_USER, "New Registry Key", nKeyHandle)
'      Call RegSetValueEx(nKeyHandle, "My Value", 0, REG_SZ, sValue, 255)
'      sValue = Space(255)
'      nLength = 255
'      Call RegQueryValueEx(nKeyHandle, "My Value", 0, nValueType, sValue, nLength)
'      MsgBox sValue
'      Call RegDeleteValue(nKeyHandle, "My Value")
'      Call RegDeleteKey(HKEY_CURRENT_USER, "New Registry Key")
'      Call RegCloseKey(nKeyHandle)
'    End Sub
