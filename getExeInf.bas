Attribute VB_Name = "modExeParser"
Option Explicit
'MsgBox GetFileVerInfo("C:\windows\notepad.exe")(0)        '版本号
'MsgBox GetFileVerInfo("C:\windows\notepad.exe")(1)        '产品名称
'MsgBox GetFileVerInfo("C:\windows\notepad.exe")(2)        '公司名称
'MsgBox GetFileVerInfo("C:\windows\notepad.exe")(3)        '版权信息
'MsgBox GetFileVerInfo("C:\windows\notepad.exe")(4)        '文件描述

Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function VerLanguageName Lib "kernel32" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
Private Const Syn_kg = " "
Private Const Syn_pzh = "\"

Public Function GetFileVerInfo(FullFileName As String) As String()
    Dim rc     As Long, lDummy       As Long, sBuffer()       As Byte
    Dim lBufferLen     As Long, lVerPointer       As Long
    Dim bytebuffer(260)     As Byte
    Dim Lang_Charset_String     As String
    Dim HexNumber     As Long, Buffer       As String
    Dim i     As Integer, strtemp       As String
    Dim strFileVer(5)     As String
    For i = 0 To 5
        strFileVer(i) = ""
    Next
    lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)
    If lBufferLen < 1 Then
        GetFileVerInfo = strFileVer
        Exit Function
    End If
    ReDim sBuffer(lBufferLen)
    rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
    If rc = 0 Then
        GetFileVerInfo = strFileVer
        Exit Function
    End If
    rc = VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lVerPointer, lBufferLen)
    If rc = 0 Then
        GetFileVerInfo = strFileVer
        Exit Function
    End If
              
    strFileVer(0) = "FileVersion"
    strFileVer(1) = "InternalName"
    strFileVer(2) = "CompanyName"
    strFileVer(3) = "LegalCopyright"
    strFileVer(4) = "FileDescription"
    
    MoveMemory bytebuffer(0), lVerPointer, lBufferLen
    HexNumber = bytebuffer(2) + bytebuffer(3) * &H100 + bytebuffer(0) * &H10000 + bytebuffer(1) * &H1000000
    rc = CLng(bytebuffer(0) + bytebuffer(1) * &H100)
    Lang_Charset_String = Hex(HexNumber)
      
    Do While Len(Lang_Charset_String) < 8
        Lang_Charset_String = "0" & Lang_Charset_String
    Loop
      
    strtemp = String(260, Asc(Syn_kg))
    rc = VerLanguageName(rc, strtemp, CLng(255))
    strFileVer(5) = StripTerminator(strtemp)
    
    strtemp = ""
    For i = 0 To 4
        Buffer = String(260, Asc(Syn_kg))
        strtemp = "\StringFileInfo\" & Lang_Charset_String & Syn_pzh & strFileVer(i)
        rc = VerQueryValue(sBuffer(0), strtemp, lVerPointer, lBufferLen)
        If rc <> 0 Then
            lstrcpy Buffer, lVerPointer
            Buffer = StripTerminator(Buffer)
        Else
            Buffer = ""
        End If
        strFileVer(i) = Buffer
    Next i
    GetFileVerInfo = strFileVer
End Function

Private Function StripTerminator(ByVal sInput As String) As String
    Dim ZeroPos     As Integer
    ZeroPos = InStr(1, sInput, vbNullChar)
    If ZeroPos > 0 Then
        StripTerminator = Left$(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
End Function

Public Function pathStringParser(ByVal Path As String) As String
pathStringParser = IIf(Right(Path, 1) = "\", Path, Path & "\")
End Function

Public Function APPPath(Optional ByVal filename As String) As String
If IsMissing(filename) Then filename = ""
APPPath = pathStringParser(App.Path) & filename
End Function
