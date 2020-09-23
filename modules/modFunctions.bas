Attribute VB_Name = "modFunctions"
'><>rbgCODE 2004

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function SearchTreeForFile Lib "IMAGEHLP.DLL" (ByVal lpRootPath As String, ByVal lpInputName As String, ByVal lpOutputName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Public Declare Function GetLogicalDrives Lib "kernel32" () As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Type BROWSEINFO
 hOwner           As Long
 pidlRoot         As Long
 pszDisplayName   As String
 lpszTitle        As String
 ulFlags          As Long
 lpfn             As Long
 lParam           As Long
 iImage           As Long
End Type

Public OurFiles As String 'Returns list of files found in DIR

Global glDoItFile(0 To 99) As String
Global glDoItFolder(0 To 99) As String
Global glDoItOther(0 To 99) As String
Global glDoItURL(0 To 99) As String

Public Function DirExists(strDir As String) As Boolean
 strDir = Dir(strDir, vbDirectory)
 If (strDir = "") Then
  DirExists = False
 Else
  DirExists = True
 End If
End Function

Public Function GetDirectory(frm As Form) As String
Dim bi As BROWSEINFO
Dim pidl As Long
Dim Path$, Pos%
 bi.hOwner = frm.hwnd
 bi.pidlRoot = 0&
 bi.lpszTitle = "Select directory..."
 bi.ulFlags = BIF_RETURNONLYFSDIRS
 pidl = SHBrowseForFolder(bi)
 Path = Space$(256)
 If SHGetPathFromIDList(ByVal pidl, ByVal Path) Then
  Pos = InStr(Path, Chr$(0))
  GetDirectory = Left(Path, Pos - 1)
 End If
 Call CoTaskMemFree(pidl)
End Function

Public Function FileExists(FileName As String) As Boolean
On Error Resume Next
 FileExists = (Dir$(UCase$((FileName))) <> "")
End Function

Public Function readFile(FileName As String) As String()
Dim f As Integer
Dim tmpStr As String
 f = FreeFile()
 FileName = App.Path & "\" & FileName
 Open FileName For Input As f
  tmpStr = Input$(LOF(f), f)
 Close f
 readFile = Split(Replace(tmpStr, """", ""), vbNewLine)
End Function

Public Function addTXT(theStr As String, TheFile As String)
Dim ff
 ff = FreeFile
 TheFile = App.Path & "\" & TheFile
 Open TheFile For Append As #ff
  Write #ff, theStr
 Close #ff
End Function

Public Sub xstart(xpath As String)
 Call ShellExecute(hwnd, "Open", xpath, "", App.Path, 1)
End Sub

