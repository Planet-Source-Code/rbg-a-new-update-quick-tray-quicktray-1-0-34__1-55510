Attribute VB_Name = "modFileControls"
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function SearchTreeForFile Lib "IMAGEHLP.DLL" (ByVal lpRootPath As String, ByVal lpInputName As String, ByVal lpOutputName As String) As Long

Public Declare Function GetLogicalDrives Lib "kernel32" () As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Public ManageActiveList As String
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

Sub SaveTextAppend(Path As String, StringName As String)
On Error Resume Next
 Open Path$ For Append As #1
  Print #1, StringName
 Close #1
End Sub
  
Sub SaveTextOutput(Path As String, StringName As String)
On Error Resume Next
 Open Path$ For Output As #1
  Print #1, StringName
 Close #1
End Sub

Public Function CHECKFORDIRECTORY(ByVal DIRECTORYNAME As String) As Boolean
On Error GoTo errhandler
 Select Case m_OBJ_FSO.FolderExists(DIRECTORYNAME)
  Case Is = True: CHECKFORDIRECTORY = True
  Case Else: CHECKFORDIRECTORY = False
 End Select
 Set m_OBJ_FSO = Nothing
 Exit Function
errhandler:
 If Not m_OBJ_FSO Is Nothing Then
  Set m_OBJ_FSO = Nothing
 End If
 MsgBox "Error Description : " & Err.Description & vbCrLf & "Error Number : " & Err.Number & vbCrLf & "Error Source : " & Err.Source, vbCritical + vbOKOnly, "ERROR"
End Function

Public Function DirExists(strDir As String) As Boolean
 strDir = Dir(strDir, vbDirectory)
 DirExists = True
 If (strDir = "") Then DirExists = False
End Function

Public Sub SaveListBox(Directory As String, theList As ListBox)
Dim savelist As Long
On Error Resume Next
 fe = FreeFile
 Open Directory$ For Output As #fe
  For savelist = 0 To theList.ListCount - 1
   bufff = theList.List(savelist)
   bufff = Replace(bufff, Chr(13), "ªä")
   Print #fe, Trim(bufff)
  Next savelist
 Close #fe
End Sub

Public Sub Save2ListBoxes(Directory As String, ListA As ListBox, ListB As ListBox)
    Dim SaveLists As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveLists& = 0 To ListA.ListCount - 1
        Print #1, ListA.List(SaveLists&) & "*" & ListB.List(SaveLists)
    Next SaveLists&
    Close #1
End Sub
Public Sub Save3ListBoxes(Directory As String, ListA As ListBox, ListB As ListBox, ListC As ListBox)
    Dim SaveLists As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveLists& = 0 To ListA.ListCount - 1
        Print #1, ListA.List(SaveLists&) & "*" & ListB.List(SaveLists) & "*" & ListC.List(SaveLists)
    Next SaveLists&
    Close #1
End Sub
Public Sub SaveComboBox(ByVal Directory As String, Combo As ComboBox)
    Dim SaveCombo As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveCombo& = 0 To Combo.ListCount - 1
        Print #1, Combo.List(SaveCombo&)
    Next SaveCombo&
    Close #1
End Sub

Public Function FileGetAttributes(TheFile As String) As Long
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        FileGetAttributes = GetAttr(TheFile$)
    End If
End Function

Public Sub FileSetNormal(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbNormal
    End If
End Sub

Public Sub FileSetReadOnly(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbReadOnly
    End If
End Sub

Public Sub FileSetHidden(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbHidden
    End If
End Sub

Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function

Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub
Function LoadList(lst, File)
On Error GoTo File
lst.Clear
Dim stuff As String
Dim Holdit As Long

 A = CurDir
 ChDir A
 Open File For Input As #1
 Do
    Input #1, stuff$
    stuff$ = Trim(stuff$)
    If stuff$ <> "" Then lst.AddItem stuff$
    Holdit = DoEvents()
  Loop Until EOF(1)
    Close #1
File:
End Function

 Function LoadText(File, TextBoxOrLabel As TextBox)
 On Error GoTo Fle
A = CurDir
ChDir A
fn = FreeFile

Open File For Input As #fn
Do
Line Input #fn, Text$
TheText = TheText & Text$ & vbNewLine
Loop Until EOF(1)
Close #fn
TextBoxOrLabel = TheText
Fle:
 End Function
 
Function LoadTextAppend(File, TextBoxOrLabel As TextBox)
 On Error GoTo Fle
A = CurDir
ChDir A
fn = FreeFile
Open File For Input As #fn
Do
Line Input #fn, Text$
TheText = TheText & Text$ & vbNewLine
Loop Until EOF(1)
Close #fn
TextBoxOrLabel = TextBoxOrLabel & vbNewLine & TheText
Fle:
 End Function
Public Function GetDriveTypes(DriveLetter As String)
    Select Case GetDriveTypes(DriveLetter)
        Case 2
            GetDriveTypes = "Removable"
        Case 3
            GetDriveTypes = "Drive Fixed"
        Case Is = 4
            Debug.Print "Remote"
        Case Is = 5
            Debug.Print "Cd-Rom"
        Case Is = 6
            Debug.Print "Ram disk"
        Case Else
            Debug.Print "Unrecognized"
    End Select
End Function

Private Function FindFile(sFile As String, sRootPath As String) As String
    ' Search for the file specified and retu
    '     rn the full path if found
    Dim sPathBuffer As String
    Dim iEnd As Integer
    
    'Allocate some buffer space (you may nee
    '     d more)
    sPathBuffer = Space(512)
    


    If SearchTreeForFile(sRootPath, sFile, sPathBuffer) Then
        'Strip off the null string that will be
        '     returned following the path name
        iEnd = InStr(1, sPathBuffer, vbNullChar, vbTextCompare)
        sPathBuffer = Left$(sPathBuffer, iEnd - 1)
        FindFile = sPathBuffer
    Else
        FindFile = vbNullString
    End If
End Function

Public Function ExtractAll(ByVal FilePath As String, Optional DefaultExtension As String = "*.*") As String
Dim RetVal As String
Dim MyName As String
Dim SubDir(500) As String

 If Mid$(FilePath, Len(FilePath), 1) <> "\" Then
  FilePath = FilePath + "\"
 End If
 i = 0
On Error GoTo nodir
 MyName = Dir(FilePath, vbDirectory)
 Do While MyName <> ""
  If MyName <> "." And MyName <> ".." And MyName <> "Directories" Then
   If (GetAttr(FilePath & MyName) And vbDirectory) = vbDirectory Then
    SubDir(i) = MyName
    i = i + 1
   End If
  End If
  MyName = Dir
  If MyName <> "." And MyName <> ".." Then
  OurFiles = OurFiles & MyName & "^"
  End If
 Loop
 DoEvents
 MyName = Dir(FilePath + DefaultExtension, vbNormal)
 Do While MyName <> ""
  If Right(LCase(MyName), 3) = Right(DefaultExtension, 3) Then
   RetVal = RetVal + FilePath + MyName + "^"
  End If
  MyName = Dir
  OurFiles = OurFiles & MyName & "^"
 Loop
nodir:
Dim RetVal2 As String
 For t = 0 To i - 1
  RetVal2 = RetVal2 + ExtractAll(FilePath & SubDir(t) + "\")
 Next t
 'ExtractAll = RetVal + RetVal2
 ExtractAll = OurFiles
 Exit Function
End Function

Public Sub Load2listboxes(Directory As String, ListA As ListBox, ListB As ListBox)
    Dim MyString As String, aString As String, bString As String
    On Error Resume Next
    If FileExists(Directory) = True Then
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        aString$ = Left(MyString$, InStr(MyString$, "*") - 1)
        bString$ = Right(MyString$, Len(MyString$) - InStr(MyString$, "*"))
        DoEvents
        ListA.AddItem aString$
        ListB.AddItem bString$
    Wend
    Close #1
    End If
End Sub
Public Sub Load3listboxes(Directory As String, ListA As ListBox, ListB As ListBox, ListC As ListBox)
    Dim MyString As String, aString As String, bString As String, cString As String
    Dim OurStringAry
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        OurStringAry = Split(MyString$, "*")
        aString$ = Left(MyString$, InStr(MyString$, "*") - 1)
        bString$ = Right(MyString$, Len(MyString$) - InStr(MyString$, "*"))
        cString$ = OurStringAry(1)
        DoEvents
        ListA.AddItem aString$
        ListB.AddItem bString$
        ListC.AddItem cString$
    Wend
    Close #1
End Sub


