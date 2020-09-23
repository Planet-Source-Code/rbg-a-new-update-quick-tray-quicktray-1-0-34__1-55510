VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4575
   ClientLeft      =   840
   ClientTop       =   615
   ClientWidth     =   3375
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4575
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkFormOnTop 
      Caption         =   "Stay On Top?"
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CheckBox chkAutoBoot 
      Caption         =   "Start with System."
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   2880
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add Item"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Frame frameAddItem 
      Caption         =   "Add a quick Tray Item"
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton btnBrowse 
         Caption         =   "..."
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   2040
         Width           =   255
      End
      Begin VB.TextBox txtFile 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtDescription 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   3015
      End
      Begin VB.ComboBox cmbItemType 
         Height          =   315
         ItemData        =   "frmMain.frx":27A2
         Left            =   240
         List            =   "frmMain.frx":27A4
         TabIndex        =   1
         Text            =   "Choose your quick Type"
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label lblFileURL 
         Caption         =   "quickTray File or Folder"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label lbl 
         Caption         =   "quickTray Description"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label lbl 
         Caption         =   "quickTray Type"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "^^Or Drag and Drop Item there^^"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   3240
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line3 
      X1              =   3240
      X2              =   3240
      Y1              =   3360
      Y2              =   4200
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   3360
      Y2              =   4200
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3240
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Menu menuquickTray 
      Caption         =   "quickTray"
      Begin VB.Menu menuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu menuManage 
         Caption         =   "Manage"
      End
      Begin VB.Menu menuScanEXE 
         Caption         =   "Scan for Files"
      End
      Begin VB.Menu mnuendsep 
         Caption         =   "-"
      End
      Begin VB.Menu menuDoItFile 
         Caption         =   "quickFILES"
         Index           =   0
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu menuDoItFolder 
         Caption         =   "quickFOLDERS"
         Index           =   0
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu menuDoItOther 
         Caption         =   "quickOTHERS"
         Index           =   0
      End
      Begin VB.Menu mnuLine4 
         Caption         =   "-"
      End
      Begin VB.Menu menuDoItURL 
         Caption         =   "quickURLs"
         Index           =   0
      End
      Begin VB.Menu mnuLine5 
         Caption         =   "-"
      End
      Begin VB.Menu menuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flagFile As Integer
Dim flagFolder As Integer
Dim flagOther As Integer
Dim flagURL As Integer
Dim i As Integer

Private Sub btnAdd_Click()
On Error Resume Next
 If txtDescription.Text = "" Or txtFile.Text = "" Or cmbItemType.ListIndex = -1 Then GoTo noAdd
 Select Case cmbItemType.ListIndex
  Case 0: flag = flagFile
  Case 1: flag = flagFolder
  Case 2: flag = flagOther
  Case 2: flag = flagURL
 End Select
 If flag = 0 Then
  Select Case cmbItemType.ListIndex
   Case 0: Me.menuDoItFile(0).Visible = True
   Case 1: Me.menuDoItFolder(0).Visible = True
   Case 2: Me.menuDoItOther(0).Visible = True
   Case 3: Me.menuDoItURL(0).Visible = True
  End Select
  flag = flag + 1
  Select Case cmbItemType.ListIndex
   Case 0
    Load menuDoItFile(flag)
    menuDoItFile(flag).Enabled = True
    menuDoItFile(flag).Visible = True
    menuDoItFile(flag).Caption = txtDescription
   Case 1
    Load menuDoItFolder(flag)
    menuDoItFolder(flag).Enabled = True
    menuDoItFolder(flag).Visible = True
    menuDoItFolder(flag).Caption = txtDescription
   Case 2
    Load menuDoItOther(flag)
    menuDoItOther(flag).Enabled = True
    menuDoItOther(flag).Visible = True
    menuDoItOther(flag).Caption = txtDescription
   Case 3
    Load menuDoItURL(flag)
    menuDoItURL(flag).Enabled = True
    menuDoItURL(flag).Visible = True
    menuDoItURL(flag).Caption = txtDescription
  End Select
 Else
  Select Case cmbItemType.ListIndex
   Case 0
    Load menuDoItFile(flag)
    menuDoItFile(flag).Enabled = True
    menuDoItFile(flag).Visible = True
    menuDoItFile(flag).Caption = txtDescription
   Case 1
    Load menuDoItFolder(flag)
    menuDoItFolder(flag).Enabled = True
    menuDoItFolder(flag).Visible = True
    menuDoItFolder(flag).Caption = txtDescription
   Case 2
    Load menuDoItOther(flag)
    menuDoItOther(flag).Enabled = True
    menuDoItOther(flag).Visible = True
    menuDoItOther(flag).Caption = txtDescription
   Case 3
    Load menuDoItURL(flag)
    menuDoItURL(flag).Enabled = True
    menuDoItURL(flag).Visible = True
    menuDoItURL(flag).Caption = txtDescription
  End Select
 End If
'><>Update TXT lists
 Select Case cmbItemType.ListIndex
  Case 0:    addTXT txtDescription & "^" & txtFile, "Files.lst"
  Case 1:    addTXT txtDescription & "^" & txtFile, "Folders.lst"
  Case 2:    addTXT txtDescription & "^" & txtFile, "oTher.lst"
  Case 3:    addTXT txtDescription & "^" & txtFile, "url.lst"
 End Select
'><>Getting There!!!
 flag = flag + 1
 txtDescription.Text = ""
 txtFile.Text = ""
'><>Update your lists!!
 Select Case cmbItemType.ListIndex
  Case 0:    flagFile = flag
  Case 1:    flagFolder = flag
  Case 2:    flagOther = flag
  Case 3:    flagURL = flag
 End Select
'><>And Im spent!
 Exit Sub
noAdd:
 MsgBox "Please ensure that all needed fields are filled in!"
End Sub

Private Sub btnBrowse_Click()
 If cmbItemType.ListIndex = -1 Then
  MsgBox "Please Specify the type of file you are openeing!"
  Exit Sub
 End If
 If cmbItemType.ListIndex = 1 Then
  txtFile = GetDirectory(Me)
 Else
  With cd
   .InitDir = App.Path
   .Filter = "*.*|*.*"
   .FilterIndex = 1
   .DefaultExt = "*"
   .ShowOpen
   If .FileName <> "" Then txtFile.Text = .FileName
  End With
 End If
End Sub

Private Sub chkAutoBoot_Click()
 SaveSetting App.EXEName, "start", "autoBoot", chkAutoBoot
 Select Case chkAutoBoot
  Case 0: DoNotStartUp App.Path & "\" & App.EXEName & ".exe", App.EXEName
  Case 1: DoStartUp App.Path & "\" & App.EXEName & ".exe", App.EXEName
 End Select
End Sub

Private Sub chkFormOnTop_Click()
 SaveSetting App.EXEName, "start", "stayOnTop", chkFormOnTop
 Select Case chkAutoBoot
  Case 0: AlwaysOnTop Me, False
  Case 1: AlwaysOnTop Me, True
 End Select
End Sub

Private Sub cmbItemType_Click()
 Select Case cmbItemType.ListIndex
  Case 0
   txtFile.Enabled = False
   lblFileURL.Caption = "quickTray File"
  Case 1
   txtFile.Enabled = False
   lblFileURL.Caption = "quickTray Folder"
  Case 2
   txtFile.Enabled = False
   lblFileURL.Caption = "quickTray (file)"
  Case 3
   txtDescription = "URL Name"
   txtFile.Enabled = True: txtFile.Text = "http://rbgCODE.com"
   lblFileURL.Caption = "quickTray URL"
  Case Else
   txtFile.Enabled = False
   lblFileURL.Caption = "Choose your quick Type"
 End Select
End Sub

Private Sub Form_Load()
 Me.Visible = False
 loadDefaults
 Me.Caption = "quick Tray " & App.Major & "." & App.Minor & "." & App.Revision
 Me.Refresh
 With nid
  .cbSize = Len(nid)
  .hWnd = Me.hWnd
  .uId = vbNull
  .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  .uCallBackMessage = WM_MOUSEMOVE
  .hIcon = Me.Icon
  .szTip = Me.Caption & vbNullChar
 End With
 Shell_NotifyIcon NIM_ADD, nid
 cmbItemType.AddItem "File"
 cmbItemType.AddItem "Folder"
 cmbItemType.AddItem "Other"
 cmbItemType.AddItem "Web Sites"
 chkAutoBoot = GetSetting(App.EXEName, "start", "autoboot", 0)
 chkFormOnTop = GetSetting(App.EXEName, "start", "stayOnTop", 0)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Result, Action As Long
 If Me.ScaleMode = vbPixels Then
  Action = x
 Else
  Action = x / Screen.TwipsPerPixelX
 End If
 Select Case Action
  Case WM_LBUTTONDBLCLK
   Me.WindowState = vbNormal
   Result = SetForegroundWindow(Me.hWnd)
   Me.Show
  Case WM_RBUTTONUP
   Result = SetForegroundWindow(Me.hWnd)
   PopupMenu menuquickTray
 End Select
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim numFiles As Integer
On Error GoTo badMonkey
 numFiles = Data.Files.Count
 Me.Caption = Data.Files(1)
 If numFiles > 1 Then
  MsgBox "Only 1 file at a time should be dropped.", vbOKOnly
  Exit Sub
 End If
'><>Time to change to work with my code!!! thanks GameBuzz.net
 Select Case (GetAttr(Data.Files(1)))
  Case vbDirectory
   cmbItemType.ListIndex = 1
  Case 32 'which means a different file!
   cmbItemType.ListIndex = 0
  Case Else
   'MsgBox (GetAttr(Data.Files(1)))' most likely a file?!?
   cmbItemType.ListIndex = 1
 End Select
 txtDescription.Text = Data.Files(1)
 txtFile.Text = Data.Files(1)
 Exit Sub
badMonkey:
 MsgBox Err.Description, vbCritical, Err.Number
 Resume Next
End Sub

Private Sub Form_Terminate()
 xstart App.Path & "\" & App.EXEName
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub menuAbout_Click()
 frmAbout.Show 1
End Sub

Private Sub menuDoItFile_Click(Index As Integer)
 xstart (glDoItFile(Index))
End Sub

Private Sub menuDoItFolder_Click(Index As Integer)
 xstart (glDoItFolder(Index))
End Sub

Private Sub menuDoItOther_Click(Index As Integer)
 xstart (glDoItOther(Index))
End Sub

Private Sub menuDoItURL_Click(Index As Integer)
 xstart (glDoItURL(Index))
End Sub

Private Sub menuExit_Click()
 Unload Me: End
End Sub

Public Function loadDefaults()
Dim i
Dim varArray() As String
On Error Resume Next
 If FileExists(App.Path & "\Files.lst") = True Then
  tmpInfo = readFile("Files.lst")
  For i = 0 To UBound(tmpInfo) - 1
   'tmpInfo(I) = Mid(tmpInfo(I), 2, Len(tmpInfo(I)) - 2)
   varArray = Split(tmpInfo(i), "^")
   flag = i
   Load menuDoItFile(flag)
   menuDoItFile(flag).Enabled = True
   menuDoItFile(flag).Visible = True
   menuDoItFile(flag).Caption = varArray(0)
   flagFile = flag
   glDoItFile(i) = varArray(1)
  Next i
 End If
 If FileExists(App.Path & "\Folders.lst") = True Then
  tmpInfo = readFile("Folders.lst")
  For i = 0 To UBound(tmpInfo) - 1
   'tmpInfo(I) = Mid(tmpInfo(I), 2, Len(tmpInfo(I)) - 2)
   varArray = Split(tmpInfo(i), "^")
   flag = i
   Load menuDoItFolder(flag)
   menuDoItFolder(flag).Enabled = True
   menuDoItFolder(flag).Visible = True
   menuDoItFolder(flag).Caption = varArray(0)
   flagFolder = flag
   glDoItFolder(i) = varArray(1)
  Next i
 End If
 If FileExists(App.Path & "\oTher.lst") = True Then
  tmpInfo = readFile("oTher.lst")
  For i = 0 To UBound(tmpInfo) - 1
   'tmpInfo(I) = Mid(tmpInfo(I), 2, Len(tmpInfo(I)) - 2)
   varArray = Split(tmpInfo(i), "^")
   flag = i
   Load menuDoItOther(flag)
   menuDoItOther(flag).Enabled = True
   menuDoItOther(flag).Visible = True
   menuDoItOther(flag).Caption = varArray(0)
   flagOther = flag
   glDoItOther(i) = varArray(1)
  Next i
 End If
 If FileExists(App.Path & "\url.lst") = True Then
  tmpInfo = readFile("url.lst")
  For i = 0 To UBound(tmpInfo) - 1
   'tmpInfo(I) = Mid(tmpInfo(I), 2, Len(tmpInfo(I)) - 2)
   varArray = Split(tmpInfo(i), "^")
   flag = i
   Load menuDoItURL(flag)
   menuDoItURL(flag).Enabled = True
   menuDoItURL(flag).Visible = True
   menuDoItURL(flag).Caption = varArray(0)
   flagURL = flag
   glDoItURL(i) = varArray(1)
  Next i
 End If
End Function

Private Sub menuHide_Click()
 Me.Visible = False
End Sub

Private Sub menuManage_Click()
 frmManage.Show 1
End Sub

Private Sub menuScanEXE_Click()
 frmScanComputer.Show
End Sub
