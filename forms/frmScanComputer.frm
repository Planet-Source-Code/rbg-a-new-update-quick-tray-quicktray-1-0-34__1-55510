VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmScanComputer 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Scan Computer"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSelectOnly 
      Caption         =   "Add only Selected Items"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Value           =   1  'Checked
      Width           =   2100
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Stop "
      Height          =   435
      Left            =   4560
      TabIndex        =   13
      Top             =   600
      Width           =   2415
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   3960
      TabIndex        =   12
      Text            =   "Choose your quick type"
      Top             =   3720
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   225
   End
   Begin MSComctlLib.ImageList imgAll 
      Left            =   6360
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScanComputer.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScanComputer.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScanComputer.frx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScanComputer.frx":0CF6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5760
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScanComputer.frx":1148
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScanComputer.frx":159A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScanComputer.frx":19EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScanComputer.frx":1E3E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnGetFiles 
      Caption         =   "Find Files"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   120
      Width           =   2415
   End
   Begin VB.Frame frameOptions 
      Caption         =   "Options"
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4335
      Begin VB.TextBox txtFilesFound 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   9
         Text            =   "0"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtExtension 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Text            =   "*.exe"
         Top             =   840
         Width           =   1095
      End
      Begin VB.DriveListBox drvSelect 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lbl 
         Caption         =   "Select Drive:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lbl 
         Caption         =   "Select Extension:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.CommandButton btnAddFile 
      Caption         =   "Add to"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   3720
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   255
      Left            =   5520
      TabIndex        =   1
      Top             =   840
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4140
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5786
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2205
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   3889
      View            =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      ColHdrIcons     =   "imgAll"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmScanComputer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bolStop As Boolean
Public Enum ImageSizingTypes
 [sizeNone] = 0
 [sizeCheckBox]
 [sizeIcon]
End Enum

Public Enum LedgerColours
 vbledgerWhite = &HF9FEFF
 vbLedgerGreen = &HD0FFCC
 vbLedgerYellow = &HE1FAFF
 vbLedgerRed = &HE1E1FF
 vbLedgerGrey = &HE0E0E0
 vbLedgerBeige = &HD9F2F7
 vbLedgerSoftWhite = &HF7F7F7
 vbledgerPureWhite = &HFFFFFF
End Enum

Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private orderedBy As String
Private m_cHdrIcons As cLVHeaderSortIcons
Private Const SM_CXVSCROLL = 2
Private Const SM_CYHSCROLL = 3
Private Const SM_CYVSCROLL = 20
Private Const SM_CXHSCROLL = 21

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Function ExtractAll(ByVal FilePath As String, Optional DefaultExtension As String = "*.*") As String
Dim i, t
Dim RetVal As String
Dim MyName As String
Dim SubDir(500) As String
Static cnt As Long
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
 Loop
 DoEvents
 If bolStop = True Then GoTo cExit
 MyName = Dir(FilePath + DefaultExtension, vbNormal)
 Do While MyName <> ""
  If Right(LCase(MyName), 3) = Right(txtExtension, 3) Then
   RetVal = RetVal + FilePath + MyName + "^"
   cnt = cnt + 1
   Status.Panels(1).Text = "Searching files " + cnt
   Status.Refresh
   DoEvents
  End If
  MyName = Dir
 Loop
nodir:
Dim RetVal2 As String
 For t = 0 To i - 1
  RetVal2 = RetVal2 + ExtractAll(FilePath & SubDir(t) + "\")
 Next t
cExit:
 ExtractAll = RetVal + RetVal2
 Exit Function
End Function

Private Sub btnAddFile_Click()
 If bolStop = False And btnGetFiles.Enabled = False Then Exit Sub
 If chkSelectOnly.Value = 0 Then
  If MsgBox("Are you then you would like to add the entire list box?  If not please check add only selected Items!", vbYesNo) = vbNo Then Exit Sub
 End If
 If cmbType.ListIndex = -1 Then
  MsgBox "Choose your quick type", vbOKOnly, App.Title
  Exit Sub
 End If
 If lv.ListItems.Count = 0 Then
  MsgBox "No Results found to add", vbOKOnly, App.Title
  Exit Sub
 End If
 Call SaveSearchContent
 MsgBox "Adding Completed!"
End Sub

Private Sub btnGetFiles_Click()
Dim x
Dim i
Dim theList
Dim iMax As Long
Dim yy() As String
 If drvSelect.ListIndex = -1 Then
  MsgBox "Select one of the drive", vbOKOnly, App.Title
  Exit Sub
 End If
 If txtExtension.Text = "" Then
  MsgBox "Extension cannot be empty", vbOKOnly, App.Title
  Exit Sub
 End If
 bolStop = False
 btnGetFiles.Enabled = False
 Status.SimpleText = "Please Wait while I read your harddrive!"
 x = ExtractAll(Mid(drvSelect.Drive, 1, 2), txtExtension)
 bolStop = False
 If Len(x) = 0 Then GoTo NoFilesFound
 theList = x
 If Right(theList, 1) = "^" Then theList = Mid(theList, 1, (Len(theList) - 1))
 yy = Split(x, "^")
'><> Use a listbox instead of a combo box!!
 loadLV lv, yy
'><>Old Code removed for combobox!
 btnGetFiles.Enabled = True
 Exit Sub
NoFilesFound:
 btnGetFiles.Enabled = True
 Status.SimpleText = "0 files found..."
End Sub

Private Sub btnCancel_Click()
 bolStop = True
End Sub

Private Sub Form_Activate()
 Call SetListViewLedger(lv, vbWhite, &HF5F5F5, sizeIcon)
End Sub

Private Sub Form_Load()
 Set m_cHdrIcons = New cLVHeaderSortIcons
 Set m_cHdrIcons.ListView = lv
      
 Status.Panels(1).Text = "Ready"
 AddProgBar pBar, Status, 2
 Visible = True
 Refresh
 Show
 
 With lv
  .ListItems.Clear
  .ColumnHeaders.Clear
  .ColumnHeaders.Add , , "ID#"
  .ColumnHeaders(1).Tag = "number"
  .ColumnHeaders.Add , , "File"
  .ColumnHeaders(2).Tag = "string"
  .View = lvwReport
  .Sorted = False
 End With
 
 cmbType.AddItem "File"
 cmbType.AddItem "Folder"
 cmbType.AddItem "Other"
 cmbType.AddItem "Web Sites"
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If bolStop = False And btnGetFiles.Enabled = False Then
  Call btnCancel_Click
 End If
 Unload frmMain
 Load frmMain
End Sub

Private Sub lv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim i As Integer
 If (lv.SortKey = ColumnHeader.Index - 1) Then
  ColumnHeader.Tag = Not Val(ColumnHeader.Tag)
 End If
 lv.SortOrder = Abs(Val(ColumnHeader.Tag))
 lv.SortKey = ColumnHeader.Index - 1
 lv.Sorted = True
 Call m_cHdrIcons.SetHeaderIcons(lv.SortKey, lv.SortOrder)
End Sub

Public Function loadLV(lv As ListView, arr() As String, Optional fieldNumber As Integer, Optional FieldValue As String)
Dim i As Integer
Dim j As Integer
Dim tmpStr() As String
 With lv
  .ListItems.Clear
  .ColumnHeaders.Clear
  .ColumnHeaders.Add = "ID#"
  .ColumnHeaders(1).Tag = "number"
  .ColumnHeaders.Add , , "File"
  .ColumnHeaders(2).Tag = "string"
  .View = lvwReport
  .Sorted = False
 End With
 txtFilesFound.Text = "0"
 pBar.Max = UBound(arr)
 pBar.Visible = True
 For i = 0 To UBound(arr) - 1
  DoEvents
  On Error Resume Next
  lv.ListItems.Add , , i + 1 ' & "", , 4
  lv.ListItems(lv.ListItems.Count).ListSubItems.Add , , arr(i) & ""
  txtFilesFound.Text = Val(txtFilesFound.Text) + 1
  Status.Panels(3).Text = CInt(i / UBound(arr) * 100) & "%"
  pBar.Value = i
  DoEvents
 Next i
 pBar.Visible = False
 Call lvAutosizeControl(lv)
End Function

Public Sub lvAutosizeControl(lv As ListView)
Dim col2adjust As Long
 For col2adjust = 0 To lv.ColumnHeaders.Count - 1
  Call SendMessage(lv.hWnd, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE_USEHEADER)
 Next col2adjust
End Sub

Public Sub SetListViewLedger(lv As ListView, Bar1Color As LedgerColours, Bar2Color As LedgerColours, nSizingType As ImageSizingTypes)
Dim iBarHeight  As Long
Dim lBarWidth   As Long
Dim diff        As Long
Dim twipsy      As Long
 iBarHeight = 0
 lBarWidth = 0
 diff = 0
On Local Error GoTo SetListViewColor_Error
 twipsy = Screen.TwipsPerPixelY
 If lv.View = lvwReport Then
  With lv
   .Picture = Nothing
   .Refresh
   .Visible = 1
   .PictureAlignment = lvwTile
   lBarWidth = .Width
  End With  ' lv
  With Picture1
   .AutoRedraw = False
   .Picture = Nothing
   .BackColor = vbWhite
   .Height = 1
   .AutoRedraw = True
   .BorderStyle = vbBSNone
   .ScaleMode = vbTwips
   .Top = Me.Top - 10000
   .Width = Screen.Width
   .Visible = False
   .Font = lv.Font
   With .Font
    .Bold = lv.Font.Bold
    .Charset = lv.Font.Charset
    .Italic = lv.Font.Italic
    .Name = lv.Font.Name
    .Strikethrough = lv.Font.Strikethrough
    .Underline = lv.Font.Underline
    .Weight = lv.Font.Weight
    .Size = lv.Font.Size
   End With  'Picture1.Font
   iBarHeight = .TextHeight("W")
   Select Case nSizingType
    Case sizeNone:
     iBarHeight = iBarHeight + twipsy
    Case sizeCheckBox:
     If (iBarHeight \ twipsy) > 18 Then
      iBarHeight = iBarHeight + twipsy
     Else
      diff = 18 - (iBarHeight \ twipsy)
      iBarHeight = iBarHeight + (diff * twipsy) + (twipsy * 1)
     End If
    Case sizeIcon:
     diff = ImageList1.ImageHeight - (iBarHeight \ twipsy)
     iBarHeight = iBarHeight + (diff * twipsy) + (twipsy * 1)
   End Select
   .Height = iBarHeight * 2
   .Width = lBarWidth * 2
   Picture1.Line (0, 0)-(lBarWidth * 2, iBarHeight), Bar1Color, BF
   Picture1.Line (0, iBarHeight)-(lBarWidth * 2, iBarHeight * 2), Bar2Color, BF
   .AutoSize = True
   .Refresh
  End With
  lv.Refresh
  lv.Picture = Picture1.Image
 Else
  lv.Picture = Nothing
 End If  'lv.View = lvwReport
SetListViewColor_Exit:
 On Local Error GoTo 0
Exit Sub
SetListViewColor_Error:
 With lv
  .Picture = Nothing
  .Refresh
 End With
 Resume SetListViewColor_Exit
End Sub

Private Sub SaveSearchContent()
Dim itemPath As String
Dim Listz As Long
Dim lpos As Variant
Dim lfilename As String
Dim itemContent As String
Dim ff
 ff = FreeFile
On Error GoTo cErr:
 Select Case cmbType.ListIndex
  Case 0
   itemPath = App.Path & "\files.lst"
  Case 1
   itemPath = App.Path & "\folders.lst"
  Case 2
   itemPath = App.Path & "\other.lst"
  Case 3
   itemPath = App.Path & "\url.lst"
 End Select
 If itemPath = "" Then Exit Sub
 If Dir(itemPath) = "" Then
  Open itemPath For Output As #ff
 Else
  Open itemPath For Append As #ff
 End If
 For Listz = 1 To lv.ListItems.Count
  If (chkSelectOnly = 1 And lv.ListItems(Listz).Selected = True) Or chkSelectOnly = 0 Then
   lfilename = ""
   itemContent = lv.ListItems(Listz).SubItems(1)
   lpos = InStrRev(itemContent, "\")
   If lpos > 0 Then
    lfilename = Mid(itemContent, lpos + 1)
   End If
   Print #ff, lfilename & "^" & lv.ListItems(Listz).SubItems(1)
  End If
 Next Listz
cErr:
 Close #ff
End Sub
