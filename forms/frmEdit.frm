VERSION 5.00
Begin VB.Form frmEdit 
   Caption         =   "Edit Entry"
   ClientHeight    =   1350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1350
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnFind 
      Caption         =   "..."
      Height          =   250
      Left            =   3555
      TabIndex        =   5
      Top             =   645
      Width           =   300
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   285
      Left            =   2760
      TabIndex        =   4
      Top             =   1020
      Width           =   1095
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   2685
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Path:"
      Height          =   375
      Left            =   255
      TabIndex        =   3
      Top             =   675
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   375
      Left            =   195
      TabIndex        =   2
      Top             =   195
      Width           =   615
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnFind_Click()
 If ManageActiveList = "lstFoldersName" Then
  txtPath = GetDirectory(Me)
 Else
  With cd
   .InitDir = App.Path
   .Filter = "*.*|*.*"
   .FilterIndex = 1
   .DefaultExt = "*"
   .ShowOpen
   If .FileName <> "" Then txtPath.Text = .FileName
  End With
 End If
End Sub

Private Sub btnOk_Click()
 If ManageActiveList = "lstFilesName" Then
  EditItem = frmManage.lstFilesName.ListIndex
  frmManage.lstFilesName.List(EditItem) = frmEdit.txtName.Text
  frmManage.lstFilesPath.List(EditItem) = frmEdit.txtPath.Text
  FilesPath = App.Path & "\files.lst"
  Kill FilesPath
  Open FilesPath For Output As #1
   For Listz& = 0 To frmManage.lstFilesName.ListCount - 1
    Print #1, frmManage.lstFilesName.List(Listz&) & "^" & frmManage.lstFilesPath.List(Listz&)
   Next Listz&
  Close #1
 End If
 If ManageActiveList = "lstFoldersName" Then
  EditItem = frmManage.lstFilesName.ListIndex
  frmManage.lstFoldersName.List(EditItem) = frmEdit.txtName.Text
  frmManage.lstFoldersPath.List(EditItem) = frmEdit.txtPath.Text
  FoldersPath = App.Path & "\folders.lst"
  Kill FoldersPath
  Open FoldersPath For Output As #1
   For Listz& = 0 To frmManage.lstFoldersName.ListCount - 1
    Print #1, frmManage.lstFoldersName.List(Listz&) & "^" & frmManage.lstFoldersPath.List(Listz&)
   Next Listz&
  Close #1
 End If
 If ManageActiveList = "lstWebName" Then
  EditItem = frmManage.lstFilesName.ListIndex
  frmManage.lstWebName.List(EditItem) = frmEdit.txtName.Text
  frmManage.lstWebPath.List(EditItem) = frmEdit.txtPath.Text
  WebPath = App.Path & "\url.lst"
  Kill WebPath
  Open WebPath For Output As #1
   For Listz& = 0 To frmManage.lstWebName.ListCount - 1
    Print #1, frmManage.lstWebName.List(Listz&) & "^" & frmManage.lstWebPath.List(Listz&)
   Next Listz&
  Close #1
 End If
 If ManageActiveList = "lstOtherName" Then
  EditItem = frmManage.lstFilesName.ListIndex
  frmManage.lstOtherName.List(EditItem) = frmEdit.txtName.Text
  frmManage.lstOtherPath.List(EditItem) = frmEdit.txtPath.Text
  OtherPath = App.Path & "\other.lst"
  Kill OtherPath
  Open OtherPath For Output As #1
   For Listz& = 0 To frmManage.lstOtherName.ListCount - 1
    Print #1, frmManage.lstOtherName.List(Listz&) & "^" & frmManage.lstOtherPath.List(Listz&)
   Next Listz&
  Close #1
 End If
 Unload Me
End Sub
