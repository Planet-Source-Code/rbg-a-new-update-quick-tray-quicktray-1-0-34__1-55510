VERSION 5.00
Begin VB.Form frmManage 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Manage"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7095
      Left            =   0
      TabIndex        =   5
      Top             =   -120
      Width           =   8175
      Begin VB.CommandButton cmdEditItem 
         Caption         =   "Edit Item"
         Height          =   375
         Left            =   2280
         TabIndex        =   16
         Top             =   6600
         Width           =   1455
      End
      Begin VB.Frame Frame5 
         Caption         =   "Web"
         Height          =   2895
         Left            =   3960
         TabIndex        =   14
         Top             =   3600
         Width           =   3495
         Begin VB.ListBox lstWebName 
            Height          =   2400
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Folders"
         Height          =   3255
         Left            =   3960
         TabIndex        =   12
         Top             =   240
         Width           =   3495
         Begin VB.ListBox lstFoldersName 
            Height          =   2595
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Other"
         Height          =   2895
         Left            =   240
         TabIndex        =   10
         Top             =   3600
         Width           =   3495
         Begin VB.ListBox lstOtherName 
            Height          =   2400
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Files"
         Height          =   3255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   3495
         Begin VB.ListBox lstFilesName 
            Height          =   2595
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   5880
         TabIndex        =   7
         Top             =   6600
         Width           =   1575
      End
      Begin VB.CommandButton cmdRemoveItem 
         Caption         =   "Remove Item"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   6600
         Width           =   1815
      End
   End
   Begin VB.TextBox txtActiveList 
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   2520
      Width           =   2055
   End
   Begin VB.ListBox lstWebPath 
      Height          =   255
      Left            =   5640
      TabIndex        =   3
      Top             =   7200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstOtherPath 
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox lstFilesPath 
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox lstFoldersPath 
      Height          =   1815
      Left            =   2760
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEditItem_Click()
Load frmEdit
 If ManageActiveList = "lstFilesName" Then
  EditItem = lstFilesName.ListIndex
  frmEdit.txtName.Text = frmManage.lstFilesName.List(EditItem)
  frmEdit.txtPath.Text = frmManage.lstFilesPath.List(EditItem)
 End If
 If ManageActiveList = "lstFoldersName" Then
  EditItem = lstFoldersName.ListIndex
  frmEdit.txtName.Text = frmManage.lstFoldersName.List(EditItem)
  frmEdit.txtPath.Text = frmManage.lstFoldersPath.List(EditItem)
 End If
 If ManageActiveList = "lstWebName" Then
  EditItem = lstWebName.ListIndex
  frmEdit.txtName.Text = frmManage.lstWebName.List(EditItem)
  frmEdit.txtPath.Text = frmManage.lstWebPath.List(EditItem)
 End If
 If ManageActiveList = "lstOtherName" Then
  EditItem = lstOtherName.ListIndex
  frmEdit.txtName.Text = frmManage.lstOtherName.List(EditItem)
  frmEdit.txtPath.Text = frmManage.lstOtherPath.List(EditItem)
 End If
 frmEdit.Show 1
End Sub

Private Sub cmdExit_Click()
 Unload Me
End Sub

Private Sub cmdRemoveItem_Click()
Dim ff
 ff = FreeFile
 If ManageActiveList = "lstFilesName" Then
  RemoveItem = lstFilesName.ListIndex
  If RemoveItem > -1 Then
   lstFilesName.RemoveItem (RemoveItem)
   lstFilesPath.RemoveItem (RemoveItem)
   FilesPath = App.Path & "\files.lst"
   Kill FilesPath
   Open FilesPath For Output As #ff
   For Listz& = 0 To frmManage.lstFilesName.ListCount - 1
    Print #ff, frmManage.lstFilesName.List(Listz&) & "^" & frmManage.lstFilesPath.List(Listz&)
    Next Listz&
   Close #ff
  End If
 End If
 If ManageActiveList = "lstFoldersName" Then
  RemoveItem = lstFoldersName.ListIndex
  If RemoveItem > -1 Then
   lstFoldersName.RemoveItem (RemoveItem)
   lstFoldersPath.RemoveItem (RemoveItem)
   FoldersPath = App.Path & "\folders.lst"
   Kill FoldersPath
   Open FoldersPath For Output As #ff
   For Listz& = 0 To frmManage.lstFoldersName.ListCount - 1
    Print #ff, frmManage.lstFoldersName.List(Listz&) & "^" & frmManage.lstFoldersPath.List(Listz&)
   Next Listz&
   Close #ff
  End If
 End If
 If ManageActiveList = "lstWebName" Then
  RemoveItem = lstWebName.ListIndex
  If RemoveItem > -1 Then
   lstWebName.RemoveItem (RemoveItem)
   lstWebPath.RemoveItem (RemoveItem)
   WebPath = App.Path & "\url.lst"
   Kill WebPath
   Open WebPath For Output As #ff
    For Listz& = 0 To frmManage.lstWebName.ListCount - 1
     Print #ff, frmManage.lstWebName.List(Listz&) & "^" & frmManage.lstWebPath.List(Listz&)
    Next Listz&
   Close #ff
  End If
 End If
 If ManageActiveList = "lstOtherName" Then
  If RemoveItem > -1 Then
   RemoveItem = lstOtherName.ListIndex
   lstOtherName.RemoveItem (RemoveItem)
   lstOtherPath.RemoveItem (RemoveItem)
   OtherPath = App.Path & "\other.lst"
   Kill OtherPath
   Open OtherPath For Output As #ff
    For Listz& = 0 To frmManage.lstOtherName.ListCount - 1
     Print #ff, frmManage.lstOtherName.List(Listz&) & "^" & frmManage.lstOtherPath.List(Listz&)
    Next Listz&
   Close #ff
  End If
 End If
 Call LoadLists
End Sub

Private Sub LoadLists()
Dim ff
 ff = FreeFile
 frmManage.lstFilesName.Clear
 frmManage.lstFilesPath.Clear
 frmManage.lstFoldersName.Clear
 frmManage.lstFoldersPath.Clear
 frmManage.lstWebName.Clear
 frmManage.lstWebPath.Clear
 frmManage.lstOtherName.Clear
 frmManage.lstOtherPath.Clear
 Open App.Path & "\files.lst" For Input As #ff
  While Not EOF(ff)
   Input #ff, MyString$
   OurStringAry = Split(MyString$, "^")
   aString$ = OurStringAry(0)
   bString$ = OurStringAry(1)
   DoEvents
   frmManage.lstFilesName.AddItem aString$
   frmManage.lstFilesPath.AddItem bString$
  Wend
 Close #ff
 Open App.Path & "\folders.lst" For Input As #ff
  While Not EOF(ff)
   Input #ff, MyString$
   OurStringAry = Split(MyString$, "^")
   aString$ = OurStringAry(0)
   bString$ = OurStringAry(1)
   DoEvents
   frmManage.lstFoldersName.AddItem aString$
   frmManage.lstFoldersPath.AddItem bString$
  Wend
 Close #ff
 Open App.Path & "\url.lst" For Input As #ff
  While Not EOF(ff)
   Input #ff, MyString$
   OurStringAry = Split(MyString$, "^")
   aString$ = OurStringAry(0)
   bString$ = OurStringAry(1)
   DoEvents
   frmManage.lstWebName.AddItem aString$
   frmManage.lstWebPath.AddItem bString$
  Wend
 Close #ff
 Open App.Path & "\other.lst" For Input As #ff
  While Not EOF(ff)
   Input #ff, MyString$
   OurStringAry = Split(MyString$, "^")
   aString$ = OurStringAry(0)
   bString$ = OurStringAry(1)
   DoEvents
   frmManage.lstOtherName.AddItem aString$
   frmManage.lstOtherPath.AddItem bString$
  Wend
 Close #1
End Sub

Private Sub Form_Load()
 Call LoadLists
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload frmMain
 Load frmMain
End Sub

Private Sub lstFilesName_Click()
 ManageActiveList = "lstFilesName"
End Sub

Private Sub lstFoldersName_Click()
 ManageActiveList = "lstFoldersName"
End Sub

Private Sub lstOtherName_Click()
 ManageActiveList = "lstOtherName"
End Sub

Private Sub lstWebName_Click()
 ManageActiveList = "lstWebName"
End Sub
