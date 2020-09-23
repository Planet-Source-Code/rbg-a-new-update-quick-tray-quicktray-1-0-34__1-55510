VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Ok!"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   3960
      Width           =   1695
   End
   Begin VB.PictureBox picRBG 
      Height          =   3975
      Left            =   0
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label lblURL 
      Caption         =   "http://rbgCODE.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   3
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmAbout.frx":1B82
      Height          =   3975
      Left            =   2760
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 Me.Icon = frmMain.Icon
 Me.Caption = "About quickTray " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub lblURL_Click()
 xstart "http://rbgCODE.com"
End Sub
