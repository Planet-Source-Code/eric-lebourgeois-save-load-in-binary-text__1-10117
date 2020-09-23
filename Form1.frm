VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Basic File Commands"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "New"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   2760
      Pattern         =   "*.txt"
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   4215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Refresh All"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Rename"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2990
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      RightMargin     =   1
      TextRTF         =   $"Form1.frx":27A2
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Qwer
Interior As String
End Type
Private beans As Qwer
Private Sub Command1_Click()
a = InputBox("Input Filename!", "Save text as...")
Open Dir1.Path & "\" & a & ".txt" For Binary As #1
beans.Interior = Text1.Text
Put #1, 1, beans
Close #1
End Sub

Private Sub Command2_Click()
Text1.Text = ""
End Sub

Private Sub Command3_Click()
'On Error GoTo closit
a = InputBox("Rename file to...", "Rename file:")
If Right(Dir1.Path, 1) = "\" Then Name Dir1.Path & File1.FileName As Dir1.Path & a & ".txt": GoTo nextstep2
Name Dir1.Path & "\" & File1.FileName As Dir1.Path & "\" & a & ".txt"
nextstep2:
Exit Sub
closit:
MsgBox "Error!"
End Sub

Private Sub Command4_Click()
Dir1.Refresh
Drive1.Refresh
File1.Refresh
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo fixit
Dir1.Path = Drive1.Drive
Exit Sub
fixit:
MsgBox "Error! Couldn't read the specified Device"
End Sub

Private Sub File1_DblClick()
If Right(Dir1.Path, 1) = "\" Then Open Dir1.Path & File1.FileName For Binary As #1: GoTo nextstep
Open Dir1.Path & "\" & File1.FileName For Binary As #1
nextstep:
Get #1, 1, beans
Close #1
Text1.Text = ""
Text1.Text = beans.Interior
End Sub
