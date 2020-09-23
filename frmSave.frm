VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save..."
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2220
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3060
      TabIndex        =   4
      Top             =   2220
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   2100
      TabIndex        =   3
      Top             =   0
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "*.jpg"
      Top             =   1860
      Width           =   4035
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim path As String
    
    path = Me.File1.path
    If Right(path, 1) <> "\" Then
        path = path & "\"
    End If
    
    On Error Resume Next
    SavePicture frmImage.Dest.Image, path & Me.Text1.Text
    Me.File1.Refresh
    
    Me.Hide
    frmMain.Enabled = True
    
End Sub

Private Sub Command2_Click()
    Me.Hide
    frmMain.Enabled = True
    
End Sub

Private Sub Dir1_Change()
    Me.File1.path = Me.Dir1.path
End Sub

Private Sub Drive1_Change()
    Me.Dir1.path = Me.Drive1
    
End Sub

Private Sub File1_Click()
    Me.Text1.Text = Me.File1.List(Me.File1.ListIndex)
End Sub
