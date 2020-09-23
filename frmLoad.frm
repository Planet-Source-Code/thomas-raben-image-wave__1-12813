VERSION 5.00
Begin VB.Form frmLoad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Load..."
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   2100
      TabIndex        =   2
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3060
      TabIndex        =   1
      Top             =   1860
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   1860
      Width           =   975
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim path As String
    
    frmMain.Enabled = True
    path = Me.File1.path
    If Right(path, 1) <> "\" Then
        path = path & "\"
    End If
    
    frmImage.Source.Picture = LoadPicture(path & Me.File1.List(Me.File1.ListIndex))
    frmImage.Dest.Picture = frmImage.Source.Picture
    frmImage.Buffer.Picture = frmImage.Source.Picture
    frmImage.Show
    frmImage.Width = frmImage.Source.Width + (frmImage.Width - frmImage.ScaleWidth)
    frmImage.Height = frmImage.Source.Height + (frmImage.Height - frmImage.ScaleHeight)

    Me.Hide
    frmMain.Enabled = True
    frmImage.Show
    
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
    If Me.File1.ListIndex > -1 Then
        Me.Command1.Enabled = True
    Else
        Me.Command1.Enabled = False
    End If
End Sub

Private Sub File1_DblClick()
    Dim path As String
    frmMain.Enabled = True
    path = Me.File1.path
    If Right(path, 1) <> "\" Then
        path = path & "\"
    End If
    
    frmImage.Source.Picture = LoadPicture(path & Me.File1.List(Me.File1.ListIndex))
    frmImage.Dest.Picture = frmImage.Source.Picture
    frmImage.Buffer.Picture = frmImage.Source.Picture
    
    frmImage.Width = frmImage.Source.Width + (frmImage.Width - frmImage.ScaleWidth)
    frmImage.Height = frmImage.Source.Height + (frmImage.Height - frmImage.ScaleHeight)
    
    Me.Hide
    frmMain.Enabled = True
    frmImage.Show
    
End Sub

Private Sub Form_Load()
    Me.File1.Pattern = "*.bmp;*.jpg;*.jpeg"
    
End Sub
