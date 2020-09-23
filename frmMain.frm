VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wave"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   3270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   2400
      TabIndex        =   11
      Top             =   420
      Width           =   855
   End
   Begin VB.CommandButton CmdRender 
      Caption         =   "Load"
      Height          =   315
      Left            =   2400
      TabIndex        =   10
      Top             =   60
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Horisontal"
      Height          =   1515
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   2355
      Begin VB.HScrollBar HEnergy 
         Height          =   255
         LargeChange     =   20
         Left            =   120
         Max             =   100
         TabIndex        =   9
         Top             =   1080
         Value           =   50
         Width           =   2055
      End
      Begin VB.HScrollBar HWave 
         Height          =   255
         LargeChange     =   20
         Left            =   120
         Max             =   100
         Min             =   1
         TabIndex        =   8
         Top             =   480
         Value           =   1
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Waves:(1)"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label3 
         Caption         =   "Energy: (0)"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vertical"
      Height          =   1515
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2355
      Begin VB.HScrollBar VEnergy 
         Height          =   255
         LargeChange     =   20
         Left            =   120
         Max             =   100
         TabIndex        =   3
         Top             =   1080
         Value           =   50
         Width           =   2055
      End
      Begin VB.HScrollBar VWave 
         Height          =   255
         LargeChange     =   20
         Left            =   120
         Max             =   100
         Min             =   1
         TabIndex        =   1
         Top             =   480
         Value           =   1
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Energy: (0)"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   2115
      End
      Begin VB.Label Label1 
         Caption         =   "Waves: (1)"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2115
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Sub CmdRender_Click()
    Me.Enabled = False

    frmLoad.Show
End Sub



Private Sub RenderWave(XW As Double, YW As Double, XE As Double, YE As Double)
    Const pi As Double = 3.1415
    Dim i As Integer
    Dim u As Integer
    Dim a As Double
    
    'reset, the images..
    frmImage.Show
    frmImage.Buffer.Cls
    frmImage.Dest.Picture = frmImage.Source.Picture
    
    'lets render the vertical first... =)
    For u = 0 To XW
        For i = 0 To frmImage.Source.ScaleHeight / (XW + 1)
            BitBlt frmImage.Buffer.hDC, (Cos(a) * XE), i + ((frmImage.Source.ScaleHeight / (XW + 1)) * u), frmImage.Source.ScaleWidth, 1, frmImage.Source.hDC, 0, i + ((frmImage.Source.ScaleHeight / (XW + 1)) * u), vbSrcCopy
            BitBlt frmImage.Buffer.hDC, (Cos(a) * XE) + frmImage.Source.ScaleWidth, i + ((frmImage.Source.ScaleHeight / (XW + 1)) * u), frmImage.Source.ScaleWidth, 1, frmImage.Source.hDC, 0, i + ((frmImage.Source.ScaleHeight / (XW + 1)) * u), vbSrcCopy
            BitBlt frmImage.Buffer.hDC, (Cos(a) * XE) - frmImage.Source.ScaleWidth, i + ((frmImage.Source.ScaleHeight / (XW + 1)) * u), frmImage.Source.ScaleWidth, 1, frmImage.Source.hDC, 0, i + ((frmImage.Source.ScaleHeight / (XW + 1)) * u), vbSrcCopy
    
            a = i / frmImage.Source.ScaleHeight * (pi * 2 * (XW + 1))
        Next i
    Next u
    'lets render the horisontal next... =)
    a = 0
    For u = 0 To YW
        For i = 0 To frmImage.Source.ScaleWidth / (YW + 1)
            BitBlt frmImage.Dest.hDC, i + ((frmImage.Source.ScaleWidth / (YW + 1)) * u), (Sin(a) * YE), 1, frmImage.Source.ScaleHeight, frmImage.Buffer.hDC, i + ((frmImage.Source.ScaleWidth / (YW + 1)) * u), 0, vbSrcCopy
            BitBlt frmImage.Dest.hDC, i + ((frmImage.Source.ScaleWidth / (YW + 1)) * u), (Sin(a) * YE) + frmImage.Source.ScaleHeight, 1, frmImage.Source.ScaleHeight, frmImage.Buffer.hDC, i + ((frmImage.Source.ScaleWidth / (YW + 1)) * u), 0, vbSrcCopy
            BitBlt frmImage.Dest.hDC, i + ((frmImage.Source.ScaleWidth / (YW + 1)) * u), (Sin(a) * YE) - frmImage.Source.ScaleHeight, 1, frmImage.Source.ScaleHeight, frmImage.Buffer.hDC, i + ((frmImage.Source.ScaleWidth / (YW + 1)) * u), 0, vbSrcCopy
            
            a = i / frmImage.Source.ScaleWidth * (pi * 2 * (YW + 1))
        Next i
    Next u
    frmImage.Dest.Refresh
End Sub



Private Sub CmdSave_Click()
    Me.Enabled = False
    frmSave.Show
    
End Sub

Private Sub Form_Load()
    frmImage.Show
    
End Sub

Private Sub HEnergy_Change()
    RenderWave Me.VWave.Value, Me.HWave.Value, 50 - Me.VEnergy.Value, 50 - Me.HEnergy.Value
    Me.Label3.Caption = "Energy: (" & Me.HEnergy.Value - 50 & ")"
End Sub

Private Sub HWave_Change()
    RenderWave Me.VWave.Value, Me.HWave.Value, 50 - Me.VEnergy.Value, 50 - Me.HEnergy.Value
    Me.Label4.Caption = "Waves: (" & Me.HWave.Value & ")"
End Sub

Private Sub VEnergy_Change()
    RenderWave Me.VWave.Value, Me.HWave.Value, 50 - Me.VEnergy.Value, 50 - Me.HEnergy.Value
    Me.Label2.Caption = "Energy: (" & Me.VEnergy.Value - 50 & ")"
End Sub

Private Sub VWave_Change()
    RenderWave Me.VWave.Value, Me.HWave.Value, 50 - Me.VEnergy.Value, 50 - Me.HEnergy.Value
    Me.Label1.Caption = "Waves: (" & Me.VWave.Value & ")"
End Sub
