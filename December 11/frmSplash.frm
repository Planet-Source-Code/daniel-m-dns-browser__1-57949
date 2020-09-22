VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4245
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7200
      Begin VB.CheckBox chkStartup 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Don't display on startup"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4800
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   5
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Timer tmr_wait 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   0
         Top             =   120
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (C) 2004"
         Height          =   255
         Left            =   5400
         TabIndex        =   3
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblCompany 
         BackStyle       =   0  'Transparent
         Caption         =   "PTK Productions"
         Height          =   255
         Left            =   5400
         TabIndex        =   2
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Initializing....."
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   3720
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "v0.91"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         TabIndex        =   4
         Top             =   240
         Width           =   600
      End
      Begin VB.Image imgLogo 
         Height          =   3585
         Left            =   120
         Picture         =   "frmSplash.frx":000C
         Top             =   360
         Width           =   6750
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Dim wbTextSize As String

Private Sub chkStartup_Click()
WriteString "Miscellaneous", "STARTUP", chkStartup.Value, "DATA\options.dat"
End Sub

Private Sub Form_Initialize()
Dim x As Long
x = InitCommonControls
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
'original interval time was .5 seconds... now .3, will set option for virtually no load time!
    'SetTrial
'SET ONLY WHEN GIVING OUT
'CheckEncryptionKey 'check encryption key!

wbTextSize = ReadString("Miscellaneous", "TEXTSIZE", "DATA\options.dat")
StartVal = ReadString("Miscellaneous", "STARTUP", "DATA\options.dat")


If StartVal <> "0" Then
    tmr_wait.Enabled = False
    frmBrowser.Show
    frmBrowser.mnuTextSizeX.Item(wbTextSize).Checked = True
    Me.Hide
Else
    tmr_wait.Enabled = True
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End If

End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub tmr_wait_Timer()
frmBrowser.Show
frmBrowser.mnuTextSizeX.Item(wbTextSize).Checked = True
tmr_wait.Enabled = False
End Sub
