VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFileMagic 
   Caption         =   "File Magic"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4845
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFileMagic.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   4080
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DNSBrowser.isButton cmdBrowse 
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Icon            =   "frmFileMagic.frx":038A
      Style           =   9
      Caption         =   "Browse"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtFile 
      Height          =   360
      Left            =   240
      TabIndex        =   1
      Text            =   "C:\"
      Top             =   840
      Width           =   3255
   End
   Begin DNSBrowser.isButton cmdHide 
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Icon            =   "frmFileMagic.frx":03A6
      Style           =   9
      Caption         =   "Hide File"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   5767216
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DNSBrowser.isButton cmdOpenFile 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Icon            =   "frmFileMagic.frx":03C2
      Style           =   9
      Caption         =   "Reveal File"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   5767216
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DNSBrowser.isButton cmdOpen 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Icon            =   "frmFileMagic.frx":03DE
      Style           =   9
      Caption         =   "Open"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   5767216
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Make File Hidden"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Shape shpHide 
      Height          =   1335
      Left            =   120
      Top             =   480
      Width           =   4575
   End
End
Attribute VB_Name = "frmFileMagic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBrowse_Click()
cDlg.ShowOpen

txtFile.Text = cDlg.FileName
End Sub

Private Sub cmdHide_Click()
Dim strPath() As String
strPath = Split(cDlg.FileName, "\", -1, 1)

Shell "cmd /c type " & vbQuote & cDlg.FileName & vbQuote & " > DATA\filemagic.dat:" & strPath(UBound(strPath)), vbHide
'MsgBox "cmd /c type " & vbQuote & cDlg.FileName & vbQuote & " > " & vbQuote & Left(cDlg.FileName, Len(cDlg.FileName) - Len(strPath(UBound(strPath)))) & ":" & strPath(UBound(strPath)) & vbQuote
End Sub

Private Sub cmdOpen_Click()
Dim strPath() As String
strPath = Split(txtFile.Text, "\", -1, 1)
Shell "notepad " & App.Path & "\" & "DATA\filemagic.dat" & ":" & strPath(UBound(strPath)), vbNormalFocus
End Sub

Private Sub cmdOpenFile_Click()
Dim strPath() As String
strPath = Split(txtFile.Text, "\", -1, 1)
FileCopy "DATA\filemagic.dat" & ":" & strPath(UBound(strPath)), txtFile.Text
End Sub

Private Sub Command1_Click()
Kill txtFile.Text
End Sub

