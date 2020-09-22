VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFTP 
   Caption         =   "FTP Client"
   ClientHeight    =   9285
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11340
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFTP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkPassive 
      Appearance      =   0  'Flat
      Caption         =   "Passive &Mode"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3720
      TabIndex        =   45
      Top             =   2760
      Width           =   1815
   End
   Begin DNSBrowser.isButton cmdMkDir 
      Height          =   615
      Left            =   3600
      TabIndex        =   42
      Top             =   0
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      Icon            =   "frmFTP.frx":038A
      Style           =   9
      Caption         =   "MKDIR"
      IconAlign       =   3
      CaptionAlign    =   4
      iNonThemeStyle  =   0
      Tooltiptitle    =   "Send File"
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
   Begin VB.PictureBox picMode 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3720
      ScaleHeight     =   855
      ScaleWidth      =   1695
      TabIndex        =   39
      Top             =   1800
      Width           =   1695
      Begin VB.OptionButton OptMode 
         Caption         =   "Auto Detect"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   44
         Top             =   120
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptMode 
         Caption         =   "Binary Mode"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   41
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton OptMode 
         Caption         =   "ASCII Mode"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   40
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.ListBox lst_cmdhist 
      Height          =   300
      Left            =   6120
      TabIndex        =   38
      Top             =   9240
      Visible         =   0   'False
      Width           =   3255
   End
   Begin DNSBrowser.XP_ProgressBar pgDL 
      Height          =   255
      Left            =   3600
      TabIndex        =   37
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   10513481
   End
   Begin DNSBrowser.isButton cmdRename 
      Height          =   615
      Left            =   3600
      TabIndex        =   36
      Top             =   3000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      Icon            =   "frmFTP.frx":0924
      Style           =   9
      Caption         =   "Rename"
      IconAlign       =   3
      CaptionAlign    =   4
      iNonThemeStyle  =   0
      Tooltiptitle    =   "Send File"
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
   Begin MSComDlg.CommonDialog Cdlg 
      Left            =   11040
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   8760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":11FE
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":1AD8
            Key             =   "Unknown"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":23B2
            Key             =   "Archive"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":308C
            Key             =   "Text"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":3966
            Key             =   "JPEG"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":4240
            Key             =   "BMP"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":4B1A
            Key             =   "DLL"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":53F4
            Key             =   "SYS"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":5CCE
            Key             =   "GIF"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":65A8
            Key             =   "CAB"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":6E82
            Key             =   "MPEG"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":775C
            Key             =   "MID"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":8036
            Key             =   "AVI"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":8910
            Key             =   "TTF"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":91EA
            Key             =   "WWW"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":9AC4
            Key             =   "Font"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":A39E
            Key             =   "HLP"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picKey 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   9240
      ScaleHeight     =   2025
      ScaleWidth      =   2025
      TabIndex        =   30
      Top             =   6765
      Width           =   2055
      Begin VB.Label lblAppMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Application Message"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Shape shpGrey 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   120
         Top             =   1380
         Width           =   135
      End
      Begin VB.Label lblKey 
         BackStyle       =   0  'Transparent
         Caption         =   "FTP Console Key ::"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblBadResponse 
         BackStyle       =   0  'Transparent
         Caption         =   "Server Bad Response"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Shape shpRed 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   120
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label lblServerResponse 
         BackStyle       =   0  'Transparent
         Caption         =   "Server Response"
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   720
         Width           =   1335
      End
      Begin VB.Shape shpGreen 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   120
         Top             =   780
         Width           =   135
      End
      Begin VB.Label lblUserCommand 
         BackStyle       =   0  'Transparent
         Caption         =   "User Command"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   420
         Width           =   1095
      End
      Begin VB.Shape shpBlue 
         BackColor       =   &H00A06C49&
         BackStyle       =   1  'Opaque
         Height          =   135
         Left            =   120
         Top             =   480
         Width           =   135
      End
   End
   Begin RichTextLib.RichTextBox txtConsole 
      Height          =   2055
      Left            =   0
      TabIndex        =   29
      Top             =   6765
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   3625
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmFTP.frx":AC78
   End
   Begin MSComctlLib.ImageList imglstIcons 
      Left            =   4800
      Top             =   8760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":AD5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":B0F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":BDD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":C36B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":C905
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":CE9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":D439
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":D9D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":DF6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":E507
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":EAA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":F03B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":F3D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":F96F
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":FF09
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":104A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":10A3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":10FD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":11571
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":11B0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":120A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFTP.frx":1263F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckData 
      Left            =   3840
      Top             =   8760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1601
   End
   Begin MSWinsockLib.Winsock sckCon 
      Left            =   4320
      Top             =   8760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   360
      Left            =   10800
      TabIndex        =   28
      Top             =   8880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtSend 
      Height          =   360
      Left            =   3360
      TabIndex        =   27
      Top             =   8880
      Width           =   7935
   End
   Begin DNSBrowser.isButton cmdDelete 
      Height          =   615
      Left            =   4560
      TabIndex        =   26
      Top             =   3600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      Icon            =   "frmFTP.frx":12BD9
      Style           =   9
      Caption         =   " Delete"
      IconAlign       =   3
      CaptionAlign    =   4
      iNonThemeStyle  =   0
      Tooltiptitle    =   "Send File"
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
   Begin DNSBrowser.isButton cmdCancel 
      Height          =   615
      Left            =   3600
      TabIndex        =   25
      Top             =   3600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      Icon            =   "frmFTP.frx":12F73
      Style           =   9
      Caption         =   " Cancel"
      IconAlign       =   3
      CaptionAlign    =   4
      iNonThemeStyle  =   0
      Tooltiptitle    =   "Send File"
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
   Begin VB.DirListBox dlstLocal 
      Appearance      =   0  'Flat
      Height          =   3870
      Left            =   0
      TabIndex        =   24
      Top             =   360
      Width           =   3495
   End
   Begin VB.FileListBox flstLocal 
      Height          =   2490
      Left            =   0
      TabIndex        =   23
      Top             =   4245
      Width           =   5535
   End
   Begin DNSBrowser.isButton cmdGET 
      Height          =   615
      Left            =   3600
      TabIndex        =   22
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      Icon            =   "frmFTP.frx":1330D
      Style           =   9
      Caption         =   "Download"
      IconAlign       =   3
      CaptionAlign    =   4
      iNonThemeStyle  =   0
      Tooltiptitle    =   "Send File"
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
   Begin DNSBrowser.isButton cmdCDUP 
      Height          =   615
      Left            =   4560
      TabIndex        =   21
      Top             =   0
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      Icon            =   "frmFTP.frx":136A7
      Style           =   9
      Caption         =   "CDUP"
      IconAlign       =   3
      CaptionAlign    =   4
      iNonThemeStyle  =   0
      Tooltiptitle    =   "Send File"
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
   Begin DNSBrowser.isButton cmdSendF 
      Height          =   615
      Left            =   4560
      TabIndex        =   20
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1085
      Icon            =   "frmFTP.frx":13A41
      Style           =   9
      Caption         =   "Upload"
      IconAlign       =   3
      CaptionAlign    =   4
      iNonThemeStyle  =   0
      Tooltiptitle    =   "Send File"
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
   Begin VB.DriveListBox drvLocal 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   3495
   End
   Begin VB.PictureBox picHeader 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7680
      ScaleHeight     =   225
      ScaleWidth      =   1785
      TabIndex        =   4
      Top             =   4440
      Width           =   1815
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FTP Connection Setup"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.PictureBox picConnect 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2070
      Left            =   6120
      ScaleHeight     =   2040
      ScaleWidth      =   4785
      TabIndex        =   3
      Top             =   4560
      Width           =   4815
      Begin VB.TextBox txtTimeOut 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A06C49&
         Height          =   360
         Left            =   960
         TabIndex        =   18
         Text            =   "60"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A06C49&
         Height          =   345
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   16
         Text            =   "21"
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox chkAnonymous 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Anonymous"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3360
         TabIndex        =   14
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtPass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "•"
         TabIndex        =   11
         Text            =   "Password"
         Top             =   1120
         Width           =   2295
      End
      Begin VB.TextBox txtUser 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   960
         TabIndex        =   9
         Text            =   "Anonymous"
         Top             =   680
         Width           =   2295
      End
      Begin VB.TextBox txtServer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   960
         TabIndex        =   7
         Text            =   "207.46.133.140"
         Top             =   240
         Width           =   2295
      End
      Begin DNSBrowser.isButton cmdClear 
         Height          =   345
         Left            =   3360
         TabIndex        =   12
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         Icon            =   "frmFTP.frx":141BB
         Style           =   9
         Caption         =   "Clear"
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
      Begin DNSBrowser.isButton cmdConnect 
         Height          =   345
         Left            =   3360
         TabIndex        =   13
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         Icon            =   "frmFTP.frx":141D7
         Style           =   9
         Caption         =   "Connect"
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
      Begin VB.Label lblTimeOut 
         BackStyle       =   0  'Transparent
         Caption         =   "Time Out:"
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   1605
         Width           =   735
      End
      Begin VB.Label lblPort 
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   255
         Left            =   2040
         TabIndex        =   15
         Top             =   1605
         Width           =   375
      End
      Begin VB.Label lblPass 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblServer 
         BackStyle       =   0  'Transparent
         Caption         =   "FTP Server:"
         Height          =   255
         Left            =   60
         TabIndex        =   6
         Top             =   270
         Width           =   975
      End
   End
   Begin DNSBrowser.isButton cmdSave 
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   8880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      Icon            =   "frmFTP.frx":141F3
      Style           =   9
      Caption         =   "Save Info..."
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
   Begin DNSBrowser.isButton cmdClrConsole 
      Height          =   345
      Left            =   1800
      TabIndex        =   0
      Top             =   8880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      Icon            =   "frmFTP.frx":1420F
      Style           =   9
      Caption         =   "Clear Console"
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
   Begin MSComctlLib.ListView lvFTP 
      Height          =   4215
      Left            =   5640
      TabIndex        =   2
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   7435
      View            =   3
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "imglstIcons"
      SmallIcons      =   "imglstIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "File Name"
         Text            =   "File Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "File Size"
         Text            =   "Size (bytes)"
         Object.Width           =   1765
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Last Modified"
         Text            =   "Last Modified"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Permissions"
         Text            =   "Permissions"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Owner"
         Text            =   "Owner"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Group"
         Text            =   "Group"
         Object.Width           =   1588
      EndProperty
   End
   Begin VB.Label lblProgress 
      Alignment       =   1  'Right Justify
      Caption         =   "0/0 bytes"
      Height          =   255
      Left            =   3600
      TabIndex        =   43
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuRemote 
      Caption         =   "Options"
      Begin VB.Menu mnuDownload 
         Caption         =   "Download"
      End
      Begin VB.Menu mnuOptionsSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
      End
      Begin VB.Menu mnuRemoteSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArrange 
         Caption         =   "Auto-Arrange"
         Begin VB.Menu mnuArrangeType 
            Caption         =   "None"
            Index           =   0
         End
         Begin VB.Menu mnuArrangeType 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuArrangeType 
            Caption         =   "Left"
            Index           =   2
         End
         Begin VB.Menu mnuArrangeType 
            Caption         =   "Top"
            Checked         =   -1  'True
            Index           =   3
         End
      End
      Begin VB.Menu mnuView 
         Caption         =   "View"
         Begin VB.Menu mnuViewType 
            Caption         =   "Icon"
            Index           =   0
         End
         Begin VB.Menu mnuViewType 
            Caption         =   "List"
            Index           =   1
         End
         Begin VB.Menu mnuViewType 
            Caption         =   "Details"
            Checked         =   -1  'True
            Index           =   2
         End
      End
   End
End
Attribute VB_Name = "frmFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declarations
Dim strLocalIP As String, strUser As String, strPass As String, strRemoteServer As String
Dim cPort As Long, strData As String, lastUsedCommand() As String, strDir$, strSaveLoc$
Dim properCommand As String, lngDLSize As Long, lngDLComplete As Long
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Dim bGetDir As Boolean, hist As Long, bCancelRename As Boolean, OldString As String
Dim i As Long, j As Long
Dim fso

Private Sub chkAnonymous_Click()
If chkAnonymous.Value = 1 Then
    txtUser.Text = "Anonymous"
    txtPass.Text = "user@site.com"
End If
End Sub

Private Sub cmdCancel_Click()
ReDim Preserve lastUsedCommand(1)
lastUsedCommand(0) = "ABOR"
sckCon.SendData "ABOR" & vbCrLf
End Sub

Private Sub cmdCDUP_Click()
    sckCon.SendData "PORT " & strLocalIP & ",6,65" & vbCrLf
    Pause (500) 'give time for response
    ReDim Preserve lastUsedCommand(1)
    lastUsedCommand(0) = "CDUP"
    sckCon.SendData "CDUP" & vbCrLf
    ConsoleUpdate "CDUP", 1
    Pause (500)
    sckCon.SendData "LIST" & vbCrLf
End Sub

Private Sub cmdClear_Click()
    txtServer.Text = "ftp."
    txtUser.Text = vbNullString
    txtPass.Text = vbNullString
    txtPort.Text = "21"
    txtTimeOut.Text = "80"
End Sub

Private Sub cmdClrConsole_Click()
txtConsole.Text = "============================================ DNS Browser - FTP Console ================================"
End Sub

Private Sub cmdConnect_Click()
If cmdConnect.Caption = "Connect" Then
    cmdConnect.Caption = "Disconnect"
    strUser$ = txtUser.Text
    strPass$ = txtPass.Text
    strRemoteServer$ = txtServer.Text
    cPort = txtPort.Text
    sckData.Close
    sckData.Listen
    
    sckCon.Close
    sckCon.Connect strRemoteServer$, cPort
    ConsoleUpdate "Connecting to " & sckCon.RemoteHost & ":" & sckCon.RemotePort & "...", 2
Else
    cmdConnect.Caption = "Connect"
    ConsoleUpdate "Connection to server has been closed.", 2
    lvFTP.ListItems.Clear
    sckData.Close
    sckCon.Close
End If
End Sub


Private Function getSckLocalIP()
sckData.Close
strLocalIP = sckData.LocalIP

    Do Until InStr(sckLocalIP, ".") = 0
        strLocalIP = Left(strLocalIP, InStr(strLocalIP, ".") - 1) & "," & Right(strLocalIP, Len(strLocalIP) - InStr(strLocalIP, "."))
    Loop
    

End Function

Private Sub cmdDebug_Click()
lvFTP.View = lvwIcon
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Delete this file? " & vbNewLine & "'" & lvFTP.SelectedItem.Text & "'", vbYesNo, "Delete File?") = vbYes Then
Select Case lvFTP.SelectedItem.Icon
    Case 1
        ConsoleUpdate "RMDIR " & lvFTP.SelectedItem.Text, 1
        sckCon.SendData "RMD " & lvFTP.SelectedItem.Text & vbCrLf
    
    Case Else
        sckCon.SendData "PORT " & strLocalIP & ",6,65" & vbCrLf
        Pause (500)
        ConsoleUpdate "DELETE " & lvFTP.SelectedItem.Text, 1
        sckCon.SendData "DELE " & lvFTP.SelectedItem.Text & vbCrLf
End Select
Else
ConsoleUpdate "Operation canceled.", 2
End If
End Sub

Private Sub cmdGET_Click()
Cdlg.FileName = lvFTP.SelectedItem.Text
Cdlg.ShowSave

    If Cdlg.FileName <> lvFTP.SelectedItem.Text Then
        strSaveLoc$ = Cdlg.FileName
        
        If fso.FileExists(strSaveLoc$) <> True Then
            fso.CreateTextFile (strSaveLoc$)
            Else
                If MsgBox("Are you sure you want to overwrite this file?", vbCritical + vbYesNo, "Overwrite File?") = vbYes Then
                    fso.DeleteFile (strSaveLoc$)
                    fso.CreateTextFile (strSaveLoc$)
                Else
                    ConsoleUpdate "Download Canceled", 2
                    Exit Sub
                End If
        End If
        sckCon.SendData "PORT " & strLocalIP & ",6,65" & vbCrLf
        Pause (500) 'give time for response
            If OptMode(2).Value = True Then
                If lvFTP.SelectedItem.Icon <> 3 Then
                    ConsoleUpdate "TYPE I - BINARY Mode set", 1
                    sckCon.SendData "TYPE I" & vbCrLf
                    Pause (500)
                Else
                    ConsoleUpdate "TYPE A - ASCII Mode set", 1
                    sckCon.SendData "TYPE A" & vbCrLf
                    Pause (500)
                End If
            Else
                If OptMode(0).Value = True Then
                    ConsoleUpdate "TYPE A - ASCII Mode set", 1
                    sckCon.SendData "TYPE A" & vbCrLf
                    Pause (500)
                Else
                    ConsoleUpdate "TYPE I - BINARY Mode set", 1
                    sckCon.SendData "TYPE I" & vbCrLf
                    Pause (500)
                End If
                
            End If
        ReDim lastUsedCommand(1)
        lngDLSize = lvFTP.SelectedItem.SubItems(1)
        pgDL.Max = lngDLSize
        lastUsedCommand(0) = "RETR"
        lastUsedCommand(1) = lvFTP.SelectedItem.Text
        properCommand$ = LCase(lastUsedCommand(0))
        sckCon.SendData "RETR " & lvFTP.SelectedItem.Text & vbCrLf
        ConsoleUpdate "RETR " & lvFTP.SelectedItem.Text, 1
    End If
End Sub

Private Sub cmdMkDir_Click()
Dim strDirName As String

strDirName$ = InputBox("What would you like the folder name to be?", "Create Folder", "New Folder")

If strDirName$ <> vbNullString Then
    ReDim lastUsedCommand(1)
    lastUsedCommand(0) = "MKDIR"
    lastUsedCommand(1) = strDirName$
    ConsoleUpdate "MKD " & """" & strDirName$ & """", 1
    sckCon.SendData "MKD " & strDirName$ & vbCrLf
Else
    ConsoleUpdate "Action canceled.", 1
End If
End Sub

Private Sub cmdRename_Click()
lvFTP.SelectedItem.Selected = True
lvFTP.StartLabelEdit
End Sub

Private Sub cmdSave_Click()
Cdlg.FileName = "FTPDATA.txt"
Cdlg.ShowSave

    If Cdlg.FileName <> "FTPDATA.txt" Then
        Open Cdlg.FileName For Append As #1
            Print #1, txtConsole.Text
        Close #1
    End If
    
End Sub

Private Sub cmdSend_Click()

'============================================================================='
'just some basic support for commands such as ABORT which ends the connection '
'or retrieve. the rest just sends the command as the user types in            '
'============================================================================='

lastUsedCommand = Split(txtSend.Text, " ", -1, 1)
If lastUsedCommand(0) = "ABOR" Then
    sckCon.SendData "ABOR" & vbCrLf
Exit Sub
End If

If lastUsedCommand(0) = "RETR" Then
    strSaveLoc$ = flstLocal.Path & "\" & lvFTP.SelectedItem.Text
    properCommand$ = LCase(lastUsedCommand(0))
    sckCon.SendData "TYPE I" & vbCrLf
    Pause (500)
End If
sckCon.SendData "PORT " & strLocalIP & ",6,65" & vbCrLf
Pause (500) 'give time for response
sckCon.SendData txtSend.Text & vbCrLf
txtSend.Text = vbNullString
End Sub

Private Sub cmdSendF_Click()
    sckCon.SendData "PORT " & strLocalIP & ",6,65" & vbCrLf
    Pause (500) 'give time for response
    ReDim lastUsedCommand(1)
    lastUsedCommand(0) = "STOR"
    lastUsedCommand(1) = lvFTP.SelectedItem.Text
    sckCon.SendData "STOR " & flstLocal.FileName & " " & vbCrLf
    ConsoleUpdate "STOR " & flstLocal.Path & "\" & flstLocal.FileName, 1
End Sub

Private Sub dlstLocal_Change()
flstLocal.Path = dlstLocal.List(dlstLocal.ListIndex)
End Sub

Private Sub drvLocal_Change()
On Error GoTo fixMe:
dlstLocal.Path = Left(drvLocal.Drive, 1) & ":\"
flstLocal.Path = Left(drvLocal.Drive, 1) & ":\"
Exit Sub
fixMe:
drvLocal.Drive = "C:"
dlstLocal.Path = "C:\"
flstLocal.Path = "C:\"

End Sub

Private Sub Form_Load()
ReDim lastUsedCommand(1)
Set fso = CreateObject("Scripting.FileSystemObject")
bGetDir = False
txtConsole.SelStart = 1
txtConsole.SelLength = Len(txtConsole.Text)
txtConsole.SelColor = &H808080
txtConsole.SelText = vbNewLine
strLocalIP$ = Replace(sckData.LocalIP, ".", ",")
bDiscard = False
bCancelRename = False
End Sub
Private Sub lvFTP_AfterLabelEdit(Cancel As Integer, NewString As String)

'====================================================================='
'Label editing will send the RNFR and RNTO (rename from/to) command   '
'====================================================================='

ReDim lastUsedCommand(1)
lastUsedCommand(0) = "RNTO" ' works like a charm... rename file from
lastUsedCommand(1) = lvFTP.SelectedItem.Text
ConsoleUpdate "REN " & lvFTP.SelectedItem.Text & " to " & NewString, 1
sckCon.SendData "RNFR " & lvFTP.SelectedItem.Text & vbCrLf 'from
Pause (500)

If bCancelRename = True Then 'check access rights
    Cancel = 1
    NewString = OldString
    bCancelRename = False
    Exit Sub
End If
Pause (100) 'if accessable then rename to
sckCon.SendData "RNTO " & NewString

End Sub


Private Sub lvFTP_BeforeLabelEdit(Cancel As Integer)
OldString = lvFTP.SelectedItem.Text
End Sub


Private Sub lvFTP_DblClick()
'==========================Description========================'
'lets user navigate through clicking or download a file       '
'if it is a TXT, otherwise must click the download button     '
'============================================================='
    Select Case lvFTP.SelectedItem.Icon
    
        Case 1
            sckCon.SendData "PORT " & strLocalIP & ",6,65" & vbCrLf
            Pause (500) 'give time for response
            ReDim lastUsedCommand(1)
            lastUsedCommand(0) = "CWD"
            lastUsedCommand(1) = lvFTP.SelectedItem.Text
            sckCon.SendData "CWD " & lvFTP.SelectedItem.Text & vbCrLf
            ConsoleUpdate "CWD", 1
            Pause (500)
            lastUsedCommand(0) = "LIST"
            sckCon.SendData "LIST" & vbCrLf
    
        Case 22
            Cdlg.FileName = lvFTP.SelectedItem.Text
            Cdlg.ShowSave
            If Cdlg.FileName <> lvFTP.SelectedItem.Text Then
                strSaveLoc$ = Cdlg.FileName
                sckCon.SendData "PORT " & strLocalIP & ",6,65" & vbCrLf
                Pause (500) 'give time for response
                ReDim lastUsedCommand(1)
                lastUsedCommand(0) = "RETR ASCII"
                lngDLSize = lvFTP.SelectedItem.SubItems(1)
                pgDL.Max = lngDLSize
                lastUsedCommand(1) = lvFTP.SelectedItem.Text
                properCommand$ = LCase(lastUsedCommand(0))
                sckCon.SendData "RETR " & lvFTP.SelectedItem.Text & vbCrLf
                ConsoleUpdate "RETR ASCII " & lvFTP.SelectedItem.Text, 1
            Else
                Exit Sub
            End If
    End Select
    
End Sub


Private Sub lvFTP_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Call mnuRemote_Click
    PopupMenu mnuRemote
End If
End Sub


Private Sub mnuArrangeName_Click()
lvFTP.Arrange = lvwAutoTop
End Sub

Private Sub mnuArrangeType_Click(Index As Integer)
'===============================================
'Allows user to change the arrangement of icons
'===============================================
Select Case Index
    Case 0
        lvFTP.Arrange = lvwNone
        mnuArrangeType.Item(Index).Checked = True
        mnuArrangeType.Item(2).Checked = False
        mnuArrangeType.Item(3).Checked = False
    Case 2
        lvFTP.Arrange = lvwAutoLeft
        mnuArrangeType.Item(Index).Checked = True
        mnuArrangeType.Item(0).Checked = False
        mnuArrangeType.Item(3).Checked = False
    Case 3
        lvFTP.Arrange = lvwAutoTop
        mnuArrangeType.Item(Index).Checked = True
        mnuArrangeType.Item(0).Checked = False
        mnuArrangeType.Item(2).Checked = False
End Select

End Sub

Private Sub mnuDelete_Click()
Call cmdDelete_Click
End Sub

Private Sub mnuRemote_Click()
If lvFTP.ListItems.Count <> 0 Then
    mnuDelete.Enabled = True
    mnuRename.Enabled = True
    If lvFTP.SelectedItem.Icon <> 1 Then
        mnuDownload.Enabled = True
    Else
        mnuDownload.Enabled = False
    End If
Else
    mnuDownload.Enabled = False
    mnuDelete.Enabled = False
    mnuRename.Enabled = False
End If
End Sub
Private Sub mnuDownload_Click()
Call cmdGET_Click
End Sub

Private Sub mnuFileClose_Click()
Unload Me
End Sub

Private Sub mnuRename_Click()
lvFTP.SelectedItem.Selected = True
lvFTP.StartLabelEdit
End Sub

Private Sub mnuViewType_Click(Index As Integer)
'===========================================
'Allows user to change view style
'===========================================
Select Case Index
    Case 0
        lvFTP.View = lvwIcon
        mnuViewType.Item(Index).Checked = True
        mnuViewType.Item(1).Checked = False
        mnuViewType.Item(2).Checked = False
    Case 1
        lvFTP.View = lvwList
        mnuViewType.Item(Index).Checked = True
        mnuViewType.Item(0).Checked = False
        mnuViewType.Item(2).Checked = False
    Case 2
        lvFTP.View = lvwReport
        mnuViewType.Item(Index).Checked = True
        mnuViewType.Item(0).Checked = False
        mnuViewType.Item(1).Checked = False
End Select
End Sub

Private Sub OptMode_Click(Index As Integer)
'===========================================
'Change data receiving type. All files will
'automatically be downloaded as binary though.
'Only ASCII is downloaded when you double click
'on a text file/rtf for viewing
'===========================================
Select Case Index
    Case 0
        ConsoleUpdate "TYPE A - ASCII Mode set", 1
        sckCon.SendData "TYPE A" & vbCrLf
    
    Case 1
        ConsoleUpdate "TYPE I - BINARY Mode Set", 1
        sckCon.SendData "TYPE I" & vbCrLf

    Case 2
        ConsoleUpdate "Auto Detect Mode Set", 1
End Select
        
    
End Sub


Private Sub sckCon_Close()
ConsoleUpdate "Disconnected from " & sckCon.RemoteHost, 2
lvFTP.ListItems.Clear
End Sub

Private Sub sckCon_Connect()
ConsoleUpdate "Connected to " & sckCon.RemoteHost & ":" & sckCon.RemotePort, 2
End Sub

Private Sub sckCon_DataArrival(ByVal bytesTotal As Long)
'===============================================
'Description: Data arrival for connection, basic
'status information is retrieved through here.
'===============================================
sckCon.GetData strData$

'Transfer complete.
'= 1 Or InStr(strData$, "226 ASCII Transfer complete.")

If InStr(strData$, "226") = 1 And bGetDir = True And lastUsedCommand(0) = "LIST" Then 'check if directory
    'Pause (500) 'make sure the data didnt just get sent too fast
    bGetDir = False 'listing is complete
    If Len(strDir$) <> 0 Then
        loadLvDir
    End If
End If

If InStr(strData$, "150 Opening ASCII mode data connection for /bin/ls.") = 1 Or _
InStr(strData$, "150 Opening BINARY mode data connection for /bin/ls.") = 1 Or _
InStr(strData$, "150 Here comes the directory listing.") = 1 Then
        bGetDir = True 'if opening then make directory listing set to true for grabbing info
End If

If InStr(strData$, "550") And InStr(strData$, OldString$) Then
    bCancelRename = True 'if 550 and contains the oldname string then you cant rename
End If


'===================================================
'The below handles status information appropriately
'===================================================
Select Case Left(strData$, 3)
    Case Is <= 399
        txtConsole.SelStart = Len(txtConsole.Text)
        txtConsole.SelColor = &H4000&
        txtConsole.SelText = strData$
        

    Case Is >= 400
        txtConsole.SelStart = Len(txtConsole.Text)
        txtConsole.SelColor = &H80&
        txtConsole.SelText = strData$
        
End Select
'====================================================
'this is for intial connection information
'====================================================
If Left(strData, 3) = "220" Then sckCon.SendData "USER " & strUser$ & vbCrLf
If Left(strData, 3) = "331" Then
   ' Pause (500)
    sckCon.SendData "PASS " & strPass$ & vbCrLf
    Pause (450)
    sckCon.SendData "PWD" & vbCrLf
    ReDim lastUsedCommand(0)
    ConsoleUpdate "PWD", 1
    Pause (450) 'give time for response
    sckCon.SendData "PORT " & strLocalIP & ",6,65" & vbCrLf
    Pause (450) 'give time for response
    ReDim lastUsedCommand(0)
    lastUsedCommand(0) = "LIST"
    sckCon.SendData "LIST" & vbCrLf
    ConsoleUpdate "LIST", 1
End If
End Sub
Sub ConsoleUpdate(strCommand As String, strType As Long)

Select Case strType
    Case 1
        txtConsole.SelStart = Len(txtConsole.Text) & vbNewLine
        txtConsole.SelColor = &H800000
        txtConsole.SelText = strCommand$ & vbNewLine

    Case 2
        txtConsole.SelStart = Len(txtConsole.Text) & vbNewLine
        txtConsole.SelColor = &H404040
        txtConsole.SelText = strCommand$ & vbNewLine
        
    Case 3
        txtConsole.SelStart = Len(txtConsole.Text) & vbNewLine
        txtConsole.SelColor = &H80&
        txtConsole.SelText = strCommand$ & vbNewLine
End Select

End Sub
Sub Pause(dwMil As Integer)
Dim initTime As Long, fTime As Long
initTime = GetTickCount
    Do Until fTime - initTime >= dwMil
        fTime = GetTickCount
        DoEvents
    Loop
End Sub

Private Sub sckCon_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If sckCon.State = sckTimedout Then
    ConsoleUpdate "Request timed out", 3
Else
    ConsoleUpdate Number & " " & Description, 3
End If
sckCon.Close
cmdConnect.Caption = "Connect"
End Sub

Private Sub sckData_Close()
sckData.Close
sckData.Listen
End Sub

Private Sub sckData_ConnectionRequest(ByVal requestID As Long)
If sckData.State <> sckClosed Then
    sckData.Close
    sckData.Accept requestID
End If
End Sub
Private Sub loadLvDir()
'==============================================
'Description: Handles all directory loading
'==============================================
lvFTP.ListItems.Clear
Dim strItems() As String, strFileName$, lngFileSize As Long, strLastModified$, strAccess$
Dim strOwner$, strGroup As String, splitItem() As String, lngIndex As Long, lvitem As ListItem
Dim bTypeTwo As Byte
strItems() = Split(strDir$, vbCrLf, -1, 1)
Dim lngCountRemove As Long
For i = 0 To UBound(strItems) - 1
   lngCountRemove = 0
   If InStr(strItems(i), "dr-") = 1 Then

        splitItem() = Split(strItems(i), " ", -1, 1)
        '===========================================================================
        'NEW ALGORITHM TO RETRIEVE DIRECTORY LISTING... SPLIT UP DATA AND RE-INDEX
        'ARRAY BY REMOVING EMPTY ARRAYS
        '===========================================================================
        For j = 0 To UBound(splitItem) - lngCountRemove 'index remove count must be subtracted
            For k = 0 To UBound(splitItem)
                If splitItem(k) = vbNullString Then 'if the item is empty then
                    If splitItem(UBound(splitItem)) = vbNullString Then 'if if the upperbound is empty
                        Do Until splitItem(UBound(splitItem)) <> vbNullString 'do until it isnt empty
                            ReDim Preserve splitItem(UBound(splitItem) - 1)
                            lngCountRemove = lngCountRemove + 1 'count = count + 1
                        Loop
                        Exit For 'then go onto next k
                    Else 'otherwise
                        splitItem(k) = splitItem(UBound(splitItem)) 'make it the upperbound data
                        ReDim Preserve splitItem(UBound(splitItem) - 1) 're-size array again
                        lngCountRemove = lngCountRemove + 1 'add count
                        Exit For
                    End If
                End If
            Next k
        Next j
        
        strAccess$ = splitItem(0)
        strFileName$ = splitItem(1)
        strOwner$ = splitItem(4)
        strLastModified$ = splitItem(6) & " " & splitItem(5) & " " & splitItem(2)
        strGroup$ = splitItem(8)
        lngFileSize = splitItem(7)
        
        '==================================================================
        'OLD FUNCTION USED TO GET DIRECTORY LISTING... DID NOT WORK WELL
        '==================================================================
        'For j = 0 To UBound(splitItem)
        'bTypeTwo = 0
        '    strAccess$ = splitItem(0)
        '    strOwner$ = splitItem(4)
        '    strGroup$ = splitItem(8)
        '    If splitItem(UBound(splitItem) - 6) <> vbNullString Then
        '        lngFileSize = splitItem(UBound(splitItem) - 6)
        '    Else
        '        lngFileSize = splitItem(UBound(splitItem) - 5)
        '        bTypeTwo = 1
        '    End If
        '
        '    If bTypeTwo <> 1 Then
        ''        strLastModified$ = splitItem(UBound(splitItem) - 5) & " " & splitItem(UBound(splitItem) - 3) & " " & splitItem(UBound(splitItem) - 1)
        '    Else
        '        strLastModified$ = splitItem(UBound(splitItem) - 4) & " " & splitItem(UBound(splitItem) - 3) & " " & splitItem(UBound(splitItem) - 1)
        '    End If
        '    strFileName$ = splitItem(UBound(splitItem))
        'Next j
        
        
        
        
        Set lvitem = lvFTP.ListItems.Add(, , strFileName$, 1, 1)
        lvitem.SubItems(1) = lngFileSize
        lvitem.SubItems(2) = strLastModified$
        lvitem.SubItems(3) = strAccess$
        lvitem.SubItems(4) = strOwner$
        lvitem.SubItems(5) = strGroup$
    End If
Next i

For i = 0 To UBound(strItems) - 1
   lngCountRemove = 0
    If InStr(strItems(i), "-r") = 1 Then
    
        Dim lngIcon As Long, strExt() As String
        splitItem() = Split(strItems(i), " ", -1, 1)
        
        For j = 0 To UBound(splitItem) - lngCountRemove 'index remove count must be subtracted
            For k = 0 To UBound(splitItem)
                If splitItem(k) = vbNullString Then 'if the item is empty then
                    If splitItem(UBound(splitItem)) = vbNullString Then 'if if the upperbound is empty
                        Do Until splitItem(UBound(splitItem)) <> vbNullString 'do until it isnt empty
                            ReDim Preserve splitItem(UBound(splitItem) - 1)
                            lngCountRemove = lngCountRemove + 1 'count = count + 1
                        Loop
                        Exit For 'then go onto next k
                    Else 'otherwise
                        splitItem(k) = splitItem(UBound(splitItem)) 'make it the upperbound data
                        ReDim Preserve splitItem(UBound(splitItem) - 1) 're-size array again
                        lngCountRemove = lngCountRemove + 1 'add count
                        Exit For
                    End If
                End If
            Next k
            
        Next j
            strAccess$ = splitItem(0)
            strFileName$ = splitItem(1)
            strOwner$ = splitItem(4)
            strLastModified$ = splitItem(6) & " " & splitItem(5) & " " & splitItem(2)
            strGroup$ = splitItem(8)
            lngFileSize = splitItem(7)
            
        strExt$ = Split(strFileName$, ".", -1, 1) 'split to get UBOUND for extention
        
        Select Case LCase(strExt$(UBound(strExt))) 'get file extentions to determine icon type
            Case "zip", "rar", "gz"
                lngIcon = 2
            
            Case "rtf"
                lngIcon = 3
            
            Case "htm", "html", "php", "asp", "shtml"
                lngIcon = 4
            
            Case "chm", "hlp"
                lngIcon = 5
            
            Case "mid"
                lngIcon = 6
            
            Case "wav"
                lngIcon = 7
            
            Case "avi", "mov"
                lngIcon = 8
            
            Case "mpeg", "mpg"
                lngIcon = 9
            
            Case "gif"
                lngIcon = 10
            
            Case "jpg", "jpeg"
                lngIcon = 11
            
            Case "bmp"
                lngIcon = 12
            
            Case "dll"
                lngIcon = 13
            
            Case "sys"
                lngIcon = 14
            
            Case "cab"
                lngIcon = 15
            
            Case "ttf"
                lngIcon = 16
            
            Case "cls"
                lngIcon = 17
            
            Case "frm", "fra"
                lngIcon = 18
            
            Case "bas"
                lngIcon = 19
            
            Case "vbp", "vbg"
                lngIcon = 20
            
            Case "txt"
                lngIcon = 22
                
            Case Else 'else make "default file" icon
                lngIcon = 21
        End Select
        
        '================================
        'Set Item Information
        '================================
        Set lvitem = lvFTP.ListItems.Add(, , strFileName$, lngIcon, lngIcon)
        lvitem.SubItems(1) = lngFileSize
        lvitem.SubItems(2) = strLastModified$
        lvitem.SubItems(3) = strAccess$
        lvitem.SubItems(4) = strOwner$
        lvitem.SubItems(5) = strGroup$

    End If
Next i
strDir$ = vbNullString
End Sub
Private Sub sckData_DataArrival(ByVal bytesTotal As Long)

If lastUsedCommand(0) = "ABOR" Then
    Exit Sub
End If
sckData.GetData strData$

If lastUsedCommand(0) = "MKDIR" And InStr(strData$, "550") = 0 Then
'input command to refresh directory
End If

If bGetDir = True Then
    'If InStr(strData$, "dr") = 1 Or InStr(strData$, "xr") = 1 _
    'Or InStr(strData$, "-r") = 1 Or InStr(strData$, "-x") = 1 Or InStr(strData$, "-w") = 1 Then
        strDir$ = strDir$ & strData$
        Exit Sub
    'End If
End If

Dim filenum As Integer, dlSpot As Long
filenum = FreeFile
'try creating a buffer for incoming binary data.. if bytesTotal + string length > max then
'put data into file and clear buffer
Select Case LCase(lastUsedCommand(0))
    Case "retr"
        If FileLen(strSaveLoc$) = 0 Then
        If lngDLComplete <> 0 Then lngDLComplete = 0
                Open strSaveLoc$ For Binary Access Write As #filenum
                    Put #filenum, 1, strData$
                Close #filenum
            lngDLComplete = lngDLComplete + bytesTotal
            lblProgress.Caption = lngDLComplete & "/" & lngDLSize & " bytes"
            pgDL.Value = lngDLComplete
        Else
                Open strSaveLoc$ For Binary Access Write As #filenum  'Binary As #filenum
                    Put #filenum, FileLen(strSaveLoc$) + 1, strData$
                    'Print #filenum, strData$
                Close #filenum
            lngDLComplete = lngDLComplete + bytesTotal
            lblProgress.Caption = lngDLComplete & "/" & lngDLSize & " bytes"
            pgDL.Value = lngDLComplete
                If lngDLComplete = lngDLSize Then
                    lblProgress.Caption = "Download Complete."
                    lngDLComplete = 0
                    lngDLSize = 0
                End If
        End If
        Exit Sub

    Case "retr ascii"
            Open strSaveLoc$ For Append As #filenum
                Print #filenum, strData$
            Close #filenum
        lngDLComplete = lngDLComplete + bytesTotal
        lblProgress.Caption = lngDLComplete & "/" & lngDLSize & " bytes"
        pgDL.Value = lngDLComplete
            If lngDLComplete = lngDLSize Then
                Shell "notepad.exe " & strSaveLoc$, vbNormalFocus
                lblProgress.Caption = "Download Complete."
                lngDLComplete = 0
                pgDL.Value = 0
            End If
        Exit Sub
    
End Select
'==========================================================================================
'PROBLEM BELOW HAS SINCE BEEN FIXED... KEEP IN CASE OF FUTURE PROBLEMS
'If bytesTotal > 1000 Then 'discard data if contains more than 3000 bytes.. if its not
'    strData$ = vbNullString 'a download or directory listing then must be leaked data..
'    Exit Sub
'End If
'==========================================================================================
'
'
'==========================================================================================
'Define information type getting received... <= 399; OK Response, >=400; Bad Response
'==========================================================================================
Select Case Left(strData$, 3)
    Case Is <= 399
        txtConsole.SelColor = &H4000&
        txtConsole.SelText = strData$
        txtConsole.SelStart = Len(txtConsole.Text)

    Case Is >= 400
        txtConsole.SelColor = &H80&
        txtConsole.SelText = strData$
        txtConsole.SelStart = Len(txtConsole.Text)
    
End Select
End Sub

Private Sub txtConsole_Change()
txtConsole.SelStart = Len(txtConsole.Text)
End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn
    If txtSend.Text = vbNullString Then
        Else
            lst_cmdhist.AddItem txtSend.Text 'creates history of used commands in listbox
            hist = lst_cmdhist.ListCount 'creates index
    End If
        Call cmdSend_Click
        
    Case vbKeyUp
    If hist = 0 Then 'if index is 0 then prevent potential error
        Exit Sub
    End If
    
    hist = hist - 1 'otherwise make index, index - 1
    txtSend.Text = lst_cmdhist.List(hist)
    txtSend.SelStart = Len(txtSend.Text)
    
    
    Case vbKeyDown
    If hist = lst_cmdhist.ListCount Then 'if index is maxed out then prevent 'out of range' error
        Exit Sub
    End If
    
    hist = hist + 1 'otherwise make index, index + 1
    txtSend.Text = lst_cmdhist.List(hist)
    txtSend.SelStart = Len(txtSend.Text)
End Select
End Sub
