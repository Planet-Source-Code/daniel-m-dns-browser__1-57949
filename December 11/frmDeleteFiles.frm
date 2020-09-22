VERSION 5.00
Begin VB.Form frmDeleteFiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Files"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDeleteFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkStartupLog 
      Appearance      =   0  'Flat
      Caption         =   "View Clean Log"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox chkLog 
      Appearance      =   0  'Flat
      Caption         =   "Create Log File"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin DNSBrowser.isButton cmdStartClean 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Icon            =   "frmDeleteFiles.frx":038A
      Style           =   9
      Caption         =   "Start Cleanup"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Enabled         =   0   'False
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
   Begin VB.CheckBox chkCleanDB 
      Appearance      =   0  'Flat
      Caption         =   "Search and Clean Database Files"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   3360
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.CheckBox chkScan 
      Appearance      =   0  'Flat
      Caption         =   "Scan for old files and prompt to delete"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   3000
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.CheckBox chkClearHist 
      Appearance      =   0  'Flat
      Caption         =   "Clear History"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   2640
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.CheckBox chkDeleteSrc 
      Appearance      =   0  'Flat
      Caption         =   "Clear WWW Source"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   2280
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.CheckBox chkDeleteBKUP 
      Appearance      =   0  'Flat
      Caption         =   "Delete Old Backups"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   1920
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin DNSBrowser.isButton cmdCancel 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Icon            =   "frmDeleteFiles.frx":03A6
      Style           =   9
      Caption         =   "Cancel"
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
   Begin VB.Label lblOptions 
      Caption         =   "Cleanup Options -"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label lblinfo 
      Caption         =   $"frmDeleteFiles.frx":03C2
      ForeColor       =   &H00000080&
      Height          =   1215
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   4095
   End
   Begin VB.Image ImgDelete 
      Height          =   720
      Left            =   240
      Picture         =   "frmDeleteFiles.frx":049A
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmDeleteFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub


