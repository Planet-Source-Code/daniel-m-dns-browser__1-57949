VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Internet Options"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider sldTransparency 
      Height          =   255
      Left            =   1320
      TabIndex        =   41
      Top             =   5880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   0
      Max             =   100
      SelStart        =   100
      TickFrequency   =   10
      Value           =   100
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmOptions.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraHomePage"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraTemp"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraMode"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "DNS Settings"
      TabPicture(1)   =   "frmOptions.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDNS2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraDNS"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Privacy"
      TabPicture(2)   =   "frmOptions.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraPB"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Filter Control"
      TabPicture(3)   =   "frmOptions.frx":05DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Miscellaneous"
      TabPicture(4)   =   "frmOptions.frx":05FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraUpdates"
      Tab(4).Control(1)=   "fraBackup"
      Tab(4).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Website Lock"
         ForeColor       =   &H00000080&
         Height          =   4935
         Left            =   -74880
         TabIndex        =   52
         Top             =   480
         Width           =   6135
         Begin VB.TextBox txtUpdatePW 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   345
            IMEMode         =   3  'DISABLE
            Left            =   1800
            PasswordChar    =   "•"
            TabIndex        =   69
            Top             =   4080
            Width           =   2775
         End
         Begin VB.TextBox txtMaskFilter 
            Enabled         =   0   'False
            Height          =   1575
            Left            =   240
            TabIndex        =   67
            Text            =   "                                   CONTENT MASKED"
            Top             =   2280
            Width           =   4335
         End
         Begin VB.PictureBox picManifestF 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   240
            ScaleHeight     =   375
            ScaleWidth      =   5775
            TabIndex        =   63
            Top             =   4440
            Width           =   5775
            Begin VB.OptionButton OptAuto 
               Caption         =   "Request password to allow site"
               Height          =   255
               Index           =   1
               Left            =   3120
               TabIndex        =   65
               Top             =   120
               Width           =   2655
            End
            Begin VB.OptionButton OptAuto 
               Caption         =   "Auto-filter all sites (no exceptions)"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   64
               Top             =   120
               Value           =   -1  'True
               Width           =   2895
            End
         End
         Begin VB.TextBox txtPass 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   345
            IMEMode         =   3  'DISABLE
            Left            =   2160
            PasswordChar    =   "•"
            TabIndex        =   61
            Top             =   960
            Width           =   2415
         End
         Begin VB.CheckBox chkFilter 
            Appearance      =   0  'Flat
            Caption         =   "Enable Filtering"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1320
            TabIndex        =   55
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtAddFilter 
            Height          =   345
            Left            =   1320
            TabIndex        =   54
            Top             =   1800
            Width           =   3255
         End
         Begin VB.ListBox lstFilter 
            Columns         =   4
            Height          =   1500
            ItemData        =   "frmOptions.frx":0616
            Left            =   240
            List            =   "frmOptions.frx":0618
            TabIndex        =   53
            Top             =   2280
            Width           =   4335
         End
         Begin DNSBrowser.isButton cmdAddF 
            Height          =   375
            Left            =   4680
            TabIndex        =   56
            Top             =   1800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Icon            =   "frmOptions.frx":061A
            Style           =   9
            Caption         =   "Add"
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
         Begin DNSBrowser.isButton cmdRemoveF 
            Height          =   375
            Left            =   4680
            TabIndex        =   57
            Top             =   2280
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Icon            =   "frmOptions.frx":0636
            Style           =   9
            Caption         =   "Remove"
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
         Begin DNSBrowser.isButton cmdRemoveAllF 
            Height          =   375
            Left            =   4680
            TabIndex        =   58
            Top             =   2760
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Icon            =   "frmOptions.frx":0652
            Style           =   9
            Caption         =   "Remove All"
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
         Begin DNSBrowser.isButton cmdLogin 
            Height          =   375
            Left            =   4680
            TabIndex        =   66
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Icon            =   "frmOptions.frx":066E
            Style           =   9
            Caption         =   "Login"
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
         Begin DNSBrowser.isButton cmdUpdatePW 
            Height          =   375
            Left            =   4680
            TabIndex        =   70
            Top             =   4080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Icon            =   "frmOptions.frx":068A
            Style           =   9
            Caption         =   "Update PW"
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
         Begin VB.Label Label1 
            Caption         =   "Change Password:"
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   4125
            Width           =   1455
         End
         Begin VB.Label lblPass 
            Caption         =   "Password:"
            Height          =   255
            Left            =   1320
            TabIndex        =   62
            Top             =   1005
            Width           =   855
         End
         Begin VB.Image Image5 
            Height          =   720
            Left            =   240
            Picture         =   "frmOptions.frx":06A6
            Top             =   480
            Width           =   720
         End
         Begin VB.Label Label2 
            Caption         =   "Filter out un-wanted websites for content control."
            Height          =   255
            Left            =   1320
            TabIndex        =   60
            Top             =   360
            Width           =   4695
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            X1              =   120
            X2              =   6000
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Label lblFilterW 
            Caption         =   "Filter Websites/Keywords"
            Height          =   255
            Left            =   1320
            TabIndex        =   59
            Top             =   1560
            Width           =   2535
         End
         Begin VB.Image Image4 
            Height          =   720
            Left            =   360
            Picture         =   "frmOptions.frx":1570
            Top             =   1560
            Width           =   720
         End
      End
      Begin VB.Frame fraUpdates 
         Caption         =   "Browser Updates"
         ForeColor       =   &H00000080&
         Height          =   2175
         Left            =   -74880
         TabIndex        =   47
         Top             =   480
         Width           =   6135
         Begin VB.PictureBox picManifest3 
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   1440
            ScaleHeight     =   855
            ScaleWidth      =   4095
            TabIndex        =   48
            Top             =   1080
            Width           =   4095
            Begin VB.OptionButton OptUp 
               Caption         =   "Don't check for updates. (Not recommended)"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   50
               Top             =   480
               Width           =   4455
            End
            Begin VB.OptionButton OptUp 
               Caption         =   "Automatically check for updates weekly."
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   49
               Top             =   120
               Value           =   -1  'True
               Width           =   4455
            End
         End
         Begin VB.Image Image1 
            Height          =   720
            Left            =   240
            Picture         =   "frmOptions.frx":243A
            Top             =   600
            Width           =   720
         End
         Begin VB.Label lblUinfo 
            Caption         =   $"frmOptions.frx":3304
            Height          =   735
            Left            =   1440
            TabIndex        =   51
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.Frame fraBackup 
         Caption         =   "Miscellaneous Configuration"
         ForeColor       =   &H00000080&
         Height          =   2535
         Left            =   -74880
         TabIndex        =   42
         Top             =   2880
         Width           =   6135
         Begin VB.CheckBox chkBackup 
            Appearance      =   0  'Flat
            Caption         =   "Backup configuration/setting files weekly."
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1440
            TabIndex        =   43
            Top             =   1080
            Width           =   3855
         End
         Begin DNSBrowser.isButton cmdRestoreD 
            Height          =   375
            Left            =   1440
            TabIndex        =   44
            Top             =   1920
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Icon            =   "frmOptions.frx":3390
            Style           =   9
            Caption         =   "Restore Defaults"
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
         Begin VB.Image Image2 
            Height          =   720
            Left            =   240
            Picture         =   "frmOptions.frx":33AC
            Top             =   600
            Width           =   720
         End
         Begin VB.Label lblBackup 
            Caption         =   $"frmOptions.frx":4276
            Height          =   735
            Left            =   1440
            TabIndex        =   46
            Top             =   240
            Width           =   4455
         End
         Begin VB.Line LineSp2 
            BorderColor     =   &H00808080&
            X1              =   120
            X2              =   5880
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label lblDefault 
            BackColor       =   &H00000080&
            BackStyle       =   0  'Transparent
            Caption         =   "Restoring Default Settings will restore all options to default. (Recommended if problems occur)"
            ForeColor       =   &H00000080&
            Height          =   735
            Left            =   3360
            TabIndex        =   45
            Top             =   1680
            Width           =   2535
         End
      End
      Begin VB.Frame fraMode 
         Caption         =   "Internet Browsing"
         ForeColor       =   &H00000080&
         Height          =   1575
         Left            =   120
         TabIndex        =   35
         Top             =   3840
         Width           =   6135
         Begin VB.PictureBox picManifest 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1200
            ScaleHeight     =   375
            ScaleWidth      =   4815
            TabIndex        =   36
            Top             =   840
            Width           =   4815
            Begin VB.OptionButton OptIB 
               Caption         =   "DNS-Based Browsing"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   38
               Top             =   0
               Value           =   -1  'True
               Width           =   1815
            End
            Begin VB.OptionButton OptIB 
               Caption         =   "Normal Browsing"
               Height          =   255
               Index           =   1
               Left            =   2040
               TabIndex        =   37
               Top             =   0
               Width           =   2295
            End
         End
         Begin VB.Label lblIB 
            Caption         =   "DNS-Based Browsing is the core of this program; however, selecting the alternative allows for regular-browser use."
            Height          =   495
            Left            =   1200
            TabIndex        =   39
            Top             =   240
            Width           =   4815
         End
         Begin VB.Image Image3 
            Height          =   720
            Left            =   240
            Picture         =   "frmOptions.frx":431A
            Top             =   360
            Width           =   720
         End
      End
      Begin VB.Frame fraTemp 
         Caption         =   "Temporary Internet Files"
         ForeColor       =   &H00000080&
         Height          =   1575
         Left            =   120
         TabIndex        =   30
         Top             =   2160
         Width           =   6135
         Begin VB.CheckBox chkDisableH 
            Appearance      =   0  'Flat
            Caption         =   "Disable History"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4560
            TabIndex        =   31
            Top             =   1000
            Width           =   1455
         End
         Begin DNSBrowser.isButton cmdDelete 
            Height          =   375
            Left            =   1200
            TabIndex        =   32
            Top             =   960
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            Icon            =   "frmOptions.frx":51E4
            Style           =   9
            Caption         =   "Delete Files..."
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
         Begin DNSBrowser.isButton cmdHist 
            Height          =   375
            Left            =   2880
            TabIndex        =   33
            Top             =   960
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            Icon            =   "frmOptions.frx":5200
            Style           =   9
            Caption         =   "Clear History"
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
         Begin VB.Image imgTemp 
            Height          =   720
            Left            =   240
            Picture         =   "frmOptions.frx":521C
            Top             =   480
            Width           =   720
         End
         Begin VB.Label lblTemp 
            Caption         =   $"frmOptions.frx":6EE6
            Height          =   735
            Left            =   1080
            TabIndex        =   34
            Top             =   240
            Width           =   4935
         End
      End
      Begin VB.Frame fraHomePage 
         Caption         =   "Home Page"
         ForeColor       =   &H00000080&
         Height          =   1575
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   6135
         Begin VB.TextBox txtHomePage 
            Height          =   345
            Left            =   1920
            TabIndex        =   24
            Text            =   "http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57949&lngWId=1"
            Top             =   600
            Width           =   3975
         End
         Begin DNSBrowser.isButton cmduCurrent 
            Height          =   375
            Left            =   1440
            TabIndex        =   25
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Icon            =   "frmOptions.frx":6F70
            Style           =   9
            Caption         =   "Use Current"
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
         Begin DNSBrowser.isButton cmduDefault 
            Height          =   375
            Left            =   3000
            TabIndex        =   26
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Icon            =   "frmOptions.frx":6F8C
            Style           =   9
            Caption         =   "Use Default"
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
         Begin DNSBrowser.isButton cmduBlank 
            Height          =   375
            Left            =   4560
            TabIndex        =   27
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Icon            =   "frmOptions.frx":6FA8
            Style           =   9
            Caption         =   "Use Blank"
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
         Begin VB.Image imgHP 
            Height          =   720
            Left            =   240
            Picture         =   "frmOptions.frx":6FC4
            Top             =   480
            Width           =   720
         End
         Begin VB.Label lblhpInfo 
            Caption         =   "You can specify which page to use for your home page."
            Height          =   255
            Left            =   1200
            TabIndex        =   29
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label lblAddress 
            Caption         =   "Address:"
            Height          =   255
            Left            =   1200
            TabIndex        =   28
            Top             =   650
            Width           =   615
         End
      End
      Begin VB.Frame fraPB 
         Caption         =   "Pop-up Blocker"
         ForeColor       =   &H00000080&
         Height          =   4695
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Width           =   6135
         Begin VB.CheckBox chkPopNotify 
            Appearance      =   0  'Flat
            Caption         =   "Notify me when pop-up is blocked"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1320
            TabIndex        =   76
            Top             =   1200
            Value           =   1  'Checked
            Width           =   2895
         End
         Begin VB.CheckBox chkPopSnd 
            Appearance      =   0  'Flat
            Caption         =   "Play sound when pop-up is blocked"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1320
            TabIndex        =   72
            Top             =   960
            Value           =   1  'Checked
            Width           =   3855
         End
         Begin VB.ListBox lstPB 
            Height          =   1980
            ItemData        =   "frmOptions.frx":7E8E
            Left            =   240
            List            =   "frmOptions.frx":7E90
            TabIndex        =   9
            Top             =   2520
            Width           =   4335
         End
         Begin VB.TextBox txtAddurl 
            Height          =   345
            Left            =   1320
            TabIndex        =   8
            Top             =   2040
            Width           =   3255
         End
         Begin VB.CheckBox chkEnable 
            Appearance      =   0  'Flat
            Caption         =   "Block Pop-ups (Recommended)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1320
            TabIndex        =   6
            Top             =   720
            Value           =   1  'Checked
            Width           =   3495
         End
         Begin DNSBrowser.isButton cmdAdd 
            Height          =   375
            Left            =   4680
            TabIndex        =   16
            Top             =   2040
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Icon            =   "frmOptions.frx":7E92
            Style           =   9
            Caption         =   "Add"
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
         Begin DNSBrowser.isButton cmdRPB 
            Height          =   375
            Left            =   4680
            TabIndex        =   17
            Top             =   3600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Icon            =   "frmOptions.frx":7EAE
            Style           =   9
            Caption         =   "Remove"
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
         Begin DNSBrowser.isButton cmdRAll 
            Height          =   375
            Left            =   4680
            TabIndex        =   18
            Top             =   4080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Icon            =   "frmOptions.frx":7ECA
            Style           =   9
            Caption         =   "Remove All"
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
         Begin DNSBrowser.isButton cmAddCur 
            Height          =   375
            Left            =   4680
            TabIndex        =   71
            Top             =   2520
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Icon            =   "frmOptions.frx":7EE6
            Style           =   9
            Caption         =   "Add Current"
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
         Begin VB.Image ImgAllow 
            Height          =   480
            Left            =   480
            Picture         =   "frmOptions.frx":7F02
            Top             =   1920
            Width           =   480
         End
         Begin VB.Label lblAllow 
            Caption         =   "Allow Pop-ups on these websites:"
            Height          =   255
            Left            =   1320
            TabIndex        =   7
            Top             =   1800
            Width           =   2535
         End
         Begin VB.Line LineSp 
            BorderColor     =   &H00808080&
            X1              =   120
            X2              =   6000
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label lblPBi 
            Caption         =   "Prevents most pop-up windows from appearing."
            Height          =   255
            Left            =   1320
            TabIndex        =   5
            Top             =   360
            Width           =   4695
         End
         Begin VB.Image ImgPB 
            Height          =   720
            Left            =   240
            Picture         =   "frmOptions.frx":87CC
            Top             =   600
            Width           =   720
         End
      End
      Begin VB.Frame fraDNS2 
         Caption         =   "Specific Options"
         ForeColor       =   &H00000080&
         Height          =   2655
         Left            =   -74880
         TabIndex        =   3
         Top             =   2760
         Width           =   6135
         Begin VB.TextBox txtDNSTO 
            Alignment       =   2  'Center
            ForeColor       =   &H00000080&
            Height          =   330
            Left            =   2280
            TabIndex        =   74
            Text            =   "1000"
            Top             =   1440
            Width           =   975
         End
         Begin VB.PictureBox picManifest2 
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   1440
            ScaleHeight     =   1095
            ScaleWidth      =   4215
            TabIndex        =   19
            Top             =   240
            Width           =   4215
            Begin VB.OptionButton OptUd 
               Caption         =   "Automatically Update DNS Database (Recommended)"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   22
               Top             =   120
               Value           =   -1  'True
               Width           =   4455
            End
            Begin VB.OptionButton OptUd 
               Caption         =   "Never Update DNS Database (Use Default)"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   21
               Top             =   480
               Width           =   4455
            End
            Begin VB.OptionButton OptUd 
               Caption         =   "Prompt me whenever update is possible."
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   20
               Top             =   840
               Width           =   4455
            End
         End
         Begin VB.Label lblms 
            Caption         =   "ms."
            Height          =   255
            Left            =   3330
            TabIndex        =   75
            Top             =   1500
            Width           =   375
         End
         Begin VB.Label lblResolveTimeout 
            Caption         =   "Time Out:"
            Height          =   255
            Left            =   1440
            TabIndex        =   73
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label lblDisable 
            Caption         =   "Disabling DNS Database Updates will only resolve a webpage temporarily but will have to resolve it again at next display."
            ForeColor       =   &H00000080&
            Height          =   735
            Left            =   1440
            TabIndex        =   10
            Top             =   1800
            Width           =   4455
         End
         Begin VB.Image ImgDB2 
            Height          =   720
            Left            =   240
            Picture         =   "frmOptions.frx":9696
            Top             =   480
            Width           =   720
         End
      End
      Begin VB.Frame fraDNS 
         Caption         =   "DNS Database"
         ForeColor       =   &H00000080&
         Height          =   2175
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   6135
         Begin DNSBrowser.isButton cmdRestore 
            Height          =   375
            Left            =   2040
            TabIndex        =   14
            Top             =   1200
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Icon            =   "frmOptions.frx":A560
            Style           =   9
            Caption         =   "Restore Default"
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
         Begin DNSBrowser.isButton cmdBackup 
            Height          =   375
            Left            =   3840
            TabIndex        =   15
            Top             =   1200
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Icon            =   "frmOptions.frx":A57C
            Style           =   9
            Caption         =   "Backup Current"
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
         Begin VB.Label lblDNSinfo 
            Caption         =   "You can restore the default DNS Database if you find the database is growing too large. A backup is kept in case of problems."
            Height          =   735
            Left            =   1320
            TabIndex        =   2
            Top             =   240
            Width           =   4575
         End
         Begin VB.Image ImgDB 
            Height          =   720
            Left            =   240
            Picture         =   "frmOptions.frx":A598
            Top             =   600
            Width           =   720
         End
      End
   End
   Begin DNSBrowser.isButton cmdOK 
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   5760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Icon            =   "frmOptions.frx":C262
      Style           =   9
      Caption         =   "OK"
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
   Begin DNSBrowser.isButton cmdCancel 
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   5760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Icon            =   "frmOptions.frx":C27E
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
   Begin DNSBrowser.isButton cmdApply 
      Height          =   375
      Left            =   5400
      TabIndex        =   13
      Top             =   5760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Icon            =   "frmOptions.frx":C29A
      Style           =   9
      Caption         =   "Apply"
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
   Begin VB.Label lblTrans 
      Caption         =   "Transparency:"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   5880
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objAlpha As clsAlpha
Dim LoginSucceeded As Byte
Dim pwattempted As Long
Dim strFilterFile As String, strToEncrypt As String, strEncrypted As String
Dim strDecrypted As String, strDecryptedArr() As String, strPassComp As String
Private Sub chkBackup_Click()
EnableApply
End Sub
Private Sub chkDisableH_Click()
EnableApply
End Sub
Private Sub chkEnable_Click()
EnableApply
End Sub
Private Sub chkFilter_Click()
EnableApply
End Sub

Private Sub chkPopNotify_Click()
EnableApply
End Sub

Private Sub chkPopSnd_Click()
EnableApply
End Sub

Private Sub cmAddCur_Click()
Dim i As Long
For i = 0 To lstPB.ListCount - 1
    If lstPB.List(i) = txtAddurl.Text Then
        MsgBox "Duplicate Entry, cannot add.", vbInformation, "Duplicate Entry"
        Exit Sub
    End If
    DoEvents
Next i
lstPB.AddItem frmBrowser.wb(curWB).LocationURL
EnableApply
End Sub

Private Sub cmdAdd_Click()
Dim i As Long
For i = 0 To lstPB.ListCount - 1
    If lstPB.List(i) = txtAddurl.Text Then
        MsgBox "Duplicate Entry, cannot add.", vbInformation, "Duplicate Entry"
        Exit Sub
    End If
    DoEvents
Next i
lstPB.AddItem txtAddurl.Text
txtAddurl.Text = vbNullString
EnableApply
End Sub
Private Sub cmdAddF_Click()
EnableApply
Dim i As Long
For i = 0 To lstFilter.ListCount - 1
    If lstFilter.List(i) = txtAddFilter.Text Then
        MsgBox "Duplicate Entry, cannot add.", vbInformation, "Duplicate Entry"
        Exit Sub
    End If
    DoEvents
Next i
lstFilter.AddItem txtAddFilter.Text
txtAddFilter.Text = vbNullString
End Sub
Private Sub cmdApply_Click()
SaveFilterSettings
SaveFilterList
ApplySettings
cmdApply.Enabled = False
End Sub
Private Sub SaveFilterSettings()
Dim EncryptOne$, EncryptTwo$, EncryptThree$

EncryptKey chkFilter.Value, EncryptOne$, "37.285.17.32"
EncryptKey strFilterPass$, EncryptTwo$, "68.158.53.60"
EncryptKey OptAuto(0).Value, EncryptThree$, "45.192.45.21"

If LoginSucceeded = 1 Then
    WriteString "Filter Settings", "ENABLEFILTER", EncryptOne$, "DATA\options.dat"
    WriteString "Filter Settings", "PASSWORD", EncryptTwo$, "DATA\options.dat"
    WriteString "Filter Settings", "AUTOFILTER", EncryptThree$, "DATA\options.dat"
End If

End Sub
Private Sub SaveFilterList()
Dim i As Long, strList As String
FileNumber = FreeFile
For i = 0 To lstFilter.ListCount - 1
    strList$ = strList$ & lstFilter.List(i) & " "
Next i


Open "DATA\filterkey.dat" For Output As #FileNumber
    Print #FileNumber, strList$
Close #FileNumber
End Sub
Private Sub cmdBackup_Click()

If MsgBox("Backup Current Database?", vbYesNo, "Backup Current?") = vbYes Then
    Dim bkup
    Set bkup = CreateObject("Scripting.FileSystemObject")
    bkup.CopyFolder "DNS Database\Current", "DNS Database\Current Backup " & Format(Now, "mm-dd-yy")
    MsgBox "Backup Created!", vbInformation, "Backup Created"
Else
    MsgBox "Action has been canceled.", vbInformation, "Action Canceled"
End If
End Sub

Private Sub cmdCancel_Click()
txtPass.Text = vbNullString
LoginSucceeded = 0
Me.Hide
End Sub

Private Sub cmdDelete_Click()
frmDeleteFiles.Show vbModal
End Sub

Private Sub cmdHist_Click()
FileNumber = FreeFile
If MsgBox("Clear History files?", vbYesNo, "Clear History?") = vbYes Then
    Open "DATA\history.dat" For Output As #FileNumber
        Print #FileNumber, "//DNS Browser - History//"
    Close #FileNumber
    Dim tempURL As String
    tempURL = frmBrowser.cboURL.Text
    frmBrowser.cboURL.Clear
    frmBrowser.cboURL.Text = tempURL
    MsgBox "History Cleared!", vbInformation, "History Cleared"
    
    Else
    MsgBox "Action has been canceled.", vbInformation, "Action Canceled"
End If
        
End Sub

Private Sub cmdLogin_Click()
Dim TempStr As String, i As Long
Dim strEncryptOne$, strEncryptTwo$, strEncryptThree$

If blnAutoFilter = False Then

    EncryptKey "1", strEncryptOne$, "37.285.17.32"
    EncryptKey txtPass.Text, strEncryptTwo$, "68.158.53.60"
    EncryptKey "True", strEncryptThree$, "45.192.45.21"
    
    WriteString "Filter Settings", "ENABLEFILTER", strEncryptOne$, "DATA\options.dat"
    WriteString "Filter Settings", "PASSWORD", strEncryptTwo$, "DATA\options.dat"
    WriteString "Filter Settings", "AUTOFILTER", strEncryptThree$, "DATA\options.dat"
        
    MsgBox "First login attempt. Password Created!", vbInformation, "First Time Use"
    LoginSucceeded = 1
    blnAutoFilter = 1
    Else
    
    If strFilterPass$ = txtPass.Text Then
        LoginSucceeded = 1
    Else
        LoginSucceeded = 0
    End If
End If

If LoginSucceeded = 1 Then
    txtMaskFilter.Visible = False
    chkFilter.Enabled = True
    txtAddFilter.Enabled = True
    cmdAddF.Enabled = True
    cmdRemoveF.Enabled = True
    cmdRemoveAllF.Enabled = True
    OptAuto(0).Enabled = True
    OptAuto(1).Enabled = True
    txtUpdatePW.Enabled = True
    cmdUpdatePW.Enabled = True
Else
    pwattempted = pwattempted + 1
    
    If pwattempted <> 3 Then
        MsgBox "The password you have entered is invalid. Please try again.", vbCritical, "Invalid Password"
    Else
        MsgBox "You have attempted to login too many times! Login attempt logged.", vbCritical
        'not yet implemented...
    End If
End If
End Sub

Private Sub cmdOK_Click()
txtPass.Text = vbNullString
If cmdApply.Enabled <> False Then
    SaveFilterSettings
    SaveFilterList
    ApplySettings
End If
LoginSucceeded = 0
LoadOptSettings
LoadFilterList
Me.Hide
End Sub

Private Sub cmdRAll_Click()
If MsgBox("Are you sure you want to clear all entries?", vbYesNo, "Clear all entries?") = vbYes Then
    lstPB.Clear
    EnableApply
    Else
End If
End Sub

Private Sub cmdRemoveAllF_Click()
If MsgBox("Are you sure you want to remove all filtered items?", vbCritical + vbYesNo, "Confirm Clear List") = vbYes Then
    EnableApply
    lstFilter.Clear
Else
    MsgBox "Action Canceled", vbInformation, "Action Canceled"
End If
End Sub

Private Sub cmdRemoveF_Click()
EnableApply
lstFilter.RemoveItem lstFilter.ListIndex
End Sub

Private Sub cmdRestore_Click()
If MsgBox("Are you sure you want to restore the default" & vbNewLine & _
            " DNS Database?", vbYesNo, "Restore Default DNS Database?") = vbYes Then
Dim cdb
Set cdb = CreateObject("Scripting.FileSystemObject")
    If MsgBox("Create backup of current Database?", vbYesNo, "Create Backup?") = vbYes Then
        cdb.CopyFolder "DNS Database\Current", "DNS Database\Current Backup " & Format(Now, "mm-dd-yy") 'Copy Current to Backup
        cdb.DeleteFolder "DNS Database\Current" 'Delete Current
    Else
        cdb.DeleteFolder "DNS Database\Current"
    End If
    cdb.CopyFolder "DNS Database\Default", "DNS Database\Current"
    MsgBox "Default Database Restored." & vbNewLine & _
    "For best results please restart browser.", vbInformation, "Default Database Restored"
    refreshDNS
Else
MsgBox "Restore Database Canceled", vbInformation, "Action Canceled"
End If
End Sub

Private Sub cmdRestoreD_Click()
If MsgBox("This option will save and close the Options Dialog Box." & vbNewLine & _
            "Are you sure you want to continue?", vbYesNo, "Restore Default Settings?") = vbYes Then
Kill "DATA\options.dat"
FileSystem.FileCopy "DATA\default.dat", "DATA\options.dat"
LoadOptSettings
MsgBox "Settings have been restored to default.", vbInformation, "Settings Restored"
Me.Hide
Else
MsgBox "Action has been canceled.", vbInformation, "Action Canceled"
End If
End Sub

Private Sub cmdRPB_Click()
lstPB.RemoveItem lstPB.ListIndex
EnableApply
End Sub

Private Sub cmduBlank_Click()
txtHomePage.Text = "about:blank"
End Sub
Private Sub cmduCurrent_Click()
txtHomePage.Text = frmBrowser.wb(curWB).LocationURL
End Sub
Private Sub cmduDefault_Click()
txtHomePage.Text = "http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57949&lngWId=1"
End Sub

Private Sub cmdUpdatePW_Click()
If MsgBox("Are you sure you want to change your password?", vbQuestion + vbYesNo, "Password Confirmation") = vbYes Then
    strFilterPass$ = txtUpdatePW.Text
    SaveFilterSettings
    MsgBox "Password Successfully Changed!", vbInformation, "Password Successfully Changed"
    txtUpdatePW.Text = vbNullString
Else
    MsgBox "Password Change Canceled!"
    txtUpdatePW.Text = vbNullString
End If
End Sub

Private Sub Form_Activate()
    txtMaskFilter.Visible = True
    txtPass.Enabled = True
    chkFilter.Enabled = False
    txtAddFilter.Enabled = False
    cmdAddF.Enabled = False
    cmdRemoveF.Enabled = False
    cmdRemoveAllF.Enabled = False
    OptAuto(0).Enabled = False
    OptAuto(1).Enabled = False
    txtUpdatePW.Enabled = False
    cmdUpdatePW.Enabled = False
    LoadFilterList
Set objAlpha = New clsAlpha
cmdApply.Enabled = False
End Sub
Private Sub OptIB_Click(Index As Integer)
EnableApply
End Sub

Private Sub OptUd_Click(Index As Integer)
EnableApply
End Sub

Private Sub OptUp_Click(Index As Integer)
EnableApply
End Sub

Private Sub sldTransparency_Change()
If sldTransparency.Value <= 10 Then
    sldTransparency.Value = 10
End If
Set objAlpha = New clsAlpha
objAlpha.SetLayered Me.hwnd, True, CByte((sldTransparency.Value * 2.5))
End Sub


Private Sub sldTransparency_Scroll()
EnableApply
End Sub

Private Sub txtDNSTO_Change()
EnableApply
End Sub

Private Sub txtHomePage_Change()
EnableApply
End Sub
Private Function EnableApply()
cmdApply.Enabled = True
End Function
