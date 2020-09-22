VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B}#4.0#0"; "MSHTML.TLB"
Begin VB.Form frmBrowser 
   AutoRedraw      =   -1  'True
   Caption         =   "DNS Browser"
   ClientHeight    =   9585
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12345
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DNSbrowser.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "wb"
   ScaleHeight     =   9585
   ScaleWidth      =   12345
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   1785
      ScaleWidth      =   3105
      TabIndex        =   29
      Top             =   3600
      Visible         =   0   'False
      Width           =   3135
      Begin DNSBrowser.isButton cmdSearchTitle 
         DragMode        =   1  'Automatic
         Height          =   345
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   609
         Icon            =   "DNSbrowser.frx":08CA
         Style           =   10
         Caption         =   "DNS Browser Search"
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
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H006C5F57&
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   120
         TabIndex        =   33
         Top             =   1050
         Width           =   2895
      End
      Begin VB.ComboBox cboSite 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         ItemData        =   "DNSbrowser.frx":08E6
         Left            =   0
         List            =   "DNSbrowser.frx":091D
         TabIndex        =   31
         Text            =   "---------Select Search Engine--------"
         Top             =   350
         Width           =   3135
      End
      Begin VB.Label cmdSearch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00616161&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Search"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   35
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblCloseSearch 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2730
         TabIndex        =   30
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblSearch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ".:: Search Query ::."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   760
         Width           =   2895
      End
   End
   Begin MSHTMLCtl.Scriptlet Scriptlet1 
      Height          =   375
      Left            =   720
      TabIndex        =   28
      Top             =   480
      Visible         =   0   'False
      Width           =   135
      Scrollbar       =   0   'False
      URL             =   "about:blank"
   End
   Begin VB.PictureBox picPopup 
      Appearance      =   0  'Flat
      BackColor       =   &H00E8FBFB&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   12345
      TabIndex        =   24
      Top             =   840
      Visible         =   0   'False
      Width           =   12375
      Begin VB.Label lblClose 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12100
         TabIndex        =   26
         Top             =   30
         Width           =   255
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Pop-up blocked. To see this pop-up or additional options click here..."
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   75
         Width           =   5295
      End
      Begin VB.Image imgPopup 
         Height          =   240
         Left            =   120
         Picture         =   "DNSbrowser.frx":09C7
         Top             =   50
         Width           =   240
      End
   End
   Begin VB.ComboBox cboURL 
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   840
      TabIndex        =   0
      Text            =   "http://"
      Top             =   480
      Width           =   10905
   End
   Begin VB.ListBox lstUnloaded 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      ItemData        =   "DNSbrowser.frx":0F51
      Left            =   6600
      List            =   "DNSbrowser.frx":0F53
      Sorted          =   -1  'True
      TabIndex        =   22
      Top             =   9600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox ColURL 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   3120
      TabIndex        =   21
      Top             =   9600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSComctlLib.ProgressBar pgBar 
      Height          =   230
      Left            =   8200
      TabIndex        =   20
      Top             =   9305
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
   End
   Begin DNSBrowser.isButton cmdGo 
      Height          =   360
      Left            =   11760
      TabIndex        =   1
      Top             =   480
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   635
      Icon            =   "DNSbrowser.frx":0F55
      Style           =   8
      Caption         =   "Go"
      IconAlign       =   1
      CaptionAlign    =   2
      iNonThemeStyle  =   0
      HighlightColor  =   255
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
   Begin VB.FileListBox lstFav 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4200
      Pattern         =   "*.fav"
      TabIndex        =   19
      Top             =   9600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar StatBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   9210
      Width           =   12345
      _ExtentX        =   21775
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14289
            MinWidth        =   14289
            Text            =   "Done"
            TextSave        =   "Done"
            Key             =   "kStatus"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2381
            MinWidth        =   2381
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "3:20 PM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "12/30/2004"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtChar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7080
      TabIndex        =   17
      Text            =   $"DNSbrowser.frx":12EF
      Top             =   10440
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   15
      Top             =   11280
      Visible         =   0   'False
      Width           =   975
   End
   Begin InetCtlsObjects.Inet findDNS 
      Left            =   7200
      Top             =   10680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AccessType      =   1
      Protocol        =   4
      URL             =   "http://"
   End
   Begin MSComctlLib.ImageList ImgHot 
      Left            =   9360
      Top             =   10560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DNSbrowser.frx":12F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DNSbrowser.frx":184E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DNSbrowser.frx":1DA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DNSbrowser.frx":2302
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DNSbrowser.frx":285C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DNSbrowser.frx":2DB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DNSbrowser.frx":3310
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DNSbrowser.frx":386A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DNSbrowser.frx":3DC4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgToolbar 
      Left            =   8640
      Top             =   10560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DNSbrowser.frx":431E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DNSbrowser.frx":4A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DNSbrowser.frx":5212
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DNSbrowser.frx":598C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DNSbrowser.frx":6106
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DNSbrowser.frx":6880
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DNSbrowser.frx":6FFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DNSbrowser.frx":7774
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DNSbrowser.frx":7EEE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolBar 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12345
      _ExtentX        =   21775
      _ExtentY        =   847
      ButtonWidth     =   2170
      ButtonHeight    =   794
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImgToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "keyBack"
            Description     =   "Go Back"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Description     =   "Go Forward"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Description     =   "Stop"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Description     =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Home"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Favorites"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "History"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Debugger"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.ListBox lstNZ 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   11400
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox lstNZ2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   11160
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.ListBox lstAM 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   10920
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox lstAM2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   10680
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.ListBox lstQZ 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   10560
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox lstQZ2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   10800
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.ListBox lstIP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   10200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox lstIP2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   10440
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox txtDebug 
      Height          =   8415
      Left            =   6910
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "DNSbrowser.frx":8668
      Top             =   840
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.TextBox txtURLx 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   10080
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.ListBox lstAH2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   10080
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.ListBox lstAH 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   9840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSComctlLib.Slider sldTransparency 
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   0
      Max             =   100
      SelStart        =   100
      TickFrequency   =   10
      Value           =   100
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   8415
      Index           =   0
      Left            =   -120
      TabIndex        =   27
      Top             =   840
      Width           =   7095
      ExtentX         =   12515
      ExtentY         =   14843
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.TextBox txtHandle 
      Height          =   330
      Left            =   240
      TabIndex        =   36
      Top             =   120
      Width           =   150
   End
   Begin VB.Label lblAddress 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   520
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Begin VB.Menu mnuFileNewWindow 
            Caption         =   "Window"
            Shortcut        =   {F1}
         End
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFilePrintPrev 
         Caption         =   "Print Pre&view..."
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Properties..."
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find (on This Page)..."
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuToolbars 
         Caption         =   "&Toolbars"
         Begin VB.Menu mnuStandardButtons 
            Caption         =   "&Standard Buttons"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuAddressBar 
            Caption         =   "&Address Bar"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuDebugConsole 
         Caption         =   "&Debug Console"
      End
      Begin VB.Menu mnuViewSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Sto&p"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTextSize 
         Caption         =   "&Text Size"
         Begin VB.Menu mnuTextSizeX 
            Caption         =   "Largest"
            Index           =   0
         End
         Begin VB.Menu mnuTextSizeX 
            Caption         =   "Larger"
            Index           =   1
         End
         Begin VB.Menu mnuTextSizeX 
            Caption         =   "Medium"
            Index           =   2
         End
         Begin VB.Menu mnuTextSizeX 
            Caption         =   "Smaller"
            Index           =   3
         End
         Begin VB.Menu mnuTextSizeX 
            Caption         =   "Smallest"
            Index           =   4
         End
      End
      Begin VB.Menu mnuViewStyle 
         Caption         =   "&Style Sheet"
         Begin VB.Menu mnuViewHC 
            Caption         =   "(&1) High Contrast"
         End
         Begin VB.Menu mnuViewBW 
            Caption         =   "(&2) Black/White"
         End
      End
      Begin VB.Menu mnuViewSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewHTML 
         Caption         =   "HTML"
         Begin VB.Menu mnuViewSource 
            Caption         =   "&Source"
         End
         Begin VB.Menu mnuViewText 
            Caption         =   "Te&xt"
         End
         Begin VB.Menu mnuViewLinks 
            Caption         =   "&Links"
         End
         Begin VB.Menu mnuViewEmails 
            Caption         =   "Emails"
         End
         Begin VB.Menu mnuViewImgSrc 
            Caption         =   "&Images"
         End
      End
   End
   Begin VB.Menu mnuFavorites 
      Caption         =   "F&avorites"
      Begin VB.Menu mnuAddFav 
         Caption         =   "&Add to Favorites..."
      End
      Begin VB.Menu mnuOrganizeFav 
         Caption         =   "&Organize Favorites..."
      End
      Begin VB.Menu mnuFavSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFavoriteT 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuQDebug 
         Caption         =   "&Quick Debug"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuRefreshDNS 
         Caption         =   "&Refresh DNS"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuPopupBlocker 
         Caption         =   "&Pop-up Blocker"
         Begin VB.Menu mnuPopupBlockSet 
            Caption         =   "Turn &Off Pop-up Blocker"
         End
      End
      Begin VB.Menu mnuOptionsSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInternetOptions 
         Caption         =   "Internet &Options..."
      End
   End
   Begin VB.Menu mnuToolS 
      Caption         =   "&Tools"
      Begin VB.Menu mnuDNS 
         Caption         =   "&DNS Editor..."
      End
      Begin VB.Menu mnuToolsFTP 
         Caption         =   "FTP Client..."
      End
      Begin VB.Menu mnuWhois 
         Caption         =   "&WHOIS Client..."
      End
      Begin VB.Menu mnuToolsSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolUpdates 
         Caption         =   "&Check For Updates"
      End
   End
   Begin VB.Menu mnuPlugins 
      Caption         =   "&Plug-Ins"
      Begin VB.Menu mnuDNSMessenger 
         Caption         =   "DNS &Messenger"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows - [1]"
      Begin VB.Menu mnuWindowsWB 
         Caption         =   "Welcome to DNS Browser..."
         Checked         =   -1  'True
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents and Index"
      End
      Begin VB.Menu mnuHelpTOTD 
         Caption         =   "&Tip of the Day"
      End
      Begin VB.Menu mnuHelpSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About DNS Browser"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Admin To&ken"
      Begin VB.Menu mnuFlashLight 
         Caption         =   "Flash Light"
      End
      Begin VB.Menu mnuWatchMovie 
         Caption         =   "Watch Movie"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuStayTop 
         Caption         =   "Stay on &Top"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuLockWorkstation 
         Caption         =   "Lock Workstation"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuAdminHide 
         Caption         =   "Quick Hide"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpAddList 
         Caption         =   "Allow This Site (Add to List)"
      End
      Begin VB.Menu mnuPopupOptions 
         Caption         =   "Preferences"
         Begin VB.Menu mnuPopUpNotification 
            Caption         =   "Disable Notification"
         End
         Begin VB.Menu mnuPopupDisableSnd 
            Caption         =   "Disable Sound Notification"
         End
         Begin VB.Menu mnuPopupTurnOff 
            Caption         =   "Turn Off Pop-up Blocker"
         End
      End
   End
   Begin VB.Menu mnuDebugPopup 
      Caption         =   "Debug Console Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuDebugCopy 
         Caption         =   "Copy Text"
      End
      Begin VB.Menu mnuDebugPaste 
         Caption         =   "Paste Text"
      End
      Begin VB.Menu mnuDebugSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearConsole 
         Caption         =   "Clear Console"
      End
   End
   Begin VB.Menu mnuWBMenu 
      Caption         =   "WB Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuWBOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuWBOpenNew 
         Caption         =   "Open In New Window"
      End
      Begin VB.Menu mnuWBBack 
         Caption         =   "&Back"
      End
      Begin VB.Menu mnuWBForward 
         Caption         =   "For&ward"
      End
      Begin VB.Menu mnuWBSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWBCut 
         Caption         =   "C&ut"
      End
      Begin VB.Menu mnuWBCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuWBPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuWBSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWBSelectAll 
         Caption         =   "&Select All"
      End
      Begin VB.Menu mnuWBFind 
         Caption         =   "&Find (on This Page)..."
      End
      Begin VB.Menu mnuWBSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWBFav 
         Caption         =   "&Add to Favorites..."
      End
      Begin VB.Menu mnuWBView 
         Caption         =   "&View HTML"
         Begin VB.Menu mnuWBViewSource 
            Caption         =   "&Source"
         End
         Begin VB.Menu mnuWBViewText 
            Caption         =   "Text"
         End
         Begin VB.Menu mnuWBViewLinks 
            Caption         =   "Links"
         End
         Begin VB.Menu mnuWBViewEmails 
            Caption         =   "Emails"
         End
         Begin VB.Menu mnuWBViewImages 
            Caption         =   "Images"
         End
      End
      Begin VB.Menu mnuWBSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWBPrint 
         Caption         =   "Pri&nt"
      End
      Begin VB.Menu mnuWBRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuWBSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWBDebug 
         Caption         =   "Quick &Debug"
      End
      Begin VB.Menu mnuWBProperties 
         Caption         =   "Pr&operties"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "Search Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuHideSearch 
         Caption         =   "Hide Search"
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===================================================================================
'Author: Daniel M.
'Project Description: This project was originally intended to be a way for
'me to access the internet at my workstation in college because they had
'disabled DNS and I had to find a way around this. So I decided to make a
'browser that would get the IP address for the DNS and would place it in a
'database which the browser would search through to decide where to navigate.
'Well, after spending over two months working on it I've added many different
'features and a lot of code that I think many people will find useful.
'
'NOTED FEATURES -
'1. Fully functional web browser which can be set to DNS or Normal Browsing Mode.
'2. Fully functional FTP client included.
'3. A whoIs client that is fully functional and includes 6 servers.
'4. Implemented own "favorites" system which isn't very good but oh well, lol.
'5. Implementation of many additional features for browser such as "style changing"
'for people who cannot read well and extracting links, emails, or image urls.
'6. Implementation of pop-up blocker as well as a feature to filter website content.
'7. Some features you will have to find on your own because I never got around to
'making a help file or anything for this. So just search through the code and see
'what you can find. Notably, check the form keypress section.
'8. PASSWORD FOR FILTERING IN BROWSER OPTIONS IS (CASE-SENSITIVE): "admin"
'9. Also included is the feature to block ad-banners and ad-text but must be uncommented
'to be active. I didn't have the time to include it in the options.
'
'BUGS: There are MANY bugs in this program as you might imagine because the code is
'quite large (over 120 pages when pasted in Microsoft Word). So please just tell me
'if you find any bugs but please don't criticize me for them. Note that I grabbed a
'few things from PSC and have noted it below.
'
'NOTED UNFINISHED WORK: (1) FTP is fully functional but the file listing does not
'always work properly and I never got around to finishing that. (2) The "clean-up" bit
'in the Browser Options was never completed and little features like the (3) updating
'browser weekly and (4) backing up files weekly was never fully completed. (5) I was
'intending on creating a mini-chat application between browser users but never got
'around to even starting it. (6) Adding whois servers has not yet been included. (7)
'Was going to add a history feature but never got around to that either. Wow, so many
'things, lol.
'
'CREDITS: I give credit to whoever's code the Alpha Transparancy is as I did not make
'that. I also give credit to the person who provided the "isButton" control as I also
'did not write that. And finally I give credit to the person who provided the .RES
'for the XP Manifest code.
'
'Extra Comments: This is my second largest project I've made and I've spent about
'two months or so working on it so I hope you appreciate me distributing
'this code. Sorry for the lack of commenting in this code! Maybe when I get around to
'it I'll submit an update with commenting. I didn't intend on submitting this code
'so yeah. And also, before you close this program and go onto doing whatever else
'you do, please vote for me! I spent many long hard hours in this program!
'
'
'CONTACT: Email - SeoulxKorean@yahoo.com      AIM - xAznHangukBoix
'=================================================================================
'Option Explicit
'Some variables may be undefined
Private Declare Function LockWorkStation Lib "user32" () As Long
Private breakURL$, breakURLbn$, maskURL$, pageURL$, Nav$, Source$
Private chkStart As Byte, skipNavigateChk As Byte, DuplicateEntry As Byte, DropDown As Byte
Private SearchProperList As Integer, IndexPages As Integer, OnPage As Integer, chkGoBack As Byte, chkNav As Byte
Private chkGoForward As Byte, chkAlreadyResolved As Byte
Private timeX As Integer, bLink As Boolean, NavigateAlreadyChecked As Boolean, MenuNewWindow As Boolean
Private ColUnloaded As New Collection, strFoundInDB As String, initLoad As Byte
Private WindowCount As Long, srchX As Long, skipResolveBadNav As Byte
Private WithEvents Web_V1 As SHDocVwCtl.WebBrowser_V1
Attribute Web_V1.VB_VarHelpID = -1
Dim strNewWindow As String, ClickURL As String
Dim ProcDispN As Byte, lngAddSubLoc
Private objAlpha As clsAlpha
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10
Private TempAllow As Byte
Private allowOnce As New Collection
Public WithEvents IEDoc As HTMLDocument
Attribute IEDoc.VB_VarHelpID = -1
Dim chkLoop As Long

'Public WithEvents IELINK As HTMLLinkElement
'As Hyperlink

Private Sub cboURL_Change()
cboURL.Tag = cboURL.Text

End Sub

Private Sub cboURL_Click()
If DropDown = 1 Then
    Call cmdGo_Click
End If
End Sub

Private Sub cboURL_DropDown()
DropDown = 1
End Sub

Private Sub cboURL_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyAdd Then
    If Shift = 1 Then
        If lngAddSubLoc = "0" Then
            If cboURL.SelLength <> Len(cboURL) Then
                lngAddSubLoc = Len(cboURL)
            Else
                lngAddSubLoc = 1
            End If
        End If
    End If
End If

If KeyCode = vbKeySubtract Then
    If Shift = 1 Then
        If lngAddSubLoc = "0" Then
            If cboURL.SelLength <> Len(cboURL) Then
                lngAddSubLoc = Len(cboURL)
            Else
                lngAddSubLoc = 1
            End If
        End If
    End If
End If

DropDown = 0
Select Case KeyCode
    Case vbKeyReturn
        If Shift <> 1 Then
        Else
            cboURL.Text = "http://www." & cboURL.Text & ".com"
        End If
        If ValHist <> 1 Then
            Dim i As Long
            For i = 0 To cboURL.ListCount - 1
                If cboURL.Text = cboURL.List(i) Then
                    DuplicateEntry = 1
                End If
            Next i
            
            If DuplicateEntry = 0 Then
                cboURL.AddItem cboURL.Text
            End If
        End If
        Call cmdGo_Click
        
        DuplicateEntry = 0
    
    Case vbKeyDelete
        If Shift = 1 Then
            If MsgBox("Delete item from history?", vbOKCancel, "Delete item from History?") = vbOK Then
                cboURL.RemoveItem cboURL.ListIndex
            Else
            End If
        End If
        
End Select

End Sub

Private Sub cboURL_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyBack Then Exit Sub
If KeyAscii = vbKeyDelete Then Exit Sub
If KeyAscii = vbKeyReturn Then Exit Sub
If cboURL.SelLength = Len(cboURL.Text) Then Exit Sub
KeyAscii = AutoComplete(cboURL, KeyAscii, False)
End Sub

Private Sub cboURL_KeyUp(KeyCode As Integer, Shift As Integer)

'handles onfocus error if user
If lngAddSubLoc <> 0 Then
    cboURL.SelStart = lngAddSubLoc
    cboURL.SelLength = Len(cboURL) - lngAddSubLoc
    cboURL.SelText = ""
    
    cboURL.SelStart = cboURL.SelStart - 1
    cboURL.SelLength = 1
    If cboURL.SelText = "+" Or cboURL.SelText = "-" Then
        cboURL.SelText = ""
    Else
        cboURL.SelStart = Len(cboURL)
    End If
    
    lngAddSubLoc = 0
End If

End Sub

Private Sub cboURL_LostFocus()
DropDown = 0
End Sub
Private Sub cmdGo_Click()
'if instr(i, cboURL.Text, ""

FilterCheck (cboURL.Text)

If xCancel = 1 Then
    xCancel = 0
    Exit Sub
End If

Dim i As Long
'If ValHist <> 1 Then
'    cboURL.AddItem cboURL.Text
'End If
If cboURL.Text = "about:blank" Then
    wb(curWB).Navigate "about:blank"
    Exit Sub
End If
If frmOptions.OptIB.Item(0).Value = True Then
    chkNav = 0
    txtDebug.Text = txtDebug.Text & vbNewLine & "Designated URL: " & cboURL.Text
    breakURL$ = Replace(cboURL.Text, "http://", "")
    ' NOTE, use FUNCTIONS to perform tasks; load is otherwise too large to work in control
    
    If Val(Left(breakURL$, 2)) > 9 Then
        'wb(curWB).Tag = breakURL$
        wb(curWB).Navigate breakURL$
        txtDebug.Text = txtDebug.Text & vbNewLine & "Target URL: " & breakURL$
        chkNav = 1
        skipNavigateChk = 1
        Exit Sub
    End If
    
    checkDNS breakURL$, "1"
    If chkNav <> 1 Then
        txtDebug.Text = txtDebug.Text & vbNewLine & "Specified URL not found in Database. Resolving..."
        StatBar.Panels(1).Text = "Specified URL not found in Database. Resolving..."
        If ValDBS = 3 Then
                If MsgBox("Would you like to Resolve/Add this website to the Database?" & vbNewLine & _
                        "URL: '" & breakURL$ & "'", vbYesNo, "Resolve/Add to Database?") = vbYes Then
                    ResolveAdd = 1
                Else
                    ResolveAdd = 0
                End If
            resolveDNS breakURL$
        Else
            If ValDBS = 1 Then
                ResolveAdd = 1
                resolveDNS breakURL$
            End If
            
            If ValDBS = 2 Then
                ResolveAdd = 0
                resolveDNS breakURL$
            End If
        End If
        
    End If
    txtDebug.Text = txtDebug.Text & vbNewLine
Else
    wb(curWB).Navigate cboURL.Text
    '
End If
End Sub
Private Function checkDNS(ByRef TheURL, TOS As String)
Dim i As Long 'this function checks the dns, sorry for the lack of commenting.. =/
On Error Resume Next
If Left(TheURL, 4) = "www." Then 'if left four is WWW. then
    SearchProperList = Asc(Mid(UCase(TheURL), 5, 1)) - 64 'get ascii val of 5th letter
    Select Case SearchProperList 'case is ascii val
        Case Is <= 13 'if A through M then
            For i = 0 To lstAM.ListCount - 1
                If lstAM.List(i) = Left(LCase(TheURL), Len(lstAM.List(i))) Then
                    Nav$ = lstAM2.List(i) & Right(TheURL, Len(TheURL) - Len(lstAM.List(i)))
                    cboURL.Text = Nav$
                    cboURL.Tag = cboURL.Text
                    txtDebug.Text = txtDebug.Text & vbNewLine & "Target URL: " & Nav$
                    strFoundInDB = "AM" & i
                    wb(curWB).Navigate Nav$
                        If TOS = "1" Then
                            chkNav = 1
                        Else
                            skipNavigateChk = 0
                        End If
                    Exit For
                    Exit Function
                End If
                
                If lstAM2.List(i) = Left(TheURL, Len(lstAM2.List(i))) Then
                    If TOS = 3 Then
                        pageURL$ = Right(TheURL, Len(TheURL) - Len(lstAM2.List(i)))
                        cboURL.Text = "http://" & lstAM.List(i) & pageURL$
                        cboURL.Tag = cboURL.Text
                        Exit Function
                    End If
                        Nav$ = lstAM2.List(i) & Right(TheURL, Len(TheURL) - Len(lstAM2.List(i)))
                        txtDebug.Text = txtDebug.Text & vbNewLine & "Target URL: " & Nav$
                        strFoundInDB = "AM" & i
                        wb(curWB).Navigate Nav$
                            If TOS = "1" Then
                                chkNav = 1
                            Else
                                skipNavigateChk = 1
                            End If
                        Exit For
                        Exit Function
                End If
                DoEvents
            Next i
                
            For i = 0 To lstNZ2.ListCount - 1
                If lstNZ2.List(i) = Left(TheURL, Len(lstNZ2.List(i))) Then
                    If TOS = 3 Then
                        pageURL$ = Right(TheURL, Len(TheURL) - Len(lstNZ2.List(i)))
                        cboURL.Text = "http://" & lstNZ.List(i) & pageURL$
                        cboURL.Tag = cboURL.Text
                        Exit Function
                    End If
                        Nav$ = lstNZ2.List(i) & Right(TheURL, Len(TheURL) - Len(lstNZ2.List(i)))
                        txtDebug.Text = txtDebug.Text & vbNewLine & "Target URL: " & Nav$
                        strFoundInDB = "NZ" & i
                        wb(curWB).Navigate Nav$
                            If TOS = "1" Then
                                chkNav = 1
                            Else
                                skipNavigateChk = 1
                            End If
                        Exit For
                        Exit Function
                End If
                DoEvents
            Next i
            
        Case Is >= 14
            For i = 0 To lstNZ.ListCount - 1
                If lstNZ.List(i) = Left(LCase(TheURL), Len(lstNZ.List(i))) Then
                    Nav$ = lstNZ2.List(i) & Right(TheURL, Len(TheURL) - Len(lstNZ.List(i)))
                    cboURL.Text = Nav$
                    cboURL.Tag = cboURL.Text
                    txtDebug.Text = txtDebug.Text & vbNewLine & "Target URL: " & Nav$
                    strFoundInDB = "NZ" & i
                    wb(curWB).Navigate Nav$
                        If TOS = "1" Then
                            chkNav = 1
                        Else
                            skipNavigateChk = 0
                        End If
                    Exit For
                    Exit Function
                End If
                DoEvents
            Next i
    End Select
    
Else

    SearchProperList = Asc(Left(UCase(TheURL), 1)) - 64
    Select Case SearchProperList
        Case Is <= 8
            For i = 0 To lstAH.ListCount - 1
                If lstAH.List(i) = Left(LCase(TheURL), Len(lstAH.List(i))) Then
                    Nav$ = lstAH2.List(i) & Right(TheURL, Len(TheURL) - Len(lstAH.List(i)))
                    cboURL.Text = Nav$
                    cboURL.Tag = cboURL.Text
                    txtDebug.Text = txtDebug.Text & vbNewLine & "Target URL: " & Nav$
                    strFoundInDB = "AH" & i
                    wb(curWB).Navigate Nav$
                        If TOS = "1" Then
                            chkNav = 1
                        Else
                            skipNavigateChk = 0
                        End If
                    Exit For
                    Exit Function
                End If
                
                If lstAH2.List(i) = Left(TheURL, Len(lstAH2.List(i))) Then
                    If TOS = 3 Then
                        pageURL$ = Right(TheURL, Len(TheURL) - Len(lstAH2.List(i)))
                        cboURL.Text = "http://" & lstAH.List(i) & pageURL$
                        cboURL.Tag = cboURL.Text
                        Exit Function
                    End If
                        Nav$ = lstAH2.List(i) & Right(TheURL, Len(TheURL) - Len(lstAH2.List(i)))
                        cboURL.Text = Nav$
                        cboURL.Tag = cboURL.Text
                        txtDebug.Text = txtDebug.Text & vbNewLine & "Target URL: " & Nav$
                        strFoundInDB = "AH" & i
                        wb(curWB).Navigate Nav$
                            If TOS = "1" Then
                                chkNav = 1
                            Else
                                skipNavigateChk = 1
                            End If
                        Exit For
                        Exit Function
                End If
             DoEvents
             Next i
             
             For i = 0 To lstIP2.ListCount - 1
                If lstIP2.List(i) = Left(TheURL, Len(lstIP2.List(i))) Then
                    If TOS = 3 Then
                        pageURL$ = Right(TheURL, Len(TheURL) - Len(lstIP2.List(i)))
                        cboURL.Text = "http://" & lstIP.List(i) & pageURL$
                        cboURL.Tag = cboURL.Text
                        Exit Function
                    End If
                        Nav$ = lstIP2.List(i) & Right(TheURL, Len(TheURL) - Len(lstIP2.List(i)))
                        cboURL.Text = Nav$
                        cboURL.Tag = cboURL.Text
                        txtDebug.Text = txtDebug.Text & vbNewLine & "Target URL: " & Nav$
                        strFoundInDB = "IP" & i
                        wb(curWB).Navigate Nav$
                            If TOS = "1" Then
                                chkNav = 1
                            Else
                                skipNavigateChk = 1
                            End If
                        Exit For
                        Exit Function
                    End If
                DoEvents
                Next i
                
            For i = 0 To lstQZ2.ListCount - 1
                If lstQZ2.List(i) = Left(TheURL, Len(lstQZ2.List(i))) Then
                    If TOS = 3 Then
                        pageURL$ = Right(TheURL, Len(TheURL) - Len(lstQZ2.List(i)))
                        cboURL.Text = "http://" & lstQZ.List(i) & pageURL$
                        cboURL.Tag = cboURL.Text
                        Exit Function
                    End If
                        Nav$ = lstQZ2.List(i) & Right(TheURL, Len(TheURL) - Len(lstQZ2.List(i)))
                        cboURL.Text = Nav$
                        cboURL.Tag = cboURL.Text
                        txtDebug.Text = txtDebug.Text & vbNewLine & "Target URL: " & Nav$
                        strFoundInDB = "QZ" & i
                        wb(curWB).Navigate Nav$
                            If TOS = "1" Then
                                chkNav = 1
                            Else
                                skipNavigateChk = 1
                            End If
                        Exit For
                        Exit Function
                End If
                DoEvents
            Next i
        
        Case Is <= 16
            For i = 0 To lstIP.ListCount - 1
                If lstIP.List(i) = Left(LCase(TheURL), Len(lstIP.List(i))) Then
                    Nav$ = lstIP2.List(i) & Right(TheURL, Len(TheURL) - Len(lstIP.List(i)))
                    cboURL.Text = Nav$
                    cboURL.Tag = cboURL.Text
                    txtDebug.Text = txtDebug.Text & vbNewLine & "Target URL: " & Nav$
                    strFoundInDB = "IP" & i
                    wb(curWB).Navigate Nav$
                        If TOS = "1" Then
                            chkNav = 1
                        Else
                            skipNavigateChk = 0
                        End If
                    Exit For
                    Exit Function
                End If
                
                DoEvents
            Next i
        
    
        Case Is >= 17
            For i = 0 To lstQZ.ListCount - 1
                If lstQZ.List(i) = Left(LCase(TheURL), Len(lstQZ.List(i))) Then
                    Nav$ = lstQZ2.List(i) & Right(TheURL, Len(TheURL) - Len(lstQZ.List(i)))
                    cboURL.Text = Nav$
                    cboURL.Tag = cboURL.Text
                    txtDebug.Text = txtDebug.Text & vbNewLine & "Target URL: " & Nav$
                    strFoundInDB = "QZ" & i
                    wb(curWB).Navigate Nav$
                        If TOS = "1" Then
                            chkNav = 1
                        Else
                            skipNavigateChk = 0
                        End If
                    Exit For
                    Exit Function
                End If
                
                DoEvents
            Next i
        
        End Select
End If

'UhOh:
'txtDebug.Text = txtDebug.Text & vbNewLine & "A possible error has occured. It may cause " & _
'"conflict with the program."
End Function

Private Sub cmdSearch_Click()
Dim strSearch As String
If Len(txtSearch) <> 0 Then
    strQuery$ = txtSearch.Text
Else
    MsgBox "Must input search query", vbInformation, "Search Query Needed"
    txtSearch.SetFocus
    Exit Sub
End If
Select Case cboSite.Text
    Case "AllTheWeb"
        strSearch$ = "http://www.alltheweb.com/search?cat=web&cs=utf8&q=" & strQuery$
    Case "AltaVista"
        strSearch$ = "http://www.altavista.com/web/results?itag=wrx&q=" & strQuery$ & "&kgs=1&kls=0"
    Case "Ask Jeeves"
        strSearch$ = "http://web.ask.com/web?q=" & strQuery$ & "&qsrc=0&o=0"
    Case "DogPile"
        strSearch$ = "http://www.dogpile.com/info.dogpl/search/web/" & strQuery$
    Case "Excite"
        strSearch$ = "http://msxml.excite.com/info.xcite/search/web/" & strQuery$
    Case "Google"
        strSearch$ = "http://www.google.com/search?hl=en&lr=&q=" & strQuery$ & "&btnG=Search"
    Case "HotBot"
        strSearch$ = "http://www.hotbot.com/default.asp?query=" & strQuery$ & "&ps=&loc=searchbox&tab=web&provKey=Inktomi&prov=HotBot"
    Case "LookSmart"
        strSearch$ = "http://search.looksmart.com/p/search?tb=web&qt=" & strQuery$
    Case "Lycos"
        strSearch$ = "http://mia-search.mia.lycos.com/default.asp?lpv=1&loc=lycoshp&tab=web&query=" & strQuery$ & "&x=11&y=1"
    Case "Mamma"
        strSearch$ = "http://www.mamma.com/Mamma?qtype=0&query=" & strQuery$
    Case "MetaCrawler"
        strSearch$ = "http://www.metacrawler.com/info.metac/search/web/" & strQuery$
    Case "MSN Search"
        strSearch$ = "http://search.msn.com/results.aspx?FORM=SRCHWB&q=" & strQuery$
    Case "Netscape"
        strSearch$ = "http://search.netscape.com/ns/search?query=" & strQuery$ & "&st=webresults&fromPage=NSCPResults&x=18&y=5"
    Case "Search.com"
        strSearch$ = "http://www.search.com/search?q=" & strQuery$
    Case "ScrubtheWeb"
        strSearch$ = "http://www.scrubtheweb.com/cgi-bin/search.cgi?keyword=" & strQuery$
    Case "Snap"
        strSearch$ = "http://www.snap.com/search.php?query=" & strQuery$ & "&f=1"
    Case "Yahoo!"
        strSearch$ = "http://search.yahoo.com/search?p=" & strQuery$ & "&ei=UTF-8&fr=FP-tab-web-t&fl=0&x=wrt"
    Case Else
        MsgBox "No Search Engine selected. Using default: Google Search", vbInformation, "Auto-Select Search"
        strSearch$ = "http://www.google.com/search?hl=en&lr=&q=" & strQuery$ & "&btnG=Search"
End Select


FilterCheck (strSearch$)

If xCancel = 1 Then
    xCancel = 0
    Exit Sub
End If

Dim i As Long

If frmOptions.OptIB.Item(0).Value = True Then
    chkNav = 0
    txtDebug.Text = txtDebug.Text & vbNewLine & "Designated URL: " & strSearch$
    strSearch$ = Replace(strSearch$, "http://", "")
    ' NOTE, use FUNCTIONS to perform tasks; load is otherwise too large to work in control
    
    If Val(Left(strSearch$, 2)) > 9 Then
        'wb(curWB).Tag = breakURL$
        wb(curWB).Navigate strSearch$
        txtDebug.Text = txtDebug.Text & vbNewLine & "Target URL: " & strSearch$
        chkNav = 1
        skipNavigateChk = 1
        Exit Sub
    End If
    

    
    checkDNS strSearch$, "1"
    If chkNav <> 1 Then
        txtDebug.Text = txtDebug.Text & vbNewLine & "Specified URL not found in Database. Resolving..."
                StatBar.Panels(1).Text = "Specified URL not found in Database. Resolving..."
        If ValDBS = 3 Then
                If MsgBox("Would you like to Resolve/Add this website to the Database?" & vbNewLine & _
                        "URL: '" & strSearch$ & "'", vbYesNo, "Resolve/Add to Database?") = vbYes Then
                    ResolveAdd = 1
                Else
                    ResolveAdd = 0
                End If
                
            'Dim strBaseSearch() As String
            'strBaseSearch = Split(strSearch$, "/", -1, 1)
    
            resolveDNS strSearch$
            If ValDBS = 1 Then
                ResolveAdd = 1
                resolveDNS strSearch$
            End If
            
            If ValDBS = 2 Then
                ResolveAdd = 0
                resolveDNS strSearch$
            End If
        End If
        
    End If
    txtDebug.Text = txtDebug.Text & vbNewLine
Else
    wb(curWB).Navigate strSearch$
    '
End If


End Sub

Private Sub cmdSearch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    cmdSearch.BackColor = &H6C5F57
End If
End Sub

Private Sub cmdSearch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdSearch.BackColor = &H616161
End Sub

Private Sub cmdSearchTitle_Click()
srchX = picSearch.Left
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF9 Then
    If frmHEHE.Visible Then
        frmHEHE.Picture1.Visible = Not frmHEHE.Picture1.Visible
    End If
End If

Dim i As Long
    If KeyCode = vbKeyControl Then
        TempAllow = 1
    End If
    
    If KeyCode = vbKeyAdd Then
        If Shift = 1 Then
           ' If Left(cboURL.Text, 1) = "+" Then cboURL.Text = Right(cboURL.Text, Len(cboURL) - 1)
           ' txtHandle.SetFocus
            If sldTransparency.Value <> sldTransparency.Max Then
                sldTransparency.Value = sldTransparency.Value + 2
            End If
        End If
    End If
    
    If KeyCode = vbKeySubtract Then
        If Shift = 1 Then
            If sldTransparency.Value <> sldTransparency.Min Then
                sldTransparency.Value = sldTransparency.Value - 2
            End If
        End If
    End If
    
    If KeyCode = vbKeyF12 Then
        If sldTransparency.Value <> sldTransparency.Max Then
            sldTransparency.Value = 100
            objAlpha.SetLayered Me.hwnd, True, 255
        End If
    End If
    
    If KeyCode = vbKeyF11 Then
        If sldTransparency.Value <> sldTransparency.Min Then
            sldTransparency.Value = 0
            objAlpha.SetLayered Me.hwnd, True, 0
        End If
    End If


If KeyCode = vbKeyAdd Then
    If Shift = 0 Then
        For i = 0 To 4
            If mnuTextSizeX(i).Checked = True And i <> 0 Then
                mnuTextSizeX(i).Checked = False
                Call mnuTextSizeX_Click(i - 1)
                Exit For
            End If
        Next i
    End If
End If

If KeyCode = vbKeySubtract Then
    If Shift = 0 Then
        For i = 0 To 4
            If mnuTextSizeX(i).Checked = True And i <> 4 Then
                mnuTextSizeX(i).Checked = False
                Call mnuTextSizeX_Click(i + 1)
                Exit For
            End If
        Next i
    End If
End If

If KeyCode = vbKey1 Then wb(curWB).Document.createStyleSheet App.Path & "\DATA\Styles\" & "stylea.css"
If KeyCode = vbKey2 Then wb(curWB).Document.createStyleSheet App.Path & "\DATA\Styles\" & "styleb.css"
If KeyCode = vbKey0 Then wb(curWB).Refresh
End Sub
Private Function BlockAdvert()
Dim lngImgWidth, lngImgHeight, AdProcess As Boolean
        
'must run first to get rid of frames!

        'Below takes care of ads contained in scripts
        Dim i As Long
        Dim allScriptCode As String
        For i = 0 To IEDoc.scripts.Length - 1
            If InStr(LCase(IEDoc.scripts(i).innerHTML), "/ads") Then AdProcess = True
            If InStr(LCase(IEDoc.scripts(i).innerHTML), "doubleclick") Then AdProcess = True
            If InStr(LCase(IEDoc.scripts(i).innerHTML), "adclick") Then AdProcess = True
            If InStr(LCase(IEDoc.scripts(i).innerHTML), "atdmt") Then AdProcess = True
            If InStr(LCase(IEDoc.scripts(i).innerHTML), "banner") Then AdProcess = True
            If InStr(LCase(IEDoc.scripts(i).innerHTML), "falkag") Then AdProcess = True
            If InStr(LCase(IEDoc.scripts(i).innerHTML), "referral") Then AdProcess = True
            
            If AdProcess = True Then
                IEDoc.scripts(i).outerHTML = "<!--// Script Removed //-->"
                Exit For
                AdProcess = False
            End If
            allScriptCode$ = allScriptCode$ & IEDoc.scripts(i).innerHTML
        Next i

    'takes care of normal ad banner images
    For i = 0 To wb(curWB).Document.images.Length - 1
        lngImgWidth = wb(curWB).Document.images(i).Width
        lngImgHeight = wb(curWB).Document.images(i).Height
        
        'find common ad banner sizes
        If lngImgWidth = "728" And lngImgHeight = "90" Then AdProcess = True 'Leaderboard banner
        If lngImgWidth = "408" And lngImgHeight = "60" Then AdProcess = True '
        If lngImgWidth = "408" And lngImgHeight = "64" Then AdProcess = True
        If lngImgWidth = "468" And lngImgHeight = "60" Then AdProcess = True
        If lngImgWidth = "392" And lngImgHeight = "72" Then AdProcess = True
        If lngImgWidth = "300" And lngImgHeight = "250" Then AdProcess = True
        If lngImgWidth = "120" And lngImgHeight = "240" Then AdProcess = True 'Vertical banner
        If lngImgWidth = "160" And lngImgHeight = "600" Then AdProcess = True 'Wide Skyscrapers
        If lngImgWidth = "120" And lngImgHeight = "600" Then AdProcess = True 'Skyscrapers
        If lngImgWidth = "234" And lngImgHeight = "60" Then AdProcess = True 'Half banner
        If lngImgWidth = "125" And lngImgHeight = "125" Then AdProcess = True 'Square banner
        If lngImgWidth = "125" And lngImgHeight = "561" Then AdProcess = True 'Another Skyscraper
        If lngImgWidth = "120" And lngImgHeight = "60" Then AdProcess = True 'Button
        If lngImgWidth = "120" And lngImgHeight = "90" Then AdProcess = True 'Button
        If lngImgWidth = "88" And lngImgHeight = "33" Then AdProcess = True 'Micro button
        
        'find images containing keywords
        If InStr(LCase(wb(curWB).Document.images(i).src), "banner") Then AdProcess = True
        If InStr(LCase(wb(curWB).Document.images(i).src), "ads") Then AdProcess = True
        If InStr(LCase(wb(curWB).Document.images(i).src), "pagead") Then AdProcess = True
        If InStr(LCase(wb(curWB).Document.images(i).src), "fastclick.net") Then AdProcess = True
        
        If AdProcess = True Then
            wb(curWB).Document.images(i).src = ""
            wb(curWB).Document.images(i).Width = 1
            wb(curWB).Document.images(i).Height = 1
        AdProcess = False
        End If
    Next i
   'code can be changed to change all links to IPs!
    AdProcess = False
    'block link ads
    For i = 0 To wb(curWB).Document.links.Length - 1
        If InStr(LCase(wb(curWB).Document.links(i).outerHTML), "referral") Then AdProcess = True
        If InStr(LCase(wb(curWB).Document.links(i).outerHTML), "ads") Then AdProcess = True
        If InStr(LCase(wb(curWB).Document.links(i).outerHTML), "/ad") Then AdProcess = True
        If InStr(LCase(wb(curWB).Document.links(i).outerHTML), "pagead") Then AdProcess = True
        If InStr(LCase(wb(curWB).Document.links(i).outerHTML), "fastclick.net") Then AdProcess = True
        
        If AdProcess = True Then
            'wb(curWB).Document.links(i).href = ""
            wb(curWB).Document.links(i).outerHTML = "<a href=" & vbQuote & "#" & vbQuote & _
            "><font size=-2 face=verdana>Ad Blocked!</font></a>"
        AdProcess = False
        End If
    Next i
    
    'block ads contained in .swf (flash)
    For i = 0 To wb(curWB).Document.embeds.Length - 1
        If InStr(LCase(wb(curWB).Document.embeds(i).src), "/ad") Then AdProcess = True
        If InStr(LCase(wb(curWB).Document.embeds(i).src), "ads") Then AdProcess = True
        If InStr(LCase(wb(curWB).Document.embeds(i).src), "pagead") Then AdProcess = True
        If InStr(LCase(wb(curWB).Document.embeds(i).src), "referral") Then AdProcess = True
        
        If AdProcess = True Then
            wb(curWB).Document.embeds(i).src = "#"
            wb(curWB).Document.embeds(i).Width = 0
            wb(curWB).Document.embeds(i).Height = 0
            AdProcess = False
        End If
    Next i
    AdProcess = False
    Do Until InStr(wb(curWB).Document.documentElement.innerHTML, "google_ads_frame") = 0
        For i = 0 To wb(curWB).Document.All.Length - 1
            If InStr(wb(curWB).Document.All.Item(i).innerHTML, "google_ads_frame") Then
                wb(curWB).Document.All.Item("google_ads_frame").outerText = ""
                Exit For
            End If
        Next i
    Loop
End Function

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
TempAllow = 0
End Sub


Private Sub Form_Load()

'If strUser$ = "user061" Then
'    Dim fso
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    If fso.DriveExists("V:") Then
'    Else
'        Shell "cmd /c net use V: " & vbQuote & "\\dilbert\storage\Project Development" & vbQuote, vbHide
'    End If
'End If
initLoad = 1
lngAddSubLoc = 0
chkAlreadyResolved = 0
Set objAlpha = New clsAlpha
WindowCount = 1
curWB = 0
timeX = 0
'| Initial SETUP |
'0. LOAD SETTINGS
'1. Load HomePage
'2. Check if form load to prevent "nav complete" error
'3. Skip Navigation check set to 0, changes if subpage to 1 to prevent infinite looping
'4. Set the DNS by loading DNSlst.txt
'5. Get information from favorites.dat and load into menu
'6. Get history and load to address bar
'7. Load Information in Debugger (DNS List Size, etc..)
LoadOptSettings
If frmSplash.Visible <> False Then
    frmSplash.lblState.Caption = "Loading Homepage..."
End If
GoHome
chkStart = 1
If frmSplash.Visible <> False Then
    frmSplash.lblState.Caption = "Initializing Browser..."
End If

Dim i As Long
For i = 0 To wb.UBound
    wb(i).Silent = True
Next i
setChecks
getFAV


    
GetWHistory
LoadFilterList
LoadFilterSettings
setDNS
If BoolAuto = True Then
    CheckDayToUpdate
Else
End If

If ValBKUP = 1 Then
    CheckDayToBackup
Else
End If


If chkTOTD = 1 Then
    frmTip.Show vbModal
End If

ColURL.AddItem curWB & "[]" & cboURL.Tag
End Sub
Private Function setChecks()
bLink = False
skipResolveBadNav = 0
ResolveAdd = 1
skipNavigateChk = 0
TempAllow = 0
chkNav = 0
DropDown = 0
chkGoBack = 0
chkGoForward = 0
chkLoop = 0
IndexPages = 0
OnPage = 0
DuplicateEntry = 0
End Function
Private Function GoHome()
Dim i As Long
For i = 1 To Len(StrHP)
    If LCase(InStr(i, StrHP, "dnsbrowser")) Then
        wb(curWB).Navigate App.Path & "\" & "DNSBrowser.html"
        cboURL.Text = "http://www.dnsbrowser.com"
        cboURL.Tag = cboURL.Text
        Exit Function
    End If
Next i

If StrHP$ = "about:blank" Then
    wb(curWB).Navigate "about:blank"
    cboURL.Text = "about:blank"
    cboURL.Tag = cboURL.Text
    Exit Function
End If
If BoolDNS = True Then
    chkNav = 0
    txtDebug.Text = txtDebug.Text & vbNewLine & "Designated URL: " & StrHP$
    breakURL$ = Replace(StrHP, "http://", "")
    ' NOTE, use FUNCTIONS to perform tasks; load is otherwise too large to work in control
    
    If Val(Left(breakURL$, 2)) > 9 Then
        wb(curWB).Navigate breakURL$
        cboURL.Tag = breakURL$
        txtDebug.Text = txtDebug.Text & vbNewLine & "Target URL: " & breakURL$
        chkNav = 1
        skipNavigateChk = 1
        Exit Function
    End If
    
    checkDNS breakURL$, "1"
    
    If chkNav <> 1 Then
        txtDebug.Text = txtDebug.Text & vbNewLine & "Specified URL not found in Database. Resolving..."
                StatBar.Panels(1).Text = "Specified URL not found in Database. Resolving..."
        If ValDBS = 3 Then
                If MsgBox("Would you like to Resolve/Add this website to the Database?" & vbNewLine & _
                        "URL: '" & breakURL$ & "'", vbYesNo, "Resolve/Add to Database?") = vbYes Then
                    ResolveAdd = 1
                Else
                    ResolveAdd = 0
                End If
            resolveDNS breakURL$
        Else
            If ValDBS = 1 Then
                ResolveAdd = 1
                resolveDNS breakURL$
            End If
            
            If ValDBS = 2 Then
                ResolveAdd = 0
                resolveDNS breakURL$
            End If
        End If
    End If
Else
    wb(curWB).Navigate StrHP$
    If Left(cboURL.Text, 7) <> "http://" Then
        cboURL.Text = "http://" & StrHP$
    Else
        cboURL.Text = StrHP$
    End If
End If
txtDebug.Text = txtDebug.Text & vbNewLine
End Function
Private Function SaveWHistory()
Dim HistSave$, i As Long
FileNumber = FreeFile
chkFExists "DATA\history.dat", "//DNS Browser - History//"

Open "DATA\history.dat" For Output As #FileNumber
    For i = 0 To cboURL.ListCount - 1
        Print #FileNumber, cboURL.List(i)
    Next i
Close #FileNumber

End Function
Private Function getFAV()
Dim i As Long
FileNumber = FreeFile
If frmSplash.Visible <> False Then
    frmSplash.lblState.Caption = "Loading Favorites..."
End If
Dim Title As String, URLf$, Folder$, TempStr As String
i = 0

lstFav.Path = "Favorites"

For i = 0 To lstFav.ListCount - 1
    Open lstFav.Path & "\" & lstFav.List(i) For Input As #FileNumber
        Do While Not EOF(FileNumber)
            Input #FileNumber, TempStr$
            If Left(TempStr$, 7) = "FOLDER=" Then Folder$ = Right(TempStr$, Len(TempStr$) - 7)
            If Left(TempStr$, 6) = "TITLE=" Then Title$ = Right(TempStr$, Len(TempStr$) - 6)
            If Left(TempStr$, 4) = "URL=" Then URLf$ = Right(TempStr$, Len(TempStr$) - 4)
            DoEvents
        Loop
    Close #FileNumber
    
    If i <> 0 Then Load mnuFavoriteT(i)
    
        If Len(Title$) <= 40 Then 'check length of caption for favorite, if greater than 40 then
            mnuFavoriteT(i).Caption = Title$ 'take the first 37 and add ...
        Else
            mnuFavoriteT(i).Caption = Left(Title$, 37) & "..."
        End If
        
    mnuFavoriteT(i).Tag = URLf$
    DoEvents
Next i

If frmSplash.Visible <> False Then
    frmSplash.lblState.Caption = "Favorites Loaded."
End If
End Function


Private Sub Form_Resize()

Dim SkipResize As Byte, i As Long, j As Long
SkipResize = 0
    
If Me.WindowState <> vbMinimized Then
If Me.Width <= 12465 Then
    Me.Width = 12465
End If

If Me.Height <= 7000 Then
    Me.Height = 7000
End If
Else
Exit Sub
End If

If Me.WindowState <> vbMinimized Then
    cboURL.Width = Me.Width - cmdGo.Width - lblAddress.Width - 100
    cmdGo.Left = Me.Width - cmdGo.Width - 100
    For i = 0 To wb.UBound
        For j = 0 To lstUnloaded.ListCount - 1
            If i = lstUnloaded.List(j) Then
                SkipResize = 1
            End If
        Next j
        
        If SkipResize <> 1 Then
            wb(i).Height = Me.Height - 800 - wb(i).Top - StatBar.Height
            SkipResize = 0
        End If
    Next i
    txtDebug.Height = Me.Height - txtDebug.Top - 800 - StatBar.Height
    txtDebug.Left = Me.Width - txtDebug.Width - 100
    StatBar.Panels.Item(1).Width = Me.Width - StatBar.Panels.Item(2).Width - StatBar.Panels.Item(3).Width - StatBar.Panels.Item(4).Width
    
    AddProgBar pgBar, StatBar, 2
 '   pgBar.Left = StatBar.Panels.Item(2).Left + 30
 '   pgBar.Width = StatBar.Panels.Item(2).Width - 60
 '   pgBar.Top = Me.ScaleHeight - pgBar.Height - 60
End If

If txtDebug.Visible = True Then
    For i = 0 To wb.UBound
        For j = 0 To lstUnloaded.ListCount - 1
            If i = lstUnloaded.List(j) Then
                SkipResize = 1
            End If
        Next j
        
        If SkipResize <> 1 Then
            wb(i).Width = txtDebug.Left
            SkipResize = 0
        End If
    Next i
Else
    For i = 0 To wb.UBound
        wb(i).Width = Me.Width - 100
    Next i
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveWHistory
CloseApplication
End Sub


Private Sub lblClose_Click()
PopUpResize 0
End Sub

Private Sub mnuAddFav_Click()
Dim i As Long
FileNumber = FreeFile
Dim CreateShortcut, Folder As String
Set CreateShortcut = CreateObject("Scripting.FileSystemObject")
Dim strSaveLoc As String, invalidChar As Byte, savestr As String
invalidChar = 1
If Len(wb(curWB).Document.Title) <= 40 Then
    strSaveLoc$ = wb(curWB).Document.Title 'prevent title from showing up too long to hint user
Else
    strSaveLoc$ = Left(wb(curWB).Document.Title, 37) & "..."
End If

savestr$ = InputBox("DNS Browser will add this url to your favorites: '" & wb(curWB).LocationURL & "'" & vbNewLine & "Please designate title.", "Add To Favorites?", strSaveLoc$)

Do Until invalidChar = 0 'make sure no invalid characters for file saving!!

        If InStr(savestr$, "\") = 1 Or InStr(savestr$, "/") = 1 Or InStr(savestr$, ":") = 1 _
        Or InStr(savestr$, "*") = 1 Or InStr(savestr$, "?") = 1 Or InStr(savestr$, """") = 1 _
        Or InStr(savestr$, "<") = 1 Or InStr(savestr$, ">") = 1 Or InStr(savestr$, "|") = 1 Then
            invalidChar = 1
        Else
            invalidChar = 0
        End If
        If invalidChar <> 0 Then 'if invalid character is true then prompt for rename
            savestr$ = InputBox("Favorites cannot contain the following characters: " & vbNewLine & _
            "\ / * ? " & """" & " < > | " & vbNewLine & "Please designate new title.", "Please rename", savestr$)
        End If
    DoEvents
Loop

If Len(savestr$) > 40 Then 'crop savestring length to prevent extensive lengths
    savestr$ = Left(savestr$, 37) & "..."
End If

For i = 0 To mnuFavoriteT.UBound - 1 'if name is already there then add a (2)
    If savestr$ = mnuFavoriteT.Item(i).Caption Then
        savestr$ = savestr$ & " (2)"
    End If
    DoEvents
Next i

If savestr$ <> vbNullString Then
    Dim SaveLoc As String
    SaveLoc$ = "Favorites\" & savestr$
    If CreateShortcut.FileExists(SaveLoc$ & ".fav") = False Then 'if not exist then create a new one
        CreateShortcut.CreateTextFile SaveLoc$ & ".fav"
    Else
    
    Do Until CreateShortcut.FileExists(SaveLoc$ & ".fav") = False 'dont want to overwrite
        SaveLoc$ = SaveLoc$ & "(2)"
        savestr$ = savestr$ & "(2)"
        DoEvents
    Loop
    
    CreateShortcut.CreateTextFile SaveLoc$ & ".fav"
    End If
    
    Open SaveLoc$ & ".fav" For Output As #FileNumber 'save information to the .fav item
        Print #FileNumber, "FOLDER=" & Folder$ & vbNewLine & _
                  "TITLE=" & savestr$ & vbNewLine & _
                  "URL=" & wb(curWB).LocationURL
    Close #FileNumber
    
    Load mnuFavoriteT(mnuFavoriteT.UBound + 1) 'load the favorite to the menu
    If Len(savestr$) <= 40 Then 'makes sure the save string isnt too large...
        mnuFavoriteT(mnuFavoriteT.UBound).Caption = savestr$
    Else 'otherwise make max 37 and add "..."
        mnuFavoriteT(mnuFavoriteT.UBound).Caption = Left(savestr$, 37) & "..."
    End If
    mnuFavoriteT(mnuFavoriteT.UBound).Tag = wb(curWB).LocationURL
End If

End Sub
Private Sub mnuAddressBar_Click()
txtDebug.Text = txtDebug.Text & vbNewLine & "Feature currently not yet implemented."
End Sub

Private Sub mnuAdminHide_Click()
If sldTransparency.Value <> 0 Then
    sldTransparency.Value = 0
Else
    sldTransparency.Value = 100
End If

App.TaskVisible = Not App.TaskVisible
End Sub

Private Sub mnuClearConsole_Click()
txtDebug.Text = "--------------------------- DNS Browser Debugger ---------------------------"

End Sub

Private Sub mnuClose_Click()
Dim i As Long

'=============================================
'LONG PROCEDURE FOR CLOSING... ON CLOSE SET A
'WINDOW TO OPEN, IF ORIGINAL WINDOW THEN PROMPT
'TO CLOSE. UPDATE EVERYTHING AT EACH CLOSE
'=============================================

If curWB <> 0 Then 'dont want to close current window, cant unload original
    WindowCount = WindowCount - 1
    Unload wb(curWB) 'if not original, unload the current
    Unload mnuWindowsWB(curWB) 'unload the menu item too
    For i = 0 To ColURL.ListCount 'collection of urls binded to menu = ColURL
        Dim ParseIndexUrl() As String 'parse index url, used to split the index (window) & url
        If ColURL.List(i) <> vbNullString Then 'if not empty then
            ParseIndexUrl$ = Split(ColURL.List(i), "[]", -1, 1) 'split it
                If ParseIndexUrl(0) = curWB Then 'the (0) means its the index bit of the parse
                    ColURL.RemoveItem i 'remove item index
                    Exit For
                End If
        End If
        DoEvents
    Next i
    lstUnloaded.AddItem curWB 'add item to unloaded list to allow windows to take its place
    mnuWindows.Caption = "Windows - " & "[" & WindowCount & "]" 'update menu's window count
    If curWB <> wb.Count Then 'if unloaded is not the max then...
        curWB = curWB + 1 'make current webbrowser the next one up
        For i = 0 To lstUnloaded.ListCount - 1
            If curWB = lstUnloaded.List(i) Then
                curWB = curWB + 1
                i = 0
            End If
        Next i
        
        With wb(curWB) 'set properties
            .Left = wb(0).Left
            .Top = wb(0).Top
            .Width = wb(0).Width
            .Height = wb(0).Height
            .Visible = True
            .SetFocus
            .ZOrder
        End With
        Me.Caption = wb(curWB).Document.Title & "- DNS Browser" 'update caption to document title
        
    For i = 0 To ColURL.ListCount 'collection of urls binded to menu = ColURL
        If ColURL.List(i) <> vbNullString Then 'if not empty then
            ParseIndexUrl$ = Split(ColURL.List(i), "[]", -1, 1) 'split it
                If ParseIndexUrl(0) = curWB Then 'the (0) means its the index bit of the parse
                    cboURL.Text = ParseIndexUrl(1)
                    Exit For 'update it to current
                End If
        End If
        DoEvents
    Next i
    Else
        curWB = curWB - 1 'otherwise make current window one less because cant load past MAX DUHH
        
        For i = 0 To lstUnloaded.ListCount - 1
            If curWB = lstUnloaded.List(i) Then
                curWB = curWB - 1
                i = 0
            End If
        Next i
        
        With wb(curWB) 'set properties
            .Left = wb(0).Left
            .Top = wb(0).Top
            .Width = wb(0).Width
            .Height = wb(0).Height
            .Visible = True
            .SetFocus
            .ZOrder
        End With
        Me.Caption = wb(curWB).Document.Title & "- DNS Browser" 'update caption
        
    For i = 0 To ColURL.ListCount 'collection of urls binded to menu = ColURL
        If ColURL.List(i) <> vbNullString Then 'if not empty then
            ParseIndexUrl$ = Split(ColURL.List(i), "[]", -1, 1) 'split it
                If ParseIndexUrl(0) = curWB Then 'the (0) means its the index bit of the parse
                    cboURL.Text = ParseIndexUrl(1)
                    Exit For 'update it to current
                End If
        End If
        DoEvents
    Next i
    End If
Else 'if it IS the current one prompt message for quiting app
    If MsgBox("Are you sure you want to exit the program?", vbYesNo + vbCritical, "Exit DNS Browser?") = vbYes Then
        CloseApplication 'close app function
    Else
    End If
End If
If picSearch.Visible = True Then
    picSearch.ZOrder
End If
End Sub
Private Sub mnuDebugConsole_Click()
Dim SkipResize As Byte
SkipResize = 0
'menu item to show/hide console, self-explanatory
If mnuDebugConsole.Checked = True Then
    mnuDebugConsole.Checked = False
    txtDebug.Visible = False
    wb(curWB).Width = Me.Width - 100 'resize this first before others to prevent visible lag time
    Dim i As Long, j As Long
    For i = 0 To wb.UBound
        For j = 0 To lstUnloaded.ListCount - 1
            If i = lstUnloaded.List(j) Then 'resizing loop check to make sure doesn't
                SkipResize = 1 'error when there is an unloaded
            End If
        Next j
        If SkipResize <> 1 Then
            wb(i).Width = Me.Width - 100
        End If
        
        SkipResize = 0
        DoEvents
    Next i
            
Else
    wb(curWB).Width = Me.Width - txtDebug.Width - 100 'resize this first before others to prevent visible lag time
    
    mnuDebugConsole.Checked = True
    txtDebug.Visible = True
    SkipResize = 0
    For i = 0 To wb.UBound 'good method for many uses
        For j = 0 To lstUnloaded.ListCount - 1
            If i = lstUnloaded.List(j) Then
                SkipResize = 1
            End If
        Next j
        
        If SkipResize <> 1 Then
            wb(i).Width = Me.Width - txtDebug.Width - 100
        End If
        SkipResize = 0
        DoEvents
    Next i
End If

End Sub

Private Sub mnuDebugCopy_Click()
Clipboard.SetText txtDebug.SelText
End Sub

Private Sub mnuDebugPaste_Click()
txtDebug.SelText = Clipboard.GetText
End Sub

Private Sub mnuDNS_Click()
frmDNS.Show 'show the form captain!
End Sub
Private Sub mnuEditCopy_Click()
wb(curWB).ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT 'COPY ME
End Sub
Private Sub mnuEditCut_Click()
wb(curWB).ExecWB OLECMDID_CUT, OLECMDEXECOPT_DODEFAULT
End Sub
Private Sub mnuEditFind_Click()
wb(curWB).SetFocus 'set focus browser
SendKeys "^f", True 'initiate keys to FIND!
End Sub
Private Sub mnuEditPaste_Click()
wb(curWB).ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DODEFAULT 'paste me momma!
End Sub
Private Sub mnuFavoriteT_Click(Index As Integer)
'ok click on a favorite then prepare to navigate!
txtDebug.Text = txtDebug.Text & vbNewLine & "Designated URL: " & mnuFavoriteT(Index).Tag
breakURL$ = Replace(mnuFavoriteT(Index).Tag, "http://", "")
checkDNS breakURL$, "2"

End Sub

Private Sub mnuFileMagic_Click()
frmFileMagic.Show
End Sub

Private Sub mnuFileNewWindow_Click()

WindowCount = WindowCount + 1 'add to window count
'=================================================
'PROCEDURE TO LOAD NEW WINDOW; CHECK FOR OPEN SPOTS
'IF NONE, THEN OPEN NEW ONE AND UPDATE EVERYTHING
'=================================================
If lstUnloaded.ListCount = 0 Then 'if there isnt any unloaded spots waiting then load a new spot
    Load mnuWindowsWB(wb.UBound + 1) 'load me!
    Load wb(wb.UBound + 1) 'yay loading
    ColURL.AddItem wb.UBound & "[]" & cboURL.Tag 'add a url for current wb in list
Else
    Load mnuWindowsWB(lstUnloaded.List(0)) 'this means there IS a spot to be filled
    Load wb(lstUnloaded.List(0)) 'go ahead and fill it
    ColURL.AddItem lstUnloaded.List(0) & "[]" & cboURL.Tag
    lstUnloaded.RemoveItem (0)
End If

mnuWindows.Caption = "Windows - " & "[" & WindowCount & "]" 'update window count
With wb(wb.UBound) 'update information
    .Left = wb(0).Left
    .Top = wb(0).Top
    .Width = wb(0).Width
    .Height = wb(0).Height
    .Visible = True
    .SetFocus
    .ZOrder
End With

curWB = wb.UBound

If picSearch.Visible = True Then
    picSearch.ZOrder
End If

If ProcDispN <> 1 Then
    GoHome
Else
        If BoolDNS = True Then
            If InStr(1, strNewWindow$, "http://") Then
                strNewWindow$ = Replace(strNewWindow$, "http://", vbNullString)
            End If
            chkNav = 0
            checkDNS strNewWindow$, "1"
            
    If chkNav <> 1 Then
        txtDebug.Text = txtDebug.Text & vbNewLine & "Specified URL not found in Database. Resolving..."
                StatBar.Panels(1).Text = "Specified URL not found in Database. Resolving..."
        If ValDBS = 3 Then
                If MsgBox("Would you like to Resolve/Add this website to the Database?" & vbNewLine & _
                        "URL: '" & strNewWindow$ & "'", vbYesNo, "Resolve/Add to Database?") = vbYes Then
                    ResolveAdd = 1
                Else
                    ResolveAdd = 0
                End If
            resolveDNS strNewWindow$
        Else
            If ValDBS = 1 Then
                ResolveAdd = 1
                resolveDNS strNewWindow$
            End If
            
            If ValDBS = 2 Then
                ResolveAdd = 0
                resolveDNS strNewWindow$
            End If
        End If
    End If
        
        Else
            wb(curWB).Navigate strNewWindow$
        End If
    ProcDispN = 0
    Exit Sub
End If

Dim SkipVisible As Byte, i As Long, j As Long
SkipVisible = 0
For i = 0 To wb.UBound - 1
    For j = 0 To lstUnloaded.ListCount - 1
        If i = lstUnloaded.List(j) Then
            SkipVisible = 1
        End If
    Next j
        
    If SkipVisible <> 1 Then
        wb(i).Visible = False
    End If
    SkipVisible = 0
    DoEvents
Next i


End Sub
Private Sub mnuFileOpen_Click()
wb(curWB).ExecWB OLECMDID_OPEN, OLECMDEXECOPT_DODEFAULT 'open me momma! (currently dont work!!)
End Sub

Private Sub mnuFilePageSetup_Click()
On Error Resume Next
wb(curWB).ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT 'PAGE ME BABY
End Sub

Private Sub mnuFilePrint_Click()
wb(curWB).ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT 'print yeah!
End Sub
Private Sub mnuFilePrintPrev_Click()
wb(curWB).ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT 'preview
End Sub
Private Sub mnuFileProperties_Click()
wb(curWB).ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_DODEFAULT 'file properties
End Sub
Private Sub mnuFileSave_Click()
wb(curWB).ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT 'save me or die!! ARRR
End Sub

Private Sub mnuFlashLight_Click()
If frmHEHE.Visible = True Then
    frmHEHE.Visible = False
Else
    sldTransparency.Value = 2
    frmBrowser.WindowState = vbMaximized
    frmHEHE.Show
End If
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuHelpTOTD_Click()
frmTip.Show
End Sub

Private Sub mnuHideSearch_Click()
picSearch.Visible = False
End Sub

Private Sub mnuInternetOptions_Click()
frmOptions.Show vbModal
End Sub

Private Sub mnuLockWorkstation_Click()
LockWorkStation
End Sub

Private Sub mnuOrganizeFav_Click()
frmFavorites.Show vbModal
End Sub

Private Sub mnuPopUpAddList_Click()
Dim strAddURL As String, strSplitURLt() As String
If Left(wb(curWB).LocationURL, 7) = "http://" Then
    strAddURL$ = Replace(wb(curWB).LocationURL, "http://", vbNullString)
Else
    strAddURL$ = wb(curWB).LocationURL
End If
strSplitURLt = Split(strAddURL$, "/", -1, 1)
frmOptions.lstPB.AddItem strSplitURLt(0)
PopUpResize 0
End Sub

Private Sub mnuPopupBlockSet_Click()
'popup blocker settings
If mnuPopupBlockSet.Caption = "Turn Off Pop-up Blocker" Then
    mnuPopupBlockSet.Caption = "Turn On Pop-up Blocker"
    StatBar.Panels.Item(1).Text = "Pop-up Blocker disabled."
    frmOptions.chkEnable.Value = 0
    WriteString "Privacy Settings", "BLOCKPOPUP", "0", "DATA\options.dat"
Else
    mnuPopupBlockSet.Caption = "Turn Off Pop-up Blocker"
    StatBar.Panels.Item(1).Text = "Pop-up Blocker enabled."
    frmOptions.chkEnable.Value = 1
    WriteString "Privacy Settings", "BLOCKPOPUP", "1", "DATA\options.dat"
End If
End Sub

Private Sub mnuPopupDisableSnd_Click()
frmOptions.chkPopSnd.Value = 0
WriteString "Privacy Settings", "PLAYSOUND", "0", "DATA\options.dat"
End Sub

Private Sub mnuPopUpNotification_Click()
frmOptions.chkPopNotify.Value = 0
WriteString "Privacy Settings", "NOTIFY", "0", "DATA\options.dat"
End Sub

Private Sub mnuPopupTurnOff_Click()
frmOptions.chkEnable.Value = 0
mnuPopupBlockSet.Caption = "Turn On Pop-up Blocker"
WriteString "Privacy Settings", "BLOCKPOPUP", "0", "DATA\options.dat"
End Sub

Private Sub mnuQDebug_Click()
chkStart = 1
skipNavigateChk = 1
chkNav = 1
End Sub
Private Sub mnuRefresh_Click()
wb(curWB).Refresh
End Sub
Private Sub mnuRefreshDNS_Click()
refreshDNS
End Sub
Private Sub mnuSelectAll_Click()
wb(curWB).ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT 'select all
End Sub
Private Sub mnuStandardButtons_Click()
'=============================================
'PROCEDURES; All just resizing/placing things
'where they belong... arrr takes awhile
'=============================================
Dim SkipResize As Byte
SkipResize = 0
If mnuStandardButtons.Checked = True Then
    mnuStandardButtons.Checked = False
    ToolBar.Visible = False
    cboURL.Top = 0
    lblAddress.Top = cboURL.Top + 45
    cmdGo.Top = cboURL.Top
    wb(curWB).Top = cboURL.Top + cboURL.Height
    wb(curWB).Height = wb(curWB).Height + ToolBar.Height
    Dim i As Long, j As Long
    For i = 0 To wb.UBound
        For j = 0 To lstUnloaded.ListCount - 1
            If i = lstUnloaded.List(j) Then 'resizing loop check to make sure doesn't
                SkipResize = 1 'error when there is an unloaded
            End If
        Next j
        If SkipResize <> 1 Then
            wb(i).Top = cboURL.Top + cboURL.Height
            wb(i).Height = wb(curWB).Height + ToolBar.Height
        End If
        
        SkipResize = 0
        DoEvents
    Next i
    
    txtDebug.Top = wb(curWB).Top
    txtDebug.Height = wb(curWB).Height
Else
    mnuStandardButtons.Checked = True
    ToolBar.Visible = True
    cboURL.Top = ToolBar.Height
    cmdGo.Top = cboURL.Top
    lblAddress.Top = cboURL.Top + 45
    wb(curWB).Top = cboURL.Top + cboURL.Height
    wb(curWB).Height = wb(curWB).Height - ToolBar.Height
    
    For i = 0 To wb.UBound
        For j = 0 To lstUnloaded.ListCount - 1
            If i = lstUnloaded.List(j) Then 'resizing loop check to make sure doesn't
                SkipResize = 1 'error when there is an unloaded
            End If
        Next j
        If SkipResize <> 1 Then
            wb(i).Top = cboURL.Top + cboURL.Height
            wb(i).Height = wb(curWB).Height + ToolBar.Height
        End If
        
        SkipResize = 0
        DoEvents
    Next i
    txtDebug.Top = wb(curWB).Top
    txtDebug.Height = wb(curWB).Height
End If
End Sub

Private Sub mnuStayTop_Click()
mnuStayTop.Checked = Not mnuStayTop.Checked
WriteString "Miscellaneous", "STAYONTOP", mnuStayTop.Checked, "DATA\options.dat"
If mnuStayTop.Checked = True Then
    SetTopMostWindow Me.hwnd, True
Else
    SetTopMostWindow Me.hwnd, False
End If
End Sub

Private Sub mnuStop_Click()
wb(curWB).Stop
End Sub

Private Sub mnuTextSizeX_Click(Index As Integer)
FileNumber = FreeFile
'================================================
'Text size changing procedure. Make it identical
'to IE, woohoo!
'================================================
    Dim i As Long
    For i = 0 To 4
        mnuTextSizeX(i).Checked = False
    DoEvents
    Next i
    mnuTextSizeX(Index).Checked = True
    Dim wbTXTSize As Byte
    Select Case Index
        Case 0
        wbTXTSize = "4"
        
        Case 1
        wbTXTSize = "3"
        
        Case 2
        wbTXTSize = "2"
        
        Case 3
        wbTXTSize = "1"
        
        Case 4
        wbTXTSize = "0"
    End Select

    wb(curWB).ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DODEFAULT, CLng(wbTXTSize)
    Dim TempStr As String, InputTXTSize As String, OtherData As String
    On Error Resume Next
    Open "DATA\pref.dat" For Input As #FileNumber
        Do While Not EOF(FileNumber)
            Input #FileNumber, TempStr$
            If Left(TempStr$, 8) = "TEXTSIZE" Then
                InputTXTSize$ = TempStr$
            Else
                OtherData$ = OtherData$ & TempStr$
            End If
        DoEvents
        Loop
    Close #FileNumber
    
    InputTXTSize$ = Left(InputTXTSize$, Len(InputTXTSize$) - 1) & Index
    
    Open "DATA\pref.dat" For Output As #FileNumber
        Print #FileNumber, InputTXTSize$ & vbNewLine & OtherData$
    Close #FileNumber
    
End Sub


Private Sub mnuToolsFTP_Click()
frmFTP.Show
End Sub

Private Sub mnuToolUpdates_Click()
CheckForUpdates ("C:\Documents and Settings\Daniel\Desktop\DNS Project Package\DNSBrowser.exe")
End Sub

Private Sub mnuViewBW_Click()
wb(curWB).Document.createStyleSheet App.Path & "\DATA\Styles\" & "styleb.css"
End Sub

Private Sub mnuViewDebugTest_Click()

'just debug testting stuff, most of this is worthless or already used!

'possible contrast feature
'wb(curWB).Document.bgcolor = "#000000" 'set bg color
'wb(curWB).Document.fgcolor = "#FFFFFF" 'set text color
'wb(curWB).Document.linkcolor = "#FFFFFF" 'set link color
'wb(curWB).Document.alinkcolor = "#FFFFFF"
'wb(curWB).Document.vlinkColor = "#FFFFFF"
frmDebug.Show
'documentsize$ = wb(curWB).Document.fileSize 'gets document size
'wb(curwb).document.title

'wb(curWB).Document.referrer 'get prev page

'wb(curWB).Document.anchors 'get page anchors

'wb(curWB).Document.images 'gets pictures!

'get cookie stuff
'cookie$ = wb(curWB).Document.cookie
'MsgBox cookie$
End Sub

Private Sub mnuViewEmails_Click()
FileNumber = FreeFile
'=======================================================
'EMAIL LINK PROCEDURE; get number of links on page then
'check the link items for mailto: tag and if so, add it
'and open in notepad. if not, give a shout
'=======================================================

Dim TempStr As String
    Dim i As Long
    For i = 0 To wb(curWB).Document.links.Length - 1
        If TempStr$ = vbNullString Then
            If Left(wb(curWB).Document.links(i).href, 7) = "mailto:" Then
                TempStr$ = Right(wb(curWB).Document.links(i).href, Len(wb(curWB).Document.links(i).href) - 7) & vbNewLine
            End If
        Else
            If Left(wb(curWB).Document.links(i).href, 7) = "mailto:" Then
                TempStr$ = TempStr$ & Right(wb(curWB).Document.links(i).href, Len(wb(curWB).Document.links(i).href) - 7) & vbNewLine
            End If
        End If
        DoEvents
    Next i
If TempStr$ <> vbNullString Then
    Open "DATA\tmp_source.txt" For Output As #FileNumber
        Print #FileNumber, TempStr$
    Close #FileNumber
    
    Shell "notepad.exe DATA\tmp_source.txt", vbNormalFocus
Else
    MsgBox "No emails were found on this page!", vbExclamation, "Emails not found."
End If
End Sub

Private Sub mnuViewHC_Click()
wb(curWB).Document.createStyleSheet App.Path & "\DATA\Styles\" & "stylea.css"

End Sub

Private Sub mnuViewImgSrc_Click()
FileNumber = FreeFile
Dim i As Long, TempStr As String
For i = 0 To wb(curWB).Document.images.Length - 1
        If TempStr$ = vbNullString Then
                TempStr$ = wb(curWB).Document.images(i).src & vbNewLine
        Else
                TempStr$ = TempStr$ & wb(curWB).Document.images(i).src & vbNewLine
        End If
        DoEvents
    Next i
If TempStr$ <> vbNullString Then
    Open "DATA\tmp_source.txt" For Output As #FileNumber
        Print #FileNumber, TempStr$
    Close #FileNumber
    
    Shell "notepad.exe DATA\tmp_source.txt", vbNormalFocus
Else
    MsgBox "No images were found on this page!", vbExclamation, "Images not found."
End If
End Sub

Private Sub mnuViewLinks_Click()
Dim TempStr As String
    FileNumber = FreeFile
    Dim i As Long
    For i = 0 To wb(curWB).Document.links.Length - 1
        If TempStr$ = vbNullString Then
            TempStr$ = wb(curWB).Document.links(i).href & vbNewLine
        Else
            TempStr$ = TempStr$ & wb(curWB).Document.links(i).href & vbNewLine
        End If
        DoEvents
    Next i
    
If TempStr$ <> vbNullString Then
    Open "DATA\tmp_source.txt" For Output As #FileNumber
        Print #FileNumber, TempStr$
    Close #FileNumber
    
    Shell "notepad.exe DATA\tmp_source.txt", vbNormalFocus
Else
    MsgBox "No links were found on this page!", vbExclamation, "Links not found."
End If
End Sub

Private Sub mnuViewSource_Click()
FileNumber = FreeFile
Dim TempStr As String
TempStr$ = wb(curWB).Document.documentElement.innerHTML
TempStr$ = Replace(TempStr$, txtChar.Text, vbNewLine)

Open "DATA\tmp_source.txt" For Output As #FileNumber
    Print #FileNumber, TempStr$
Close #FileNumber

Shell "notepad.exe DATA\tmp_source.txt", vbNormalFocus
'frmSource.Show
End Sub


Private Sub mnuViewText_Click()
FileNumber = FreeFile
Dim TempStr As String
TempStr$ = wb(curWB).Document.documentElement.innerText
TempStr$ = Replace(TempStr$, txtChar.Text, vbNewLine)

Open "DATA\tmp_source.txt" For Output As #FileNumber
    Print #FileNumber, TempStr$
Close #FileNumber

Shell "notepad.exe DATA\tmp_source.txt", vbNormalFocus
End Sub

Private Sub mnuWatchMovie_Click()
Dim TempStr As String, SplitToMovie() As String, SplitPart2() As String
TempStr$ = wb(curWB).Document.documentElement.innerHTML
TempStr$ = Replace(TempStr$, txtChar.Text, vbNewLine)

SplitToMovie = Split(TempStr$, "<PARAM NAME=movie VALUE=" & """" & "http://uploads.ungrounded.net", -1, 1)
If UBound(SplitToMovie) >= 1 Then
    SplitPart2 = Split(SplitToMovie(1), """", -1, 1)
    TempStr$ = "uploads.ungrounded.net" & SplitPart2(0)
    checkDNS TempStr$, 1
Else
    MsgBox "No Movie could be found!", vbCritical, "Action Canceled"
End If
End Sub

Private Sub mnuWBBack_Click()
On Error Resume Next
wb(curWB).GoBack
End Sub

Private Sub mnuWBCopy_Click()
If bLink <> True Then
    Call mnuEditCopy_Click
Else
    Clipboard.SetText ClickURL$
End If
End Sub

Private Sub mnuWBCut_Click()
Call mnuEditCut_Click
End Sub

Private Sub mnuWBDebug_Click()
chkStart = 1
skipNavigateChk = 1
chkNav = 1
End Sub

Private Sub mnuWBFav_Click()
If bLink <> True Then
    Call mnuAddFav_Click
Else

  'Implement this
End If
End Sub

Private Sub mnuWBFind_Click()
Call mnuEditFind_Click
End Sub

Private Sub mnuWBForward_Click()
On Error Resume Next
wb(curWB).GoForward
End Sub

Private Sub mnuWBOpen_Click()
Call IEDoc_onclick
End Sub

Private Sub mnuWBOpenNew_Click()
ProcDispN = 1
strNewWindow$ = IEDoc.activeElement
MenuNewWindow = True
Call mnuFileNewWindow_Click
End Sub

Private Sub mnuWBPaste_Click()
Call mnuEditPaste_Click
End Sub

Private Sub mnuWBPrint_Click()
Call mnuFilePrint_Click
End Sub

Private Sub mnuWBProperties_Click()
Call mnuFileProperties_Click
End Sub

Private Sub mnuWBRefresh_Click()
wb(curWB).Refresh
End Sub

Private Sub mnuWBSelectAll_Click()
Call mnuSelectAll_Click
End Sub

Private Sub mnuWBViewEmails_Click()
Call mnuViewEmails_Click
End Sub

Private Sub mnuWBViewImages_Click()
Call mnuViewImgSrc_Click
End Sub

Private Sub mnuWBViewLinks_Click()
Call mnuViewLinks_Click
End Sub

Private Sub mnuWBViewSource_Click()
Call mnuViewSource_Click
End Sub

Private Sub mnuWBViewText_Click()
Call mnuViewText_Click
End Sub

Private Sub mnuWhois_Click()
frmWhois.Show
End Sub

Private Sub mnuWindows_Click()
Dim i As Long
Dim j As Long
Dim SkipTry As Byte
SkipTry = 0

For i = 0 To wb.UBound
    For j = 0 To lstUnloaded.ListCount - 1
        If i <> lstUnloaded.List(j) Then
        Else
        SkipTry = 1
        End If
    Next j
    
    If SkipTry <> 1 Then
        If wb(i).LocationName <> "DNSBrowser.html" Then
            If Len(wb(i).Document.Title) <= 40 Then
                If Len(wb(i).Document.Title) = 0 Then
                    mnuWindowsWB.Item(i).Caption = "Blank"
                Else
                    mnuWindowsWB.Item(i).Caption = wb(i).Document.Title
                End If
            Else
                mnuWindowsWB.Item(i).Caption = Left(wb(i).Document.Title, 37) & "..."
            End If
            mnuWindowsWB.Item(i).Checked = False
        Else
            mnuWindowsWB.Item(i).Caption = "DNS Browser Home Page"
            mnuWindowsWB.Item(i).Checked = False
        End If
    End If
    SkipTry = 0
Next i

If mnuWindowsWB.Count = 1 And mnuWindowsWB.Item(0).Checked = False Then
    mnuWindowsWB.Item(0).Checked = True
Else
    mnuWindowsWB.Item(curWB).Checked = True
End If
End Sub

Private Sub mnuWindowsWB_Click(Index As Integer)
StatBar.Panels.Item(1).Text = vbNullString

If wb(Index).Document.Title <> vbNullString Then
    Me.Caption = wb(Index).Document.Title & " - DNS Browser"
Else
    Me.Caption = "Hello there, " & strUserName$ & "! Welcome to DNS Browser - DNS Browser"
End If

Dim ParseIndexUrl() As String, DontAddMe As Byte
DontAddMe = 0
For i = 0 To ColURL.ListCount - 1
    ParseIndexUrl$ = Split(ColURL.List(i), "[]", -1, 1)
    If ParseIndexUrl(0) = curWB Then
        ColURL.List(i) = curWB & "[]" & cboURL.Tag
        'cboURL.Text = ParseIndexUrl(1)
        DontAddMe = 1
    End If
Next i

curWB = Index

With wb(Index)
    .Left = wb(0).Left
    .Top = wb(0).Top
    .Width = wb(0).Width
    .Height = wb(0).Height
    .Visible = True
    .ZOrder
End With

If DontAddMe <> 1 Then
    ColURL.AddItem curWB & "[]" & cboURL.Tag
End If

'=====================================================================================
'Below commented code slows down processing time a bit... need to work on that!
'=====================================================================================
'Dim SkipVisible As Byte, j As Long
'SkipVisible = 0
'For i = 0 To wb.UBound
'    For j = 0 To lstUnloaded.ListCount
'        If i = lstUnloaded.List(j) Then
'            SkipVisible = 1
'        End If
'    Next j
'
'    If SkipVisible <> 1 Then
'        If i <> curWB Then wb(i).Visible = False
'    End If
'    SkipVisible = 0
'Next i

For i = 0 To ColURL.ListCount - 1
    ParseIndexUrl$ = Split(ColURL.List(i), "[]", -1, 1)
    If ParseIndexUrl(0) = Index Then
        cboURL.Text = ParseIndexUrl(1)
        cboURL.Tag = ParseIndexUrl(1)
        Exit For
    End If
    DoEvents
Next i

End Sub
Private Sub picPopup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnuPopup
End If
End Sub





Private Sub picSearch_DblClick()
PopupMenu mnuSearch
End Sub

Private Sub sldTransparency_Change()
If sldTransparency.Value <> 100 Then
    objAlpha.SetLayered Me.hwnd, True, CByte((sldTransparency.Value * 2.5))
Else
    objAlpha.SetLayered Me.hwnd, True, 255
End If
End Sub



Private Sub ToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button
    Case "Back"
    On Error Resume Next
    'OnPage = OnPage - 1
    wb(curWB).GoBack
    
    'If OnPage = 0 Then
    '    ToolBar.Buttons.Item(1).Enabled = False
    'End If
    
    'If OnPage <> IndexPages Then
    '    ToolBar.Buttons.Item(2).Enabled = True
    'End If
    chkGoBack = 1

    Case "Forward"
    On Error Resume Next
    'OnPage = OnPage + 1
    wb(curWB).GoForward
    'If OnPage = IndexPages Then
    '    ToolBar.Buttons.Item(2).Enabled = False
    'End If
    
    'If OnPage <> 0 Then
    '    ToolBar.Buttons.Item(1).Enabled = True
    'End If
    chkGoForward = 1
    
    Case "Stop"
    wb(curWB).Stop
    
    Case "Refresh"
    wb(curWB).Refresh
    
    Case "Home"
    GoHome
    
    Case "Search"
        picSearch.Visible = Not picSearch.Visible
        If picSearch.Visible Then
            txtSearch.SetFocus
        End If
    
    Case "Favorites"
    Call mnuOrganizeFav_Click
    
    Case "History"
    ' get history
    
    Case "Debugger"
    Call mnuDebugConsole_Click
    
End Select

'Exit Sub
'DisableBack:
'ToolBar.Buttons.Item(1).Enabled = False
'Exit Sub
'DisableForward:
'ToolBar.Buttons.Item(2).Enabled = False
End Sub
Private Sub txtDebug_Change()
txtDebug.SelStart = Len(txtDebug.Text)
End Sub

Private Sub txtDebug_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
txtDebug.Enabled = False
txtDebug.Enabled = True
txtDebug.SetFocus
PopupMenu mnuDebugPopup
End If
End Sub
Private Function IEDoc_oncontextmenu() As Boolean
   IEDoc_oncontextmenu = False
    If IEDoc.activeElement <> "[object]" Then
        mnuWBOpen.Visible = True
        mnuWBOpenNew.Visible = True
        mnuWBBack.Visible = False
        mnuWBForward.Visible = False
        mnuWBFav.Visible = False
        bLink = True
    Else
        mnuWBOpen.Visible = False
        mnuWBOpenNew.Visible = False
        mnuWBBack.Visible = True
        mnuWBForward.Visible = True
        mnuWBFav.Visible = True
        bLink = False
    End If
   PopupMenu mnuWBMenu
End Function

Private Function IEDoc_onclick() As Boolean


Dim strCheckCom As String
ClickURL$ = IEDoc.activeElement
If ClickURL$ = "[object]" Then
    IEDoc_onclick = True
    Exit Function
End If
NavigateAlreadyChecked = True
If Right(ClickURL$, 1) = "/" Then
strCheckCom$ = LCase(Mid(ClickURL, Len(ClickURL$) - 4, 4))
    If strCheckCom$ = ".com" Or strCheckCom$ = ".net" Or strCheckCom$ = ".org" Or strCheckCom$ = ".edu" Or strCheckCom$ = ".gov" Then
        ClickURL$ = Left(ClickURL$, Len(ClickURL$) - 1)
        Else
        strCheckCom$ = LCase(Mid(ClickURL, Len(ClickURL$) - 4, 3))
            If strCheckCom$ = ".tv" Or strCheckCom$ = ".nu" Or strCheckCom$ = ".au" Or strCheckCom$ = ".to" Or strCheckCom$ = ".co" Then
                ClickURL$ = Left(ClickURL$, Len(ClickURL$) - 1)
            End If
    End If
End If

IEDoc_onclick = False

FilterCheck (ClickURL$)

If xCancel = 1 Then
    xCancel = 0
    Exit Function
End If

Dim i As Long

If frmOptions.OptIB.Item(0).Value = True Then
    chkNav = 0
    txtDebug.Text = txtDebug.Text & vbNewLine & "Designated URL: " & ClickURL$
    ClickURL$ = Replace(ClickURL$, "http://", "")
    ' NOTE, use FUNCTIONS to perform tasks; load is otherwise too large to work in control
    
    If Val(Left(ClickURL$, 2)) > 9 Then
        'wb(curWB).Tag = breakURL$
        wb(curWB).Navigate ClickURL$
        txtDebug.Text = txtDebug.Text & vbNewLine & "Target URL: " & ClickURL$
        chkNav = 1
        skipNavigateChk = 1
        Exit Function
    End If
    
    checkDNS ClickURL$, "1"
    If chkNav <> 1 Then
        txtDebug.Text = txtDebug.Text & vbNewLine & "Specified URL not found in Database. Resolving..."
                StatBar.Panels(1).Text = "Specified URL not found in Database. Resolving..."
        If ValDBS = 3 Then
                If MsgBox("Would you like to Resolve/Add this website to the Database?" & vbNewLine & _
                        "URL: '" & ClickURL$ & "'", vbYesNo, "Resolve/Add to Database?") = vbYes Then
                    ResolveAdd = 1
                Else
                    ResolveAdd = 0
                End If
            resolveDNS ClickURL$
        Else
            If ValDBS = 1 Then
                ResolveAdd = 1
                resolveDNS ClickURL$
            End If
            
            If ValDBS = 2 Then
                ResolveAdd = 0
                resolveDNS ClickURL$
            End If
        End If
        
    End If
    txtDebug.Text = txtDebug.Text & vbNewLine
Else
    wb(curWB).Navigate ClickURL$
    '
End If

End Function
Private Sub wb_BeforeNavigate2(Index As Integer, ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
If picPopup.Visible = True Then
    PopUpResize 0
End If

If NavigateAlreadyChecked = True Then
    NavigateAlreadyChecked = False
    Exit Sub
End If

FilterCheck (URL)

If xCancel = 1 Then
    Cancel = True
    xCancel = 1
    Exit Sub
End If

If frmOptions.OptIB.Item(0).Value = True Then
    If skipNavigateChk = 1 Then
        skipNavigateChk = 0
        Exit Sub
    End If
    
    If URL = App.Path & "\" & "DNSBrowser.html" Then
        cboURL.Text = "http://www.dnsbrowser.com"
        cboURL.Tag = "http://www.dnsbrowser.com"
        Exit Sub
    End If
    
    If URL = "about:blank" Then
        Exit Sub
    End If
    'Cancel = True
    Dim i As Long
    

        If InStr(URL, "/search?hl=en&lr=&q=related:") = 1 Then
            Dim strParseSearchURL$
            strParseSearchURL$ = Right(URL, Len(URL) - i - 27)
            checkDNS strParseSearchURL$, "2"
            Exit Sub
        End If

    
    checkDNS URL, "1"

Else
End If
End Sub
Private Sub wb_DocumentComplete(Index As Integer, ByVal pDisp As Object, URL As Variant)
If URL <> "about:blank" Then
    Set IEDoc = wb(curWB).Document

'=============================================
'Description: The below commented code blocks
'advertisement banners and text. Remove comment
'to enable it. I have a newer source code at
'my college that included this feature in options
'but I will have to wait till I get off break.
'=============================================
   ' If initLoad <> 1 Then
   '     Dim i As Long
   '     For i = 0 To IEDoc.frames.Length - 1
   '         On Error Resume Next
   '         IEDoc.frames(i).window.location.href = ""
   '     Next i
   '         BlockAdvert
   '         Exit Sub
   '     Else
   '         initLoad = 0
   ' End If
    StatBar.Panels.Item(1).Text = "Done"
    
    mnuWindowsWB(Index).Caption = wb(curWB).LocationName
    
    If wb(curWB).LocationName <> "DNSBrowser.html" Then
        If wb(curWB).LocationURL <> "about:blank" Then
            Me.Caption = wb(curWB).LocationName & " - " & "DNS Browser"
        Else
            Me.Caption = "Hello there, " & strUserName$ & "! Welcome to DNS Browser - DNS Browser"
        End If
    Else
        Me.Caption = wb(curWB).Document.Title
    End If
    
    If chkStart = 1 Then
        chkStart = 0
        Exit Sub
        Else
    End If
    
    If chkGoBack = 1 Then
        chkGoBack = 0
    Else
        If chkGoForward = 1 Then
            chkGoForward = 0
        Else
            IndexPages = IndexPages + 1
            OnPage = OnPage + 1
            ToolBar.Buttons.Item(1).Enabled = True
        End If
    End If
    
    
    If URL = App.Path & "\" & "DNSBrowser.html" Then
        cboURL.Text = "http://www.dnsbrowser.com"
        cboURL.Tag = "http://www.dnsbrowser.com"
        Exit Sub
    End If
    If frmOptions.OptIB.Item(0).Value = True Then
        breakURL$ = Replace(URL, "http://", "")
        
        checkDNS breakURL$, 3
        
        If URL <> "http:///" Then
            txtDebug.Text = txtDebug.Text & vbNewLine & "Current: " & URL
        End If
    Else
    cboURL.Text = URL
    cboURL.Tag = URL
    End If
End If
End Sub


Private Sub wb_DownloadBegin(Index As Integer)
pgBar.Visible = True
End Sub

Private Sub wb_DownloadComplete(Index As Integer)
pgBar.Visible = False
End Sub

Private Sub wb_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
picSearch.Move X - (picSearch.Width \ 2), Y + (picSearch.Height \ 2)
If picSearch.Left < 0 Then picSearch.Left = 0
If picSearch.Top < wb(curWB).Top Then picSearch.Top = wb(curWB).Top
If picSearch.Left + picSearch.Width > Me.Width Then picSearch.Left = Me.Width - picSearch.Width - 120
If picSearch.Top + picSearch.Height > StatBar.Top Then picSearch.Top = StatBar.Top - picSearch.Height

End Sub

Private Sub wb_GotFocus(Index As Integer)
mnuEditCut.Enabled = 1
mnuEditCopy.Enabled = 1
mnuEditPaste.Enabled = 1
End Sub
Private Sub wb_LostFocus(Index As Integer)
mnuEditCut.Enabled = 0
mnuEditCopy.Enabled = 0
mnuEditPaste.Enabled = 0
End Sub
Private Sub wb_NavigateError(Index As Integer, ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
Silent = True

'=================================================================
'Prompt User to Remove Current Item if Navigate Error and Resolve
'=================================================================
'strFoundInDB = "AH" & i
'Allow this option to be set as well...
If skipResolveBadNav <> 1 Then
    Dim strResolveAgain As String
    FileNumber = FreeFile
    Select Case Left(strFoundInDB, 2)
        Case "AM"
            If MsgBox("URL was found in database but page cannot be displayed. Remove Item and Resolve?", vbYesNo, _
            "Found But Cannot Be Displayed") = vbYes Then
                
                strResolveAgain$ = lstAM.List(Right(strFoundInDB, Len(strFoundInDB) - 2))
                DeleteString "DNS Browser - DNS Database", strResolveAgain$, "DNS Database\Current\WWW_A-M.dat"
                lstAM.RemoveItem Right(strFoundInDB, Len(strFoundInDB) - 2)
                lstAM2.RemoveItem Right(strFoundInDB, Len(strFoundInDB) - 2)
            
            cboURL.Text = strResolveAgain$
            Call cmdGo_Click
            End If
            strFoundInDB$ = vbNullString
            Exit Sub
            
        Case "NZ"
            If MsgBox("URL was found in database but page cannot be displayed. Remove Item and Resolve?", vbYesNo, _
            "Found But Cannot Be Displayed") = vbYes Then
                
                strResolveAgain$ = lstNZ.List(Right(strFoundInDB, Len(strFoundInDB) - 2))
                DeleteString "DNS Browser - DNS Database", strResolveAgain$, "DNS Database\Current\WWW_N-Z.dat"
                lstNZ.RemoveItem Right(strFoundInDB, Len(strFoundInDB) - 2)
                lstNZ2.RemoveItem Right(strFoundInDB, Len(strFoundInDB) - 2)
                
            cboURL.Text = strResolveAgain$
            Call cmdGo_Click
            End If
            strFoundInDB$ = vbNullString
            Exit Sub
                   
        Case "AH"
            If MsgBox("URL was found in database but page cannot be displayed. Remove Item and Resolve?", vbYesNo, _
            "Found But Cannot Be Displayed") = vbYes Then
                
                strResolveAgain$ = lstAH.List(Right(strFoundInDB, Len(strFoundInDB) - 2))
                DeleteString "DNS Browser - DNS Database", strResolveAgain$, "DNS Database\Current\A-H.dat"
                lstAH.RemoveItem Right(strFoundInDB, Len(strFoundInDB) - 2)
                lstAH2.RemoveItem Right(strFoundInDB, Len(strFoundInDB) - 2)
            
            cboURL.Text = strResolveAgain$
            Call cmdGo_Click
            End If
            strFoundInDB$ = vbNullString
            Exit Sub
                    
        Case "IP"
            If MsgBox("URL was found in database but page cannot be displayed. Remove Item and Resolve?", vbYesNo, _
            "Found But Cannot Be Displayed") = vbYes Then
                
                strResolveAgain$ = lstIP.List(Right(strFoundInDB, Len(strFoundInDB) - 2))
                DeleteString "DNS Browser - DNS Database", strResolveAgain$, "DNS Database\Current\I-P.dat"
                lstIP.RemoveItem Right(strFoundInDB, Len(strFoundInDB) - 2)
                lstIP2.RemoveItem Right(strFoundInDB, Len(strFoundInDB) - 2)
                
            cboURL.Text = strResolveAgain$
            Call cmdGo_Click
            End If
            strFoundInDB$ = vbNullString
            Exit Sub
                    
        Case "QZ"
            If MsgBox("URL was found in database but page cannot be displayed. Remove Item and Resolve?", vbYesNo, _
            "Found But Cannot Be Displayed") = vbYes Then
                
                strResolveAgain$ = lstQZ.List(Right(strFoundInDB, Len(strFoundInDB) - 2))
                DeleteString "DNS Browser - DNS Database", strResolveAgain$, "DNS Database\Current\Q-Z.dat"
                lstQZ.RemoveItem Right(strFoundInDB, Len(strFoundInDB) - 2)
                lstQZ2.RemoveItem Right(strFoundInDB, Len(strFoundInDB) - 2)
                
            cboURL.Text = strResolveAgain$
            Call cmdGo_Click
            End If
            strFoundInDB$ = vbNullString
            Exit Sub
            
    End Select
Else
    skipResolveBadNav = 0
    Exit Sub
End If
'==================================================
'Mission complete =)
'==================================================

If chkAlreadyResolved = 5 Then
    Dim bURL As String, i As Long
    If Left(URL, 7) = "http://" Then
        bURL$ = Replace(URL, "http://", vbNullString)
            For i = 1 To Len(bURL$)
                If InStr(i, bURL$, "/") = 1 Then
                    bURL$ = Left(bURL$, i - 1)
                    Exit For
                End If
                DoEvents
            Next i
            MsgBox bURL$
    Else
        bURL$ = URL
            For i = 1 To Len(bURL$)
                If InStr(i, bURL$, "/") Then
                    bURL$ = Left(bURL$, i - 1)
                    Exit For
                End If
                DoEvents
            Next i
            MsgBox bURL$
    End If
    
         txtDebug.Text = txtDebug.Text & vbNewLine & "Specified URL not found in Database. Resolving..."
                 StatBar.Panels(1).Text = "Specified URL not found in Database. Resolving..."
            If ValDBS = 3 Then
                    If MsgBox("Would you like to Resolve/Add this website to the Database?" & vbNewLine & _
                            "URL: '" & URL & "'", vbYesNo, "Resolve/Add to Database?") = vbYes Then
                        ResolveAdd = 1
                    Else
                        ResolveAdd = 0
                    End If
                resolveDNS bURL$
            Else
                If ValDBS = 1 Then
                    ResolveAdd = 1
                    resolveDNS bURL$
                End If
                
                If ValDBS = 2 Then
                    ResolveAdd = 0
                    resolveDNS bURL$
                End If
            End If
            
    wb(Index).Silent = True
    chkAlreadyResolved = 1
Else
    chkAlreadyResolved = 0
End If
End Sub
Private Sub wb_NewWindow2(Index As Integer, ppdisp As Object, Cancel As Boolean)
Dim exec As Long
Dim i As Long, j As Long, k As Long


    If frmOptions.lstPB.ListCount <> 0 Then
        If mnuPopupBlockSet.Caption = "Turn Off Pop-up Blocker" Then
            For i = 0 To frmOptions.lstPB.ListCount - 1
                    If InStr(wb(curWB).LocationURL, frmOptions.lstPB.List(i)) = 1 Then
                        GoTo openIt
                    End If
            
                    If InStr(cboURL.Text, frmOptions.lstPB.List(i)) = 1 Then
                        GoTo openIt
                    End If
                DoEvents
            Next i
                    If TempAllow <> 1 Then
                        Cancel = True
                        wb(Index).Silent = True
                        If frmOptions.chkPopSnd.Value = 1 Then
                            exec = sndPlaySound("DATA\Sounds\popup-blocked.wav", &H1)
                            StatBar.Panels.Item(1).Text = "Pop-up blocked!"
                        End If
                        If frmOptions.chkPopNotify.Value = 1 Then
                            PopUpResize 1
                        End If
                    Else
                        Set Web_V1 = wb(Index).object
                        Exit Sub
                    End If
        Else
                Cancel = False
                Set Web_V1 = wb(Index).object
        End If
    
    Else
        If mnuPopupBlockSet.Caption = "Turn Off Pop-up Blocker" Then
            If TempAllow <> 1 Then
                Cancel = True
                wb(Index).Silent = True
                    If frmOptions.chkPopSnd.Value = 1 Then
                        exec = sndPlaySound("DATA\Sounds\popup-blocked.wav", &H1)
                        StatBar.Panels.Item(1).Text = "Pop-up blocked!"
                    End If
                    If frmOptions.chkPopNotify.Value = 1 Then
                        PopUpResize 1
                    End If
            Else
                Set Web_V1 = wb(Index).object
            End If
        Else
openIt:
            Set Web_V1 = wb(Index).object
        End If
    End If
End Sub
Private Function PopUpResize(ByVal bVisible As Byte)
With picPopup
    .Width = Me.Width
    lblClose.Left = .Width - lblClose.Width - 100
    If bVisible = 1 Then
    
         .Visible = True
         txtDebug.Top = .Top + .Height - 30
         
             For i = 0 To wb.UBound
                 For j = 0 To lstUnloaded.ListCount - 1
                     If i = lstUnloaded.List(j) Then
                         SkipResize = 1
                     End If
                 Next j
                 
                 If SkipResize <> 1 Then
                     wb(i).Top = .Top + .Height - 30
                     wb(i).Height = wb(i).Height - .Height + 30
                     SkipResize = 0
                 End If
             Next i
    
    Else
         wb(curWB).Top = .Top
         wb(curWB).Height = wb(curWB).Height + .Height
         picPopup.Visible = False
         txtDebug.Top = .Top
         
             For i = 0 To wb.UBound
                 For j = 0 To lstUnloaded.ListCount - 1
                     If i = lstUnloaded.List(j) Then
                         SkipResize = 1
                     End If
                 Next j
                 
                 If SkipResize <> 1 Then
                     wb(i).Top = .Top
                     wb(i).Height = wb(i).Height + .Height - 30
                     SkipResize = 0
                 End If
             Next i
    End If
End With
End Function


Private Sub Web_V1_NewWindow(ByVal URL As String, ByVal Flags As Long, ByVal TargetFrameName As String, _
                            PostData As Variant, ByVal Headers As String, Processed As Boolean)

FilterCheck (URL)
If xCancel = 1 Then
    Cancel = True
    xCancel = 0
    Exit Sub
End If
Processed = True
    
strNewWindow$ = URL
ProcDispN = 1

mnuFileNewWindow_Click
End Sub
Private Sub wb_ProgressChange(Index As Integer, ByVal Progress As Long, ByVal ProgressMax As Long)
If Index = curWB Then
    On Error Resume Next
    pgBar.Max = ProgressMax
    pgBar.Value = Progress
End If
End Sub
Private Function resolveDNS(rURL As String)
Dim DeniedResolverLV1 As Byte
Dim DNSServer As New Collection
DNSServer.Add "151.164.64.201"
DNSServer.Add "168.156.219.7"
DNSServer.Add "168.156.220.7"
DNSServer.Add "216.151.192.222"

DeniedResolverLV1 = 1
FileNumber = FreeFile
If Left(LCase(rURL), 4) <> "www." Then
    rURL = "www." & rURL
End If

'ok first we need to get the IP somehow, we will do this using PING tool on DNS Stuff
Dim StripData() As String, Chunk$, ChunkChk() As String, IP$, CategorizeURL As Integer
Dim bExtention, i As Long, j As Long
Dim ProperURL$, getBaseURL() As String

            ProperURL$ = rURL
                If InStr(rURL, "/") Then
                    getBaseURL$ = Split(rURL, "/", -1, 1)
                    ProperURL$ = getBaseURL(0)
                    bExtention = 1
                End If
                
Do Until DeniedResolverLV1 = 3 Or DeniedResolverLV1 = 4
    For i = 1 To DNSServer.Count
        bExtention = 0
        
        If DeniedResolverLV1 = 1 Then

            ' NEW ATTEMPT TO RESOLVE!
          Shell "cmd /c nslookup " & ProperURL$ & " " & DNSServer(i) & " > C:\tmplookup.dat", vbHide
          txtDebug.Text = txtDebug.Text & vbNewLine & "Attempting Server #" & i & "..."
                  StatBar.Panels(1).Text = "Attempting Server #" & i & "..."
          TimeOut (IntTimeOut)
            
            Dim TempStr$
            Open "C:\tmplookup.dat" For Input As #FileNumber
                Do While Not EOF(FileNumber)
                    Input #FileNumber, TempStr$
                    If Left(TempStr$, 8) = "Address:" Then
                        IP$ = Trim(Right(TempStr$, Len(TempStr$) - 8))
                    Else
                        If Left(TempStr$, 10) = "Addresses:" Then
                            StripData$ = Split(TempStr$, ",", -1, 1)
                            IP$ = Trim(Right(StripData(0), Len(StripData(0)) - 10))
                            DeniedResolverLV1 = 3
                        End If
                    End If
                Loop
            Close #FileNumber
        Else
            DeniedResolverLV1 = 1
        End If
        
        If IP$ = vbNullString Or Len(IP$) > 15 Then
            If i = DNSServer.Count Then
                DeniedResolverLV1 = 4
                Exit Do
            End If
        Else
            Exit Do
            DeniedResolverLV1 = 0
        End If
    Next i
DoEvents
Loop

If DeniedResolverLV1 = 4 Then ' OLD ONE
        txtDebug.Text = txtDebug.Text & vbNewLine & "Attempting Final Server..."
        StatBar.Panels(1).Text = "Attempting Final Server..."
    Do Until findDNS.StillExecuting = False 'prevent from getting SRC before inet can finish executing
        TimeOut (IntTimeOut)
        DoEvents
    Loop
    
    findDNS.Execute "http://69.2.200.183/tools/ping.ch?ip=" & rURL

Do Until findDNS.StillExecuting = False
    TimeOut (IntTimeOut) 'prevent error if inet takes longer than expected to execute
    DoEvents 'less vital
Loop

Chunk$ = frmBrowser.findDNS.GetChunk(1024)


    If InStr(Chunk$, "Sorry") = 1 Then
        frmBrowser.txtDebug.Text = frmBrowser.txtDebug.Text & vbNewLine & "DNS could not be resolved..."
        StatBar.Panels(1).Text = "DNS could not be resolved..."
        wb(curWB).Navigate App.Path & "\" & "404.html"
        Exit Function
    End If
    DoEvents

If Len(Chunk$) = 0 Then
    frmBrowser.txtDebug.Text = frmBrowser.txtDebug.Text & vbNewLine & "DNS could not be resolved..."
    StatBar.Panels(1).Text = "DNS could not be resolved..."
Exit Function
End If

'grab Chunk
StripData$ = Split(Chunk$, "[", -1, 1)
StripData2 = StripData(1)
StripData2 = Split(StripData2, "]", -1, 1) 'ok now we split up the junk and got the IP!
IP$ = Trim(StripData2(0))
End If

CategorizeURL = Asc(Mid(UCase(ProperURL$), 5, 1)) - 64

If ResolveAdd = 1 Then
    skipResolveBadNav = 1
    Select Case CategorizeURL
        Case Is <= 13
        frmBrowser.lstAM.AddItem ProperURL$
        frmBrowser.lstAM2.AddItem IP$
            
        Open "DNS Database\Current\WWW_A-M.dat" For Append As #FileNumber
            Print #FileNumber, ProperURL$ & "=" & IP$
        Close #FileNumber
        
        Case Is >= 14
        frmBrowser.lstNZ.AddItem ProperURL$
        frmBrowser.lstNZ2.AddItem IP$
    
        Open "DNS Database\Current\WWW_N-Z.dat" For Append As #FileNumber
            Print #FileNumber, ProperURL$ & "=" & IP$
        Close #FileNumber
    End Select
    
    ProperURL$ = Replace(ProperURL$, "www.", "")
    
    Select Case CategorizeURL
        Case Is <= 8
        frmBrowser.lstAH.AddItem ProperURL$
        frmBrowser.lstAH2.AddItem IP$
        
        Open "DNS Database\Current\A-H.dat" For Append As #FileNumber
            Print #FileNumber, ProperURL$ & "=" & IP$
        Close #FileNumber
        
        frmBrowser.txtDebug.Text = frmBrowser.txtDebug.Text & vbNewLine & "DNS Successfully Resolved!"
        StatBar.Panels(1).Text = "DNS Successfully Resolved!"
        checkDNS rURL, "1"
        Exit Function
            
        Case Is <= 16
        frmBrowser.lstIP.AddItem ProperURL$
        frmBrowser.lstIP2.AddItem IP$
        
        Open "DNS Database\Current\I-P.dat" For Append As #FileNumber
            Print #FileNumber, ProperURL$ & "=" & IP$
        Close #FileNumber
        
        frmBrowser.txtDebug.Text = frmBrowser.txtDebug.Text & vbNewLine & "DNS Successfully Resolved!"
        StatBar.Panels(1).Text = "DNS Successfully Resolved!"
        checkDNS rURL, "1"
        Exit Function
            
        Case Is >= 17
        frmBrowser.lstNZ.AddItem ProperURL$
        frmBrowser.lstNZ2.AddItem IP$
        
        Open "DNS Database\Current\Q-Z.dat" For Append As #FileNumber
            Print #FileNumber, ProperURL$ & "=" & IP$
        Close #FileNumber
        
        frmBrowser.txtDebug.Text = frmBrowser.txtDebug.Text & vbNewLine & "DNS Successfully Resolved!"
        StatBar.Panels(1).Text = "DNS Successfully Resolved!"
        checkDNS rURL, "1"
        Exit Function
    End Select
Else
    frmBrowser.txtDebug.Text = frmBrowser.txtDebug.Text & vbNewLine & "DNS Successfully Resolved!"
    StatBar.Panels(1).Text = "DNS Successfully Resolved!"
    If bExtention = 0 Then wb(curWB).Navigate IP$
    If bExtention = 1 Then wb(curWB).Navigate IP$ & "/" & getBaseURL(1)
    cboURL.Text = rURL
    cboURL.Tag = rURL
End If
End Function

Private Sub wb_StatusTextChange(Index As Integer, ByVal Text As String)
If Index = curWB Then
    If Text <> "Done" Then
        StatBar.Panels(1).Text = Text
    End If
End If
End Sub

Private Sub wb_TitleChange(Index As Integer, ByVal Text As String)
If Index = curWB Then
    Me.Caption = Text & "- DNS Browser"
End If
End Sub
