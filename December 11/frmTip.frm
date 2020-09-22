VERSION 5.00
Begin VB.Form frmTip 
   Caption         =   "Tip of the Day"
   ClientHeight    =   2685
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5415
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin DNSBrowser.isButton cmdNextTip 
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Style           =   9
      Caption         =   "&Next Tip"
      IconAlign       =   1
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
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
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Appearance      =   0  'Flat
      Caption         =   "&Show Tips at Startup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   120
      Picture         =   "frmTip.frx":5D52
      ScaleHeight     =   2055
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know..."
         Height          =   255
         Left            =   540
         TabIndex        =   3
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1155
         Left            =   180
         TabIndex        =   2
         Top             =   840
         Width           =   3255
      End
   End
   Begin DNSBrowser.isButton cmdOK 
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Icon            =   "frmTip.frx":605C
      Style           =   9
      Caption         =   "OK"
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
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Long
Dim TipCollection As New Collection

Function LoadTips()

Dim TempStr As String
FileNumber = FreeFile
Open "DATA\tips.dat" For Input As #FileNumber
    Do While Not EOF(FileNumber)
    Input #FileNumber, TempStr$
        TipCollection.Add TempStr$
    DoEvents
    Loop
Close #FileNumber
lblTipText.Caption = TipCollection.Item(1)
End Function

Private Sub chkLoadTipsAtStartup_Click()
WriteString "Miscellaneous", "TOTD", chkLoadTipsAtStartup.Value, "DATA\options.dat"
End Sub

Private Sub cmdNextTip_Click()
If i = TipCollection.Count Then
    i = 1
End If

i = i + 1
lblTipText.Caption = TipCollection.Item(i)
End Sub

Private Sub cmdOK_Click()
Me.Hide
End Sub

Private Sub Form_Load()

i = 1
LoadTips
    
End Sub
