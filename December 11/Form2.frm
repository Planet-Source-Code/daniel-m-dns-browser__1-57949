VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Port Scanner"
   ClientHeight    =   7545
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8565
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4920
      TabIndex        =   13
      Top             =   5880
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Text            =   "20000"
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop!!"
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   600
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2760
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   2280
      Top             =   1080
   End
   Begin VB.CommandButton Cmdstart 
      Caption         =   "Start"
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   0
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "1"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      MaxLength       =   15
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      CausesValidation=   0   'False
      ForeColor       =   &H00000000&
      Height          =   4320
      ItemData        =   "Form2.frx":08CA
      Left            =   120
      List            =   "Form2.frx":08CC
      TabIndex        =   0
      Top             =   1200
      Width           =   6975
   End
   Begin VB.Label Label7 
      Caption         =   "Port# End:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Ports Found"
      Height          =   240
      Left            =   360
      TabIndex        =   10
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Port Number Scanning"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Port # Start:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Ip Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu mnumestuff 
      Caption         =   "Stuff"
      Begin VB.Menu mnuIF 
         Caption         =   "IP Finder"
      End
      Begin VB.Menu datmenu 
         Caption         =   "Data Checker"
      End
      Begin VB.Menu mnulps 
         Caption         =   "Local Port Scanner"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim port As Integer
Private Sub datmenu_Click()
Form1.Show
Unload Me
End Sub

Private Sub Form_Load()
Timer5.Enabled = False
Label3 = "0"
Label5 = "0"
List1.Clear
Text1 = Winsock1.LocalIP
Cmdstart.Enabled = True
End Sub

Private Sub Cmdstart_Click()
Cmdstart.Enabled = False
Command2.Enabled = True
Winsock1.Close
List1.Clear
port = Text2.Text
Label3 = "0" 'lblopen =label3
Label5 = "0" 'lblport = label5
Timer5.Enabled = True
End Sub

Private Sub mnuIF_Click()
Form5.Show
Unload Me
End Sub

Private Sub mnulps_Click()
Form4.Show
Unload Me
End Sub

Private Sub Timer5_Timer()

If port = "65256" Then
Timer5.Enabled = False
Text2.Text = "1"
End If
port = port + 1
Label5.Caption = port
Winsock1.Close
Winsock1.Connect Text1.Text, port

End Sub

Private Sub Command2_Click()
Cmdstart.Enabled = True
Command2.Enabled = False
port = "1"
Timer5.Enabled = False
End Sub

Private Sub Winsock1_Connect()
List1.AddItem ("Port # : " & Winsock1.RemotePort & strdata)
Winsock1.Close
Label3 = Label3 + 1
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strdata As String
Winsock1.GetData strdata
Text4.Text = strdata
End Sub
