VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWhois 
   Caption         =   "Whois Client"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11025
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstCSet 
      Appearance      =   0  'Flat
      ForeColor       =   &H00404040&
      Height          =   1710
      ItemData        =   "frmWhois.frx":0000
      Left            =   7920
      List            =   "frmWhois.frx":0002
      TabIndex        =   8
      Top             =   2400
      Width           =   3015
   End
   Begin DNSBrowser.isButton cmdAdd 
      Height          =   300
      Left            =   9720
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Style           =   8
      Caption         =   "Add Server"
      IconAlign       =   1
      iNonThemeStyle  =   0
      HighlightColor  =   4210752
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
   Begin MSComctlLib.StatusBar sbInfo 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   4905
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7145
            MinWidth        =   7145
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4764
            MinWidth        =   4764
            Text            =   "Packet Size"
            TextSave        =   "Packet Size"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3794
            MinWidth        =   3794
            Text            =   "Response Time"
            TextSave        =   "Response Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3529
            MinWidth        =   3529
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstServers 
      Appearance      =   0  'Flat
      ForeColor       =   &H00404040&
      Height          =   2190
      ItemData        =   "frmWhois.frx":0004
      Left            =   7920
      List            =   "frmWhois.frx":001A
      TabIndex        =   5
      Top             =   0
      Width           =   3015
   End
   Begin VB.ListBox lst_hist 
      Height          =   540
      ItemData        =   "frmWhois.frx":00F3
      Left            =   480
      List            =   "frmWhois.frx":00F5
      TabIndex        =   3
      Top             =   5400
      Width           =   7335
   End
   Begin MSWinsockLib.Winsock sckWhois 
      Left            =   7920
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtQuery 
      CausesValidation=   0   'False
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   4560
      Width           =   7215
   End
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   4455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   7815
   End
   Begin VB.Label lblSelect 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select WHOIS Server"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   7920
      TabIndex        =   4
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label lblQuery 
      Caption         =   "Query:"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4560
      Width           =   735
   End
End
Attribute VB_Name = "frmWhois"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hist As Integer, strQuery As String
Dim FileNumber As Integer, strResponse As String
Dim initResponse As Long, cResponse As Long

Private Sub cmdAdd_Click()
Dim strServerLoc As String, strServerIP As String
strServerLoc$ = InputBox("Please enter a server location:", "Add whois Server", "whois.site.com")
strServerIP$ = InputBox("Please enter the server's IP Address.", "Configure IP", "0.0.0.0")
If strServerLoc$ <> vbNullString And strServerIP$ <> vbNullString Then
    lstCSet.AddItem strServerLoc$ & "[ " & strServerIP$ & " ]"
End If
End Sub

Private Sub Form_Load()
hist = 0
End Sub

Private Sub sckWhois_Connect()
sbInfo.Panels(1).Text = "Connect: OK"
sckWhois.SendData strQuery$
sckWhois.SendData vbCrLf
sbInfo.Panels(1).Text = "Retrieving data..."
End Sub

Private Sub sckWhois_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
sckWhois.GetData strData$

If Len(strData$) <= 0 Then Exit Sub

strData$ = Replace(strData$, vbLf, vbCrLf)

For i = 1 To Len(strData$)
    If InStr(i, LCase(strData$), "timeout") Then
        strResponse$ = "Response: Connection timed out"
    End If
    DoEvents
Next i

If strResponse$ = vbNullString Then
    strResponse$ = "Response received"
End If

txtData.Text = txtData.Text & Replace(strData$, frmBrowser.txtChar.Text, vbNewLine)
sbInfo.Panels(1).Text = strResponse$
sbInfo.Panels(2).Text = "Packet Size: " & Len(txtData) & " bytes" 'bytesTotal & " bytes"
cResponse = GetTickCount
sbInfo.Panels(3).Text = "Response Time: " & cResponse - initResponse & " ms"
End Sub

Private Sub sckWhois_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
sbInfo.Panels(1).Text = "Response: " & Description
End Sub

Private Sub txtQuery_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyUp Then
    If hist = 0 Then 'if index is 0 then prevent potential error
        Exit Sub
    End If
    hist = hist - 1
    txtQuery.Text = lst_hist.List(hist)
    txtQuery.SelStart = Len(txtQuery.Text)
End If

If KeyCode = vbKeyDown Then
    If hist = lst_hist.ListCount Then 'if index is 0 then prevent potential error
        Exit Sub
    End If
    hist = hist + 1
    txtQuery.Text = lst_hist.List(hist)
    txtQuery.SelStart = Len(txtQuery.Text)
End If

If KeyCode = vbKeyReturn Then
        strResponse$ = vbNullString
        'whois.alldomains.com [ 64.124.14.21 ]
        'whois.apnic.net [ 202.12.29.13 ]
        'whois.arin.net [ 192.149.252.44 ]
        'whois.internic.net [ 198.41.0.6 ]
        'whois.names4ever.com [216.55.191.55]
        'whois.ripe.net [ 193.0.0.135 ]
        
        sckHandleInfo
        
        If sckServer$ = vbNullString Then
            sckServer$ = "192.149.252.44"
            sckHandle$ = "whois.arin.net"
        End If
        
        If Left(LCase(txtQuery.Text), 4) = "save" Then
            FileNumber = FreeFile
            Open Right(LCase(txtQuery.Text), Len(txtQuery.Text) - 5) For Append As #FileNumber
                Print #FileNumber, txtData.Text
            Close #FileNumber
        Else
            strQuery$ = txtQuery.Text
            txtQuery.Text = vbNullString
            sckWhois.Close
            sckWhois.Connect sckServer$, sckPort
            sbInfo.Panels(4).Text = sckHandle$
            initResponse = GetTickCount
        End If
    If strQuery$ <> vbNullString Then
        lst_hist.AddItem strQuery$ 'creates history of used commands in listbox
        hist = lst_hist.ListCount 'creates index
        txtData.Text = vbNullString
    End If
 '   txtQuery.Text = vbNullString
    
End If
End Sub
