VERSION 5.00
Begin VB.Form frmDNS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DNS Tool"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DNSlst.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   10020
   StartUpPosition =   1  'CenterOwner
   Begin DNSBrowser.isButton cmdResolve 
      Height          =   375
      Left            =   6480
      TabIndex        =   18
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Icon            =   "DNSlst.frx":058A
      Style           =   9
      Caption         =   "Resolve And Add"
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
   Begin VB.Frame fraADD 
      Caption         =   "Add DNS Information to Database"
      Height          =   2055
      Left            =   5880
      TabIndex        =   12
      Top             =   0
      Width           =   4095
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   600
         TabIndex        =   16
         Text            =   "0.0.0.0"
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtURL 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   600
         TabIndex        =   14
         Text            =   "http://www."
         Top             =   480
         Width           =   3375
      End
      Begin DNSBrowser.isButton cmdAdd 
         Height          =   375
         Left            =   2400
         TabIndex        =   19
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Icon            =   "DNSlst.frx":05A6
         Style           =   9
         Caption         =   "Add to Database"
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
      Begin VB.Label lblipx 
         Caption         =   "IP:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1010
         Width           =   255
      End
      Begin VB.Label lblurl 
         Caption         =   "URL:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   510
         Width           =   375
      End
   End
   Begin VB.Frame fraDNSD 
      Caption         =   "DNS Database List Sorting Tool"
      Height          =   8055
      Left            =   40
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.ListBox lstAM 
         Height          =   1020
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   5415
      End
      Begin VB.ListBox lstAH 
         Height          =   1020
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   3360
         Width           =   5415
      End
      Begin VB.ListBox lstIP 
         Height          =   1020
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   4800
         Width           =   5415
      End
      Begin VB.ListBox lstQZ 
         Height          =   1020
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   6240
         Width           =   5415
      End
      Begin VB.ListBox lstNZ 
         Height          =   1020
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   1920
         Width           =   5415
      End
      Begin DNSBrowser.isButton cmdSort 
         Height          =   375
         Left            =   1680
         TabIndex        =   20
         Top             =   7560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Icon            =   "DNSlst.frx":05C2
         Style           =   9
         Caption         =   "Sort and Save Lists"
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
      Begin DNSBrowser.isButton cmdRefresh 
         Height          =   375
         Left            =   3720
         TabIndex        =   21
         Top             =   7560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Icon            =   "DNSlst.frx":05DE
         Style           =   9
         Caption         =   "Refresh Lists"
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
      Begin VB.Label lblAM 
         BackStyle       =   0  'Transparent
         Caption         =   "WWW A to M Websites"
         Height          =   255
         Left            =   165
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "WWW N to Z Websites"
         Height          =   255
         Left            =   165
         TabIndex        =   10
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label lblAH 
         BackStyle       =   0  'Transparent
         Caption         =   "A to H Websites"
         Height          =   255
         Left            =   165
         TabIndex        =   9
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Label lblIP 
         BackStyle       =   0  'Transparent
         Caption         =   "I to P Websites"
         Height          =   255
         Left            =   165
         TabIndex        =   8
         Top             =   4560
         Width           =   2775
      End
      Begin VB.Label lblQZ 
         BackStyle       =   0  'Transparent
         Caption         =   "Q to Z Websites"
         Height          =   255
         Left            =   165
         TabIndex        =   7
         Top             =   6000
         Width           =   2655
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "Ready."
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   7320
         Width           =   3375
      End
   End
   Begin VB.Label lblinfo 
      Caption         =   "Add DNS Information does not currently work at the moment.The List Sorting Tool, however works great."
      Height          =   5895
      Left            =   5880
      TabIndex        =   17
      Top             =   2160
      Width           =   4095
   End
End
Attribute VB_Name = "frmDNS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGo_Click()
'wb.Navigate txtUrl.Text
'Shell "cmd /c ping" & txtUrl.Text & "> txturl.txt"
txtURLx.Text = txtURL.Text

wb.Navigate txtURL.Text
Winsock1.Close
Winsock1.Connect txtURL.Text, 80

End Sub


Private Sub cmdGo2_Click()
wb.Navigate txtURLx.Text
End Sub

Private Sub cmdLoad_Click()
Dim dnsInfo$, lst1Item, lst2Item, splitDNS
Open "DNSlst.txt" For Input As #FileNumber
Do Until EOF(FileNumber)
    Input #FileNumber, dnsInfo$
    If dnsInfo$ <> "[DNS Browser - DNS Database]" And dnsInfo$ <> vbNullString Then
        splitDNS = Split(dnsInfo$, "=", -1, 1)
        lst1Item = splitDNS(0)
        lst2Item = splitDNS(1)
        lst1.AddItem lst1Item
        lst2.AddItem lst2Item
    End If
Loop
Close #FileNumber

End Sub

Private Sub Command1_Click()
Open "DNSlst.txt" For Append As #FileNumber
    For i = 0 To lst1.ListCount - 1
        Print #FileNumber, lst1.List(i) & "=" & lst2.List(i)
    Next i
Close #FileNumber

End Sub

Private Sub wb_DocumentComplete(ByVal pDisp As Object, URL As Variant)
txtURLx.Text = URL
End Sub


Private Sub Winsock1_Connect()

lst1.AddItem Winsock1.RemoteHost
lst2.AddItem Winsock1.RemoteHostIP

If Left(Winsock1.RemoteHost, 4) = "www." Then
    lst1.AddItem Right(Winsock1.RemoteHost, Len(Winsock1.RemoteHost) - 4)
    lst2.AddItem Winsock1.RemoteHostIP
End If
txtURL.Text = "www."
End Sub

Private Sub cmdRefresh_Click()
refreshDNS
End Sub

Private Sub cmdResolve_Click()
txtIP.Enabled = False

'If resolveDNS(txtURL.Text, Resolved) = 1 Then
'    MsgBox "DNS Successfully Resolved!"
'End If

txtIP.Enabled = True
End Sub

Private Sub cmdSort_Click()
lblState.Caption = "Loading Database Information..."
For i = 0 To frmBrowser.lstAM.ListCount - 1
    lstAM.AddItem frmBrowser.lstAM.List(i) & "=" & frmBrowser.lstAM2.List(i)
Next i
    
For i = 0 To frmBrowser.lstNZ.ListCount - 1
    lstNZ.AddItem frmBrowser.lstNZ.List(i) & "=" & frmBrowser.lstNZ2.List(i)
Next i

For i = 0 To frmBrowser.lstAH.ListCount - 1
    lstAH.AddItem frmBrowser.lstAH.List(i) & "=" & frmBrowser.lstAH2.List(i)
Next i

For i = 0 To frmBrowser.lstIP.ListCount - 1
    lstIP.AddItem frmBrowser.lstIP.List(i) & "=" & frmBrowser.lstIP2.List(i)
Next i

For i = 0 To frmBrowser.lstQZ.ListCount - 1
    lstQZ.AddItem frmBrowser.lstQZ.List(i) & "=" & frmBrowser.lstQZ2.List(i)
Next i
lblState.Caption = "Sorting Complete!"
lblState.Caption = "Updating Database Information..."
Open "DNS Database\Current\WWW_A-M.dat" For Output As #FileNumber
    Print #FileNumber, "[DNS Browser - DNS Database]"
        For i = 0 To lstAM.ListCount - 1
            Print #FileNumber, lstAM.List(i)
        Next i
Close #FileNumber

Open "DNS Database\Current\WWW_N-Z.dat" For Output As #FileNumber
    Print #FileNumber, "[DNS Browser - DNS Database]"
        For i = 0 To lstNZ.ListCount - 1
            Print #FileNumber, lstNZ.List(i)
        Next i
Close #FileNumber

Open "DNS Database\Current\A-H.dat" For Output As #FileNumber
    Print #FileNumber, "[DNS Browser - DNS Database]"
        For i = 0 To lstAH.ListCount - 1
            Print #FileNumber, lstAH.List(i)
        Next i
Close #FileNumber


Open "DNS Database\Current\I-P.dat" For Output As #FileNumber
    Print #FileNumber, "[DNS Browser - DNS Database]"
        For i = 0 To lstIP.ListCount - 1
            Print #FileNumber, lstIP.List(i)
        Next i
Close #FileNumber


Open "DNS Database\Current\Q-Z.dat" For Output As #FileNumber
    Print #FileNumber, "[DNS Browser - DNS Database]"
        For i = 0 To lstQZ.ListCount - 1
            Print #FileNumber, lstQZ.List(i)
        Next i
Close #FileNumber

lblState.Caption = "Database Sorted and Updated."
End Sub

Private Sub txtIP_Click()
If txtIP.Text = "0.0.0.0" Then
    txtIP.Text = vbNullString
End If

End Sub
