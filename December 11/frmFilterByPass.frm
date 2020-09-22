VERSION 5.00
Begin VB.Form frmFilterByPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filter Control Password Request"
   ClientHeight    =   1755
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3870
   Icon            =   "frmFilterByPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1036.912
   ScaleMode       =   0  'User
   ScaleWidth      =   3633.72
   StartUpPosition =   1  'CenterOwner
   Begin DNSBrowser.isButton cmdOK 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Icon            =   "frmFilterByPass.frx":058A
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
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
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
      Left            =   1320
      PasswordChar    =   "•"
      TabIndex        =   1
      Top             =   720
      Width           =   2325
   End
   Begin DNSBrowser.isButton cmdCancel 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Icon            =   "frmFilterByPass.frx":05A6
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
   Begin VB.Label lblAllow 
      Caption         =   "Allow:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3735
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   225
      TabIndex        =   0
      Top             =   720
      Width           =   1080
   End
End
Attribute VB_Name = "frmFilterByPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    ByPassFilter = 0
    Me.Hide
End Sub
Private Sub cmdOK_Click()
    'check for correct password
    If txtPassword = strFilterPass Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        ByPassFilter = 1
        Me.Hide
    Else
        MsgBox "Invalid Password, try again!", vbCritical, "Filter Control"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Select Case UnloadMode
    Case 1
    ByPassFilter = 0
    
    Case Else
    ByPassFilter = 0
End Select
End Sub
