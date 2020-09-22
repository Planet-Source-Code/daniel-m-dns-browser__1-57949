VERSION 5.00
Begin VB.Form frmHEHE 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   663
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   1800
      Picture         =   "frmHEHE.frx":0000
      ScaleHeight     =   4425
      ScaleWidth      =   4185
      TabIndex        =   1
      Top             =   4200
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   1320
      Picture         =   "frmHEHE.frx":41EF2
      ScaleHeight     =   4455
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3960
      Top             =   5760
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   5640
      Picture         =   "frmHEHE.frx":83DE4
      Top             =   1560
      Visible         =   0   'False
      Width           =   4500
   End
End
Attribute VB_Name = "frmHEHE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=====================================================
'Description: Hehe, I wrote this because some people in
'my class wanted to browse some websites that the
'instructors may find inappropriate. Does not work well
'on all monitors.
'=====================================================

Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long


Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long 'self explanatory
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Dim POS As POINTAPI
Dim lngStart

Private Sub Form_Load()
lngStart = 0
End Sub

Private Sub Timer1_Timer()
If lngStart <> 1 Then
    lngStart = 1
    BringWindowToTop (frmBrowser.hwnd)
Else
End If
GetCursorPos POS
Picture1.Move POS.X - (Picture1.Width / 2) - 25, POS.Y - (Picture1.Height / 2) - 25
End Sub
