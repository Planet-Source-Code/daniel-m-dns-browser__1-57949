VERSION 5.00
Begin VB.Form frmFavorites 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Organize Favorites"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFavorites.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstFav 
      Height          =   4380
      Left            =   3120
      TabIndex        =   11
      Top             =   120
      Width           =   3135
   End
   Begin VB.FileListBox lstFavorites 
      Height          =   240
      Left            =   6960
      Pattern         =   "*.fav"
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.ListBox lstSURLs 
      Height          =   4380
      Left            =   9600
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
   Begin VB.ListBox lstSort 
      Height          =   4140
      Left            =   8400
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.ListBox lstURLs 
      Height          =   4380
      Left            =   6480
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame frmInfo 
      Caption         =   "Information"
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   2895
      Begin VB.Label lblMod 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label lblTitle 
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label lblModified 
         Caption         =   "Modified:"
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
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblURL 
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.site.com"
         Height          =   480
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   2655
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblname 
         Caption         =   "Website Title"
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2655
         WordWrap        =   -1  'True
      End
   End
   Begin DNSBrowser.isButton cmdRename 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Icon            =   "frmFavorites.frx":038A
      Style           =   9
      Caption         =   "Rename"
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
   Begin DNSBrowser.isButton cmdDelete 
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Icon            =   "frmFavorites.frx":03A6
      Style           =   9
      Caption         =   "Delete"
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
   Begin DNSBrowser.isButton cmdClose 
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Icon            =   "frmFavorites.frx":03C2
      Style           =   9
      Caption         =   "Close"
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
   Begin VB.Label lblInfo 
      Caption         =   "There is currently no folder support for favorites. Favorites may support more features in the future."
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim chkSorted As Byte
Private Sub cmdClose_Click()
Dim Title As String, URLf$, Folder$, TempStr As String
FileNumber = FreeFile
For i = 1 To frmBrowser.mnuFavoriteT.Count - 1
    Unload frmBrowser.mnuFavoriteT(i)
DoEvents
Next i

frmBrowser.lstFav.Refresh

For i = 0 To lstFav.ListCount - 1
    Open lstFavorites.Path & "\" & lstFavorites.List(i) For Input As #FileNumber
        Do While Not EOF(FileNumber)
            Input #FileNumber, TempStr$
            If Left(TempStr$, 7) = "FOLDER=" Then Folder$ = Right(TempStr$, Len(TempStr$) - 7)
            If Left(TempStr$, 6) = "TITLE=" Then Title$ = Right(TempStr$, Len(TempStr$) - 6)
            If Left(TempStr$, 4) = "URL=" Then URLf$ = Right(TempStr$, Len(TempStr$) - 4)
            DoEvents
        Loop
    Close #FileNumber
    
    If i <> 0 Then Load frmBrowser.mnuFavoriteT(i)
    frmBrowser.mnuFavoriteT(i).Caption = Title$
    frmBrowser.mnuFavoriteT(i).Tag = URLf$
    DoEvents
Next i

Unload Me
End Sub

Private Sub cmdDelete_Click()

If MsgBox("Remove item from Favorites?", vbOKCancel, "Remove Favorite?") = vbOK Then
    Kill "Favorites\" & lstFavorites.List(lstFav.ListIndex)
    lstFav.RemoveItem (lstFav.ListIndex)
    lstFavorites.Refresh
Else
End If


End Sub

Private Sub cmdRename_Click()
Dim InputR$
InputR$ = InputBox("What would you like to rename it?", "Rename Favorite", lstFav.List(lstFav.ListIndex))
FileNumber = FreeFile
If InputR$ <> vbNullString Then
    Dim fso, OldName As String, NewName As String, TempStr As String, Title$, URLf$, Folder$
    
    OldName$ = lstFavorites.List(lstFav.ListIndex)
    NewName$ = InputR$
    
    Open "Favorites\" & lstFavorites.List(lstFav.ListIndex) For Input As #FileNumber
        Do While Not EOF(FileNumber)
            Input #FileNumber, TempStr$
            If Left(TempStr$, 7) = "FOLDER=" Then Folder$ = TempStr$
            If Left(TempStr$, 4) = "URL=" Then URLf$ = TempStr$
        DoEvents
        Loop
    Close #FileNumber
    
    Open "Favorites\" & lstFavorites.List(lstFav.ListIndex) For Output As #FileNumber
        Print #FileNumber, Folder$ & vbNewLine & "TITLE=" & NewName$ & vbNewLine & URLf$
    Close #FileNumber
    
    lstFav.List(lstFav.ListIndex) = NewName$
    lstFavorites.Refresh
End If

End Sub



Private Sub Form_Load()
FileNumber = FreeFile
Dim TempStr As String
lstFavorites.Path = "Favorites"

For i = 0 To lstFavorites.ListCount - 1
    Open "Favorites\" & lstFavorites.List(i) For Input As #FileNumber
        Do While Not EOF(FileNumber)
        Input #FileNumber, TempStr$
            If Left(TempStr$, 6) = "TITLE=" Then
                lstFav.AddItem Right(TempStr$, Len(TempStr$) - 6)
                Exit Do
            End If
        DoEvents
        Loop
    Close #FileNumber
DoEvents
Next i

End Sub

Private Sub lstFav_Click()
Dim GetFavInfo As String, TempStr As String, GetURL As String, GetMod, GetTitle As String
Set GetMod = CreateObject("Scripting.FileSystemObject")
FileNumber = FreeFile

GetFavInfo$ = "Favorites\" & lstFavorites.List(lstFav.ListIndex)

Open GetFavInfo$ For Input As #FileNumber
    Do While Not EOF(FileNumber)
        Input #FileNumber, TempStr$
        If Left(TempStr$, 6) = "TITLE=" Then GetTitle$ = Right(TempStr$, Len(TempStr$) - 6)
        If Left(TempStr$, 4) = "URL=" Then GetURL$ = Right(TempStr$, Len(TempStr$) - 4)
    DoEvents
    Loop
Close #FileNumber

lblTitle.Caption = GetTitle$
lblURL.Caption = GetURL$
lblMod.Caption = FormatDateTime(GetMod.GetFile(GetFavInfo$).DateLastModified, 1)
End Sub

