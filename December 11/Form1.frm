VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   ScaleHeight     =   10185
   ScaleWidth      =   12630
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Text            =   "http://"
      Top             =   8520
      Width           =   12015
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   255
      Left            =   12000
      TabIndex        =   5
      Top             =   8520
      Width           =   615
   End
   Begin VB.ListBox lst1 
      Height          =   450
      Left            =   0
      TabIndex        =   4
      Top             =   9240
      Width           =   3015
   End
   Begin VB.ListBox lst2 
      Height          =   450
      Left            =   3120
      TabIndex        =   3
      Top             =   9240
      Width           =   3735
   End
   Begin VB.TextBox txtURLx 
      Height          =   285
      Left            =   6960
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   9240
      Width           =   3855
   End
   Begin VB.CheckBox chkPopup 
      Caption         =   "Block Popups"
      Height          =   255
      Left            =   11280
      TabIndex        =   1
      Top             =   8880
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12615
      ExtentX         =   22251
      ExtentY         =   14843
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
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
      Location        =   "http:///"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim breakURL$, breakURLbn$, maskURL$, pageURL$, chkStart As Byte, skipNavigateChk As Byte

Private Sub cmdGo_Click()
breakURL$ = Replace(txtURL.Text, "http://", "")

For i = 0 To lst1.ListCount - 1
    If lst1.List(i) = Left(breakURL$, Len(lst1.List(i))) Then
        wb.Navigate lst2.List(i) & Right(breakURL$, Len(breakURL$) - Len(lst1.List(i)))
        Exit For
    End If
    DoEvents
Next i

End Sub
Private Function setDNS()

Dim dnsInfo$, lst1Item, lst2Item, splitDNS
Open "DNSlst.txt" For Input As #1
Do Until EOF(1)
    Input #1, dnsInfo$
    splitDNS = Split(dnsInfo$, vbTab, -1, 1)
    lst1Item = splitDNS(0)
    lst2Item = splitDNS(1)
    lst1.AddItem lst1Item
    lst2.AddItem lst2Item
    DoEvents
Loop
Close #1


End Function
Private Sub Form_Load()
chkStart = 1
skipNavigateChk = 0
setDNS
getFAV
End Sub

Private Function getFAV()
Dim FAVt$, splitFAV, titleF$, URLf$
i = 0
Open "favorites.dat" For Input As #1
Do Until EOF(1)
    Input #1, FAVt$
    splitFAV = Split(FAVt$, "::", -1, 1)
    titleF$ = splitFAV(0)
    URLf$ = splitFAV(1)
    mnuFavoriteT(i).Caption = titleF$
    mnuFavoriteT(i).Tag = URLf$
    Load mnuFavoriteT(i + 1)
    i = i + 1
    DoEvents
Loop
Close #1

End Function







Private Sub mnuAddFav_Click()

savestr$ = InputBox("DNS Browser will add this url to your favorites: '" & wb.LocationURL & "'" & vbNewLine & "Please designate title.", "Add To Favorites?", wb.LocationName)
If savestr$ <> vbNullString Then
'("Add '" & wb.LocationURL & "' to your favorites?", vbOKCancel, "Add to Favorites?") = vbOK Then
    Open "favorites.dat" For Append As #1
        Print #1, vbNewLine & savestr$ & "::" & wb.LocationURL
    Close #1
    Load mnuFavoriteT(mnuFavoriteT.UBound + 1)
    mnuFavoriteT(mnuFavoriteT.UBound).Caption = savestr$
    mnuFavoriteT(mnuFavoriteT.UBound).Tag = wb.LocationURL
End If
End Sub

Private Sub mnuFavoriteT_Click(Index As Integer)
breakURL$ = Replace(mnuFavoriteT(Index).Tag, "http://", "")
For i = 0 To lst1.ListCount - 1
    If lst1.List(i) = Left(breakURL$, Len(lst1.List(i))) Then
        wb.Navigate lst2.List(i)
        Exit Sub
    End If
    DoEvents
Next i

End Sub

Private Sub txtURL_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call cmdGo_Click
End If
End Sub

Private Sub wb_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

If skipNavigateChk = 1 Then
    skipNavigateChk = 0
    Exit Sub
End If

breakURLbn$ = Replace(URL, "http://", "")
For i = 0 To lst1.ListCount - 1
    If lst1.List(i) = Left(breakURLbn$, Len(lst1.List(i))) Then
        Cancel = True
        wb.Navigate lst2.List(i) & Right(breakURLbn$, Len(breakURLbn$) - Len(lst1.List(i)))
        skipNavigateChk = 1
        Exit For
        Exit Sub
    End If
    
    If lst2.List(i) = Left(breakURLbn$, Len(lst2.List(i))) Then
        Cancel = True
        wb.Navigate lst2.List(i) & Right(breakURLbn$, Len(breakURLbn$) - Len(lst2.List(i)))
        skipNavigateChk = 0
        Exit For
        Exit Sub
    End If
    DoEvents
Next i

End Sub

Private Sub wb_DocumentComplete(ByVal pDisp As Object, URL As Variant)
If chkStart = 1 Then
    chkStart = 0
    Exit Sub
End If

breakURL$ = Replace(URL, "http://", "")
For i = 0 To lst1.ListCount
If lst2.List(i) = Left(breakURL$, Len(lst2.List(i))) Then
    pageURL$ = Right(URL, Len(URL) - Len(lst2.List(i)) - 7)
    txtURL.Text = "http://" & lst1.List(i) & pageURL$
    Exit For
End If
Next i

End Sub

Private Sub wb_NewWindow2(ppDisp As Object, Cancel As Boolean)

If chkPopup.Value = 1 Then
    Cancel = True
Else
    Cancel = False
End If

End Sub

