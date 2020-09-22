Attribute VB_Name = "ModDNS"
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Public Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Public curWB As Integer 'used to hold the current web browser number
Public IntTimeOut As Integer
Public Const vbQuote = """"
'=======================================
'DECLARE SETTINGS FOR PREFERENCES
'=======================================
Public StrHP As String 'homepage string
Public BoolDNS As String, BoolAuto As String, blnAutoFilter As String, bStayTop As String
Public ValHist As Byte, ValDBS As Byte, ValBP As Byte, ValSnd As Byte, ValBKUP As Byte, ResolveAdd As Byte
Public ValEnFilter As String, strFilterPass As String, blnLoginFilter As Byte, ByPassFilter As Byte
Public strUserName As String, StartVal As String, chkTOTD As Byte, ValNotify As Byte



Public xCancel As Byte
' WRITING/READ STUFF; first is writing stuff like [general], second is HOMEPAGE=thing
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public FileNumber As Integer 'FOR OPENING FILES

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Public Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, _
ByVal Y As Long, _
ByVal cX As Long, _
ByVal cY As Long, _
ByVal wFlags As Long) As Long
Public Function SetTopMostWindow(hwnd As Long, TopMost As Boolean) As Long
    If TopMost = True Then
        SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
    Else
        SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
        SetTopMostWindow = True
    End If
End Function

Public Function ReadString(Header As String, strKey As String, strFileLoc As String)
    Dim nChars As Long
    Dim sBuffer As String
    Dim strResult
    sBuffer = String$(255, 0)
    nChars = GetPrivateProfileString(Header$, strKey$, strKey$, sBuffer$, Len(sBuffer$), strFileLoc$)
    strResult = Replace(sBuffer, Chr(0), vbNullString)
    ReadString = strResult
End Function
Public Function WriteString(strHeader As String, strKey As String, strkeyval As String, strFileLoc As String)
    Dim exec As Long
    exec = WritePrivateProfileString(strHeader$, strKey$, strkeyval$, strFileLoc$)
End Function

Public Function DeleteString(strHeader As String, strKey As String, strFileLoc As String)
WritePrivateProfileString strHeader$, strKey$, vbNullString, strFileLoc$
End Function
Public Function refreshDNS()
frmBrowser.txtDebug.Text = frmBrowser.txtDebug.Text & vbNewLine & _
"Refreshing DNS List. Please wait..." & vbNewLine
'clear lists
frmBrowser.lstAH.Clear
frmBrowser.lstAH2.Clear

frmBrowser.lstIP.Clear
frmBrowser.lstIP2.Clear

frmBrowser.lstQZ.Clear
frmBrowser.lstQZ2.Clear

frmBrowser.lstAM.Clear
frmBrowser.lstAM2.Clear

frmBrowser.lstNZ.Clear
frmBrowser.lstNZ2.Clear
'reload dns list
setDNS
End Function
Public Function setDNS()
FileNumber = FreeFile
If frmSplash.Visible <> False Then
    frmSplash.lblState.Caption = "Initializing DNS Database..."
End If

Dim dnsInfo$, lst1Item, lst2Item, splitDNS() As String

chkFExists "DNS Database\Current\A-H.dat", "[DNS Browser - DNS Database]"
Open "DNS Database\Current\A-H.dat" For Input As FreeFile
Do While Not EOF(FileNumber)
    Input #FileNumber, dnsInfo$
    
    If dnsInfo$ <> "[DNS Browser - DNS Database]" And dnsInfo$ <> vbNullString Then
        splitDNS$ = Split(dnsInfo$, "=", -1, 1)
        If UBound(splitDNS) = 1 Then
            lst1Item = splitDNS(0)
            lst2Item = splitDNS(1)
            frmBrowser.lstAH.AddItem lst1Item
            frmBrowser.lstAH2.AddItem lst2Item
        Else
            MsgBox "Error loading entry '" & splitDNS(0) & "'."
        End If
    Else
    End If
    
    DoEvents
Loop
Close #FileNumber


chkFExists "DNS Database\Current\I-P.dat", "[DNS Browser - DNS Database]"
Open "DNS Database\Current\I-P.dat" For Input As #FileNumber
Do While Not EOF(FileNumber)
    Input #FileNumber, dnsInfo$
    
    If dnsInfo$ <> "[DNS Browser - DNS Database]" And dnsInfo$ <> vbNullString Then
        splitDNS$ = Split(dnsInfo$, "=", -1, 1)
        If UBound(splitDNS) = 1 Then
            lst1Item = splitDNS(0)
            lst2Item = splitDNS(1)
            frmBrowser.lstIP.AddItem lst1Item
            frmBrowser.lstIP2.AddItem lst2Item
        Else
            MsgBox "Error loading entry '" & splitDNS(0) & "'."
        End If
    Else
    End If
    
    DoEvents
Loop
Close #FileNumber


chkFExists "DNS Database\Current\Q-Z.dat", "[DNS Browser - DNS Database]"
Open "DNS Database\Current\Q-Z.dat" For Input As #FileNumber
Do While Not EOF(FileNumber)
    Input #FileNumber, dnsInfo$
    
    If dnsInfo$ <> "[DNS Browser - DNS Database]" And dnsInfo$ <> vbNullString Then
        splitDNS$ = Split(dnsInfo$, "=", -1, 1)
        If UBound(splitDNS) = 1 Then
            lst1Item = splitDNS(0)
            lst2Item = splitDNS(1)
            frmBrowser.lstQZ.AddItem lst1Item
            frmBrowser.lstQZ2.AddItem lst2Item
        Else
            MsgBox "Error loading entry '" & splitDNS(0) & "'."
        End If
    Else
    End If
    
    DoEvents
Loop
Close #FileNumber


chkFExists "DNS Database\Current\WWW_A-M.dat", "[DNS Browser - DNS Database]"
Open "DNS Database\Current\WWW_A-M.dat" For Input As #FileNumber
Do While Not EOF(FileNumber)
    Input #FileNumber, dnsInfo$
    
    If dnsInfo$ <> "[DNS Browser - DNS Database]" And dnsInfo$ <> vbNullString Then
        splitDNS$ = Split(dnsInfo$, "=", -1, 1)
        If UBound(splitDNS) = 1 Then
            lst1Item = splitDNS(0)
            lst2Item = splitDNS(1)
            frmBrowser.lstAM.AddItem lst1Item
            frmBrowser.lstAM2.AddItem lst2Item
        Else
            MsgBox "Error loading entry '" & splitDNS(0) & "'."
        End If
    Else
    End If
    
    DoEvents
Loop
Close #FileNumber


chkFExists "DNS Database\Current\WWW_N-Z.dat", "[DNS Browser - DNS Database]"
Open "DNS Database\Current\WWW_N-Z.dat" For Input As #FileNumber
Do While Not EOF(FileNumber)
    Input #FileNumber, dnsInfo$
    
    If dnsInfo$ <> "[DNS Browser - DNS Database]" And dnsInfo$ <> vbNullString Then
        splitDNS$ = Split(dnsInfo$, "=", -1, 1)
        If UBound(splitDNS) = 1 Then
            lst1Item = splitDNS(0)
            lst2Item = splitDNS(1)
            frmBrowser.lstNZ.AddItem lst1Item
            frmBrowser.lstNZ2.AddItem lst2Item
        Else
            MsgBox "Error loading entry '" & splitDNS(0) & "'."
        End If
    Else
    End If
    
    DoEvents
Loop
Close #FileNumber

If frmSplash.Visible <> False Then
    frmSplash.lblState.Caption = "Loading Complete!"
    Unload frmSplash
End If

DebugInfo
End Function
Public Function chkFExists(File As String, Header As String)
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
FileNumber = FreeFile
If fso.FileExists(File) Then
    Else
    frmBrowser.txtDebug.Text = frmBrowser.txtDebug.Text & vbNewLine & File & _
    " does not exist. Creating file..."
    
    fso.CreateTextFile (File)
    Pause (0.3) 'only pause if creating text file! why pause else wise!?!
    Open File For Output As #FileNumber
        Print #FileNumber, Header
    Close #FileNumber
    frmBrowser.txtDebug.Text = frmBrowser.txtDebug.Text & vbNewLine & File & _
    " created! Please re-load browser before use."
End If

End Function

Public Function Pause(ByVal TimeToWait As Long) 'Time In seconds
    Dim EndTime As Long
    EndTime = GetTickCount + TimeToWait * 1000 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds


    Do Until GetTickCount > EndTime


        DoEvents
    Loop
End Function
Public Function TimeOut(ByVal TimeToWait As Long) 'Time In seconds
    Dim EndTime As Long
    EndTime = GetTickCount + TimeToWait * 1 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds


    Do Until GetTickCount > EndTime


        DoEvents
    Loop
End Function
Public Function DebugInfo()
frmBrowser.txtDebug.Text = frmBrowser.txtDebug.Text & vbTab & vbTab & vbTab & _
vbTab & vbTab & "Time: " & Format(time, hh, mm, ss) & _
vbNewLine & vbNewLine & "DNS List Loaded: " & frmBrowser.lstAH.ListCount + frmBrowser.lstIP.ListCount _
+ frmBrowser.lstQZ.ListCount & " entries." & vbNewLine & "Favorites Loaded: " & _
frmBrowser.mnuFavoriteT.Count & " entries." & vbNewLine
End Function

Public Function GetWHistory()
FileNumber = FreeFile
If frmSplash.Visible <> False Then
    frmSplash.lblState.Caption = "Loading History..."
End If

Dim hist$

chkFExists "DATA\history.dat", "//DNS Browser - History//"
Open "DATA\history.dat" For Input As #FileNumber
Do While Not EOF(FileNumber)
    Input #FileNumber, hist$
    
    If hist$ <> "//DNS Browser - History//" And hist$ <> vbNullString Then
        frmBrowser.cboURL.AddItem hist$
    Else
    End If
    DoEvents
Loop
Close #FileNumber

If frmSplash.Visible <> False Then
    frmSplash.lblState.Caption = "History Loaded."
End If
End Function
Public Function SetTrial()
Dim DateStr$, InputStr$, fso, UseStr$, CurMDY
Dim i As Long
FileNumber = FreeFile
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists("H:\:DNS" & App.Revision & ".dat") Then

    Open "H:\:DNS" & App.Revision & ".dat" For Input As #FileNumber
        Do While Not EOF(FileNumber)
            Input #FileNumber, InputStr$
                If Len(InputStr$) > 2 Then
                    DateStr$ = InputStr$
                Else
                    UseStr$ = InputStr$
                End If
        DoEvents
        Loop
    Close #FileNumber
    
    UseStr$ = UseStr$ + 1
    Open "H:\:DNS" & App.Revision & ".dat" For Output As #FileNumber
        Print #FileNumber, DateStr$ & vbNewLine & UseStr$
    Close #FileNumber
    
    CurMDY = Format(Now, "ddddd")
    
    If UseStr >= 30 Then
        MsgBox "Sorry, the trial for this program has ended. Please visit 'http://www.dnsbrowser.com'" & _
        " for purchase info, or contact seoulxkorean@yahoo.com.", vbCritical, "Program used '30' times."
        Unload frmSplash
        Unload frmBrowser
    End If
    
    If DateDiff("d", DateStr, CurMDY) >= 10 Or DateDiff("d", DateStr$, CurMDY) < 0 Then
        MsgBox "Sorry, the trial for this program has ended. Please visit 'http://www.dnsbrowser.com'" & _
        " for purchase info, or contact seoulxkorean@yahoo.com.", vbCritical, "10 Day Trial has ended."
        Unload frmSplash
        Unload frmBrowser
    End If
    
Else
    fso.CreateTextFile "H:\:DNS" & App.Revision & ".dat"
        Open "H:\:DNS" & App.Revision & ".dat" For Output As #FileNumber
            Print #FileNumber, Format(Now, "ddddd") & vbNewLine & "1"
        Close #FileNumber
    
End If
End Function
Public Function CheckDayToUpdate()
Dim TempStr$, NowDate$
FileNumber = FreeFile

If Format(Now, "w") = "3" Then
    NowDate$ = Format(Now, "short date")
    Open "DATA\updatechk.dat" For Input As #FileNumber
        Do While Not EOF(FileNumber)
            Input #FileNumber, TempStr$
            If TempStr$ = NowDate$ Then Exit Function
        DoEvents
        Loop
    Close #FileNumber
Else
    Exit Function
End If

    Open "DATA\updatechk.dat" For Append As #FileNumber
        Print #FileNumber, NowDate$
    Close #FileNumber
    If MsgBox("DNS Browser Auto-Update has started. DNS Browser will check for latest version..." & vbNewLine & _
            "Do you wish to continue?", vbYesNo, "Auto-Update File Check") = vbYes Then
        CheckForUpdates ("C:\Documents and Settings\Daniel\Desktop\DNS Project Package\DNSBrowser.exe")
    Else
        MsgBox "Auto-Update Canceled. Auto-Update will check again in 7 days.", vbInformation
    End If
    
End Function
Public Function CheckDayToBackup()
Dim TempStr$, NowDate$
FileNumber = FreeFile

If Format(Now, "w") = "3" Then
    NowDate$ = Format(Now, "short date")
    Open "DATA\bkupchk.dat" For Input As #FileNumber
        Do While Not EOF(FileNumber)
        Input #FileNumber, TempStr$
        If TempStr$ = NowDate$ Then Exit Function
        DoEvents
        Loop
    Close #FileNumber
Else
    Exit Function
End If

Open "DATA\bkupchk.dat" For Append As #FileNumber
    Print #FileNumber, NowDate$
Close #FileNumber

If MsgBox("DNS Browser has been set to automatically backup your files." & vbNewLine & _
            "Would you like to perform backup?", vbYesNo, "Auto-Backup Query") = vbYes Then
    PerformBackup
Else
    MsgBox "Auto-Backup Canceled", vbInformation, "Action Canceled"
End If
End Function
Public Function PerformBackup()
Dim bkupdata
Set bkupdata = CreateObject("Scripting.FileSystemObject")

If bkupdata.FolderExists("DATA\Backup") Then
    bkupdata.DeleteFolder ("DATA\Backup")
Else
End If

bkupdata.CopyFolder "DATA", "DATA\" & Format(Now, "short date")

MsgBox "Auto-Backup Complete!", vbInformation, "Auto-Backup Successful!"
End Function
Public Function CheckForUpdates(FileLoc As String)
Dim ChkU, VersionInfo As String
Set ChkU = CreateObject("Scripting.FileSystemObject")


If ChkU.FileExists(FileLoc) Then

    VersionInfo$ = ChkU.GetFileVersion(FileLoc)
    Dim MyVersion As String
    MyVersion$ = App.Major & "." & App.Minor & "." & App.Revision

    If VersionInfo$ <> MyVersion$ Then
        If MsgBox("You currently have an outdated version of DNS Browser." & vbCrLf & _
            "Would you like to install the Update?", vbYesNo, "Update Alert!") = vbYes Then
            
            If ChkU.FileExists("autoupdate.exe") Then
                Shell "autoupdate.exe", vbNormalFocus
                Unload frmBrowser
            Else
                MsgBox "Critical File Missing: Auto Update File is missing. Download it manually at" & vbCrLf & _
                "'" & Left(FileLoc, Len(FileLoc) - 14) & "'."
            End If
            
        Else
        
        End If
    Else
        MsgBox "You're Browser is currently up to date.", vbInformation, "DNS Browser Info"
    End If
End If
End Function
Public Function LoadOptSettings()
Dim nChars As Long
Dim sBuffer As String
Dim lBuffer As String
FileNumber = FreeFile
sBuffer$ = String$(255, 0)
nChars = GetUserName(sBuffer$, Len(sBuffer$))
strUserName$ = Replace(sBuffer$, Chr(0), "")

If frmSplash.Visible <> False Then
    frmSplash.lblState.Caption = "Loading Settings..."
End If

Dim TempStr As String
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists("DATA\popaccess.dat") <> True Then
    fso.CreateTextFile (App.Path & "\" & "DATA\popaccess.dat")
Else
frmOptions.lstPB.Clear

    Open "DATA\popaccess.dat" For Input As #FileNumber
        Do While Not EOF(FileNumber)
            Input #FileNumber, TempStr$
            If TempStr$ <> vbNullString Then frmOptions.lstPB.AddItem TempStr$
        Loop
    Close #FileNumber
End If


If fso.FileExists("DATA\options.dat") <> True Then
    fso.FileCopy "DATA\default.dat", "DATA\options.dat"
End If

'=========================================================
'Description: Read all settings from options.dat file
'=========================================================
StrHP$ = ReadString("General Options", "HOMEPAGE", "DATA\options.dat")
ValHist = ReadString("General Options", "DISABLEHIST", "DATA\options.dat")
BoolDNS = ReadString("General Options", "DNSMODE", "DATA\options.dat")
ValDBS = ReadString("DNS Settings", "DBSPECIFIC", "DATA\options.dat")
IntTimeOut = ReadString("DNS Settings", "TIMEOUT", "DATA\options.dat")
ValBP = ReadString("Privacy Settings", "BLOCKPOPUP", "DATA\options.dat")
ValSnd = ReadString("Privacy Settings", "PLAYSOUND", "DATA\options.dat")
ValNotify = ReadString("Privacy Settings", "NOTIFY", "DATA\options.dat")
BoolAuto = ReadString("Miscellaneous", "AUTOUPDATE", "DATA\options.dat")
ValBKUP = ReadString("Miscellaneous", "BACKUPCONF", "DATA\options.dat")
bStayTop = ReadString("Miscellaneous", "STAYONTOP", "DATA\options.dat")
chkTOTD = ReadString("Miscellaneous", "TOTD", "DATA\options.dat")
frmBrowser.sldTransparency.Value = ReadString("Miscellaneous", "FORMTRANSPARENCY", "DATA\options.dat")
frmOptions.sldTransparency.Value = ReadString("Miscellaneous", "OPTTRANSPARENCY", "DATA\options.dat")

frmTip.chkLoadTipsAtStartup.Value = chkTOTD
frmBrowser.mnuStayTop.Checked = bStayTop
If bStayTop = True Then SetTopMostWindow frmBrowser.hwnd, True
With frmOptions
    .txtHomePage.Text = StrHP$
    .chkDisableH.Value = ValHist
    .OptIB(0).Value = BoolDNS
        If .OptIB(0).Value = 0 Then .OptIB(1).Value = 1
    .OptUd(ValDBS - 1).Value = True
    .chkEnable.Value = ValBP
    .chkPopSnd.Value = ValSnd
    .chkPopNotify.Value = ValNotify
    .OptUp(0).Value = BoolAuto
    .txtDNSTO.Text = IntTimeOut
        If .OptUp(0).Value = 0 Then .OptUp(1).Value = 1
    .chkBackup.Value = ValBKUP
End With

If ValBP = 0 Then
    frmBrowser.mnuPopupBlockSet.Caption = "Turn On Pop-up Blocker"
Else
    frmBrowser.mnuPopupBlockSet.Caption = "Turn Off Pop-up Blocker"
End If

End Function
Public Function ApplySettings()
Dim DBSVal As String
Dim i As Long
FileNumber = FreeFile
Open "DATA\popaccess.dat" For Output As #FileNumber
    For i = 0 To frmOptions.lstPB.ListCount - 1
        Print #FileNumber, frmOptions.lstPB.List(i)
    DoEvents
    Next i
Close #FileNumber

'GENERAL OPTIONS
'HOMEPAGE: URL SPECIFIED;
'DISABLEHIST: TRUE/FALSE
'DNS MODE: TRUE/FALSE

'DNS SETTINGS
'DBSPECIFIC: 1 - AUTOUPDATE; 2 - NEVER UPDATE; 3 - PROMPT USER UPDATE;

'PRIVACY SETTINGS
'BLOCKPOPUP: 0 - NO; 1 - YES;

'MISCELLANEOUS OPTIONS
'AUTOUPDATE: TRUE/FALSE
'BACKUPCONF: 0 - NO; 1 - YES;

For i = 0 To 2
    If frmOptions.OptUd.Item(i).Value = True Then
        DBSVal = i + 1
        Exit For
    End If
DoEvents
Next i

If frmOptions.chkEnable = True Then
    frmBrowser.mnuPopupBlockSet.Caption = "Turn Off Pop-up Blocker"
Else
    frmBrowser.mnuPopupBlockSet.Caption = "Turn On Pop-up Blocker"
End If
Dim strSetLoc As String
strSetLoc = "DATA\options.dat"

'=========================================================
'Description: Write all settings to options.dat file
'=========================================================
WriteString "General Options", "HOMEPAGE", frmOptions.txtHomePage.Text, strSetLoc$
WriteString "General Options", "DISABLEHIST", frmOptions.chkDisableH.Value, strSetLoc$
WriteString "General Options", "DNSMODE", frmOptions.OptIB.Item(0).Value, strSetLoc$
WriteString "DNS Settings", "DBSPECIFIC", DBSVal, strSetLoc$
WriteString "DNS Settings", "TIMEOUT", frmOptions.txtDNSTO.Text, strSetLoc$
WriteString "Privacy Settings", "BLOCKPOPUP", frmOptions.chkEnable.Value, strSetLoc$
WriteString "Privacy Settings", "PLAYSOUND", frmOptions.chkPopSnd.Value, strSetLoc$
WriteString "Privacy Settings", "NOTIFY", frmOptions.chkPopNotify.Value, strSetLoc$
WriteString "Miscellaneous", "AUTOUPDATE", frmOptions.OptUp(0).Value, strSetLoc$
WriteString "Miscellaneous", "BACKUPCONF", frmOptions.chkBackup.Value, strSetLoc$
WriteString "Miscellaneous", "OPTTRANSPARENCY", frmOptions.sldTransparency.Value, strSetLoc$
End Function

Public Function CheckEncryptionKey()

Dim fso
Dim nChars As Long
Dim sBuffer As String
Dim lBuffer As String
Dim strUserComp As String
sBuffer$ = String$(255, 0)
nChars = GetUserName(sBuffer$, Len(sBuffer$))
strUserComp$ = Replace(sBuffer$, Chr(0), "")
Dim strKey As String, strMatchKey As String

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists("DATA\key.dat") <> True Then
    MsgBox "You have an invalid key! You cannot run this program.", vbCritical, "Invalid Key"
    CloseApplication
End If

Open "DATA\key.dat" For Input As #FileNumber
    Do While Not EOF(FileNumber)
        Input #FileNumber, TempStr$
            strKey$ = strKey$ & TempStr$
    Loop
Close #FileNumber

DecryptKey strKey$, strMatchKey$, "59.853.36.6"
DecryptKey strMatchKey$, strMatchKey$, "59.853.36.6"
DecryptKey strMatchKey$, strMatchKey$, "59.853.36.6"
DecryptKey strMatchKey$, strMatchKey$, "59.853.36.6"
If strMatchKey$ <> strUserComp$ Then
    MsgBox "You have an invalid key! You cannot run this program.", vbCritical, "Invalid Key"
    CloseApplication
Else
End If

End Function
Public Function EncryptKey(EncryptStr As String, EncryptLoc As String, UniqueID As String)
'Encryption Stuff
Dim EncryptL1 As String
Dim EncryptL3 As String
Dim EncryptL4 As String
Dim EncryptL2A
Dim EncryptL2l As Integer
Dim EncryptL2r As Integer
Dim L2BitShft As String
Dim L2BitShftP As String
Dim EncryptA As String
Dim EncryptCom As String
Dim EncryptChar As String
Dim L2BitShftC As String
Dim EncryptL4R As String
Dim GRndBitShft As String
Dim GRndXorVal As String


'first make sure all variables are empty...
    EncryptL4R = vbNullString
    EncryptL4 = vbNullString
    EncryptL3 = vbNullString
    L2BitShftP = vbNullString
    L2BitShft = vbNullString
    EncryptL1 = vbNullString
'ok here are the steps we will be doing:
'convert to ascii
'do random bitshifting depending on client ip; extract first 3 and last 192.168.0.8
    EncryptA = EncryptStr
    EncryptCom = EncryptA ' Split up the text to command & syntax
'EncryptSyntax = EncryptA(1) ' splitting
'OK now we need to make all chars within the "EncryptCom" into ascii and xor first set of IP

'we will need to split the IP
    EncryptL2A = Split(UniqueID$, ".", -1, 1)
    EncryptL2l = EncryptL2A(0)
    EncryptL2r = EncryptL2A(3)
    Randomize
    RndXorVal = Int(Rnd * 256)
    For i = 1 To Len(EncryptCom)
        EncryptChar = Mid$(EncryptCom, i, 1)
        If EncryptChar = " " Then
            EncryptL1 = EncryptL1 & "x;" 'TEST THIS OMG OMG OMG
            Else
            If i Mod 2 = 0 Then
            EncryptL1 = EncryptL1 & (Asc(EncryptChar) Xor EncryptL2A(0)) & ";"
            Else
            EncryptL1 = EncryptL1 & (Asc(EncryptChar) Xor RndXorVal) & ";"
            End If
        End If
        DoEvents
    Next i
'Now comes the difficult part. Bitshifting dependent on IP
'ABCDEFG
    Randomize
    RndBitShft = Int(Rnd * 256)
    For i = 1 To RndBitShft 'Left$(EncryptL2l, 1) 'bitshift w/ first number (move rightmost to Left$most)
        If i = 1 Then
            L2BitShft = Right(EncryptL1, 1)
            L2BitShftP = L2BitShft & Left$(EncryptL1, Len(EncryptL1) - 1)
            Else
            L2BitShft = Right(L2BitShftP, 1)
            L2BitShftP = L2BitShft & Left$(L2BitShftP, Len(L2BitShftP) - 1)
        End If
        DoEvents
    Next i
'bitshift two will be taking center numbers and put them on either end
    If Len(L2BitShftP) Mod 2 = 0 Then
        L2BitShftC = Mid$(L2BitShftP, Len(L2BitShftP) \ 2 - 1, 2)
        L2BitShftP = Left$(L2BitShftC, 1) & Left$(L2BitShftP, Len(L2BitShftP) \ 2 - 2) & Right(L2BitShftP, Len(L2BitShftP) \ 2) & Right(L2BitShftC, 1)
    Else

    End If

'End If
'ABCDEFGHI
'next bitshift same as first
    For i = 1 To Mid$(EncryptL2l, 2, 1) 'bit shift w/ second number
        L2BitShft = Right(L2BitShftP, 1)
        L2BitShftP = L2BitShft & Left$(L2BitShftP, Len(L2BitShftP) - 1)
        DoEvents
    Next i
'Now we will take the length of EncryptL2r and determine how many times we will split our product with
'a random char (&) random times
Dim Addr
    Randomize

    For i = 1 To Len(EncryptL2r) + Int(Rnd * 5 + 1)
        Randomize
        Addr = Int(Rnd * Len(L2BitShftP) + 1)
        If i = 1 Then
            EncryptL3 = Left$(L2BitShftP, Addr) & "%" & Right(L2BitShftP, Len(L2BitShftP) - Addr)
            Else
            EncryptL3 = Left$(EncryptL3, Addr) & "%" & Right(EncryptL3, Len(EncryptL3) - Addr)
        End If
        DoEvents
    Next i
'Simple Reverse
    EncryptL4R = StrReverse(EncryptL3)
'Simple Replacement Scheme
'EncryptL4
    EncryptL4 = Replace(EncryptL4R, "0", "ÿ")
    EncryptL4 = Replace(EncryptL4, "1", "Ä")
    EncryptL4 = Replace(EncryptL4, "2", "Ñ")
    EncryptL4 = Replace(EncryptL4, "3", "æ")
    EncryptL4 = Replace(EncryptL4, "4", "ò")
    EncryptL4 = Replace(EncryptL4, "5", "§")
    EncryptL4 = Replace(EncryptL4, "6", "¼")
    EncryptL4 = Replace(EncryptL4, "7", "ü")
    EncryptL4 = Replace(EncryptL4, "8", "µ")
    EncryptL4 = Replace(EncryptL4, "9", "¤")
    If Len(EncryptL4) Mod 2 = 0 Then
        EncryptL4 = Left$(EncryptL4, Len(EncryptL4) \ 2 - 1) & Chr(RndXorVal) & Chr(RndBitShft) & Right(EncryptL4, Len(EncryptL4) \ 2 + 1)
    Else
        EncryptL4 = Chr(RndXorVal) & EncryptL4 & Chr(RndBitShft)
    End If
    EncryptLoc = EncryptL4
End Function

Public Function DecryptKey(DecryptStr As String, DecryptLoc As String, UniqueID As String)
On Error GoTo InvalidKey:
Dim DecryptL1 As String
Dim DecryptL4 As String
Dim DecryptLX As String
Dim D2BitShft As String
Dim D2BitShftP As String
Dim D2BitShftC As String
Dim D3BitShft As String
Dim RndBitShft As Long
Dim RndXorVal As Integer
Dim DecryptL2A
Dim DecryptL2l As Integer
Dim DecryptL2r As Integer
Dim D1BitShft As String
Dim D1BitShftP As String
Dim i As Long

    D2BitShftP = vbNullString
    D3BitShft = vbNullString
    D1BitShftP = vbNullString
    DecryptLX = vbNullString
    
    DecryptL4 = DecryptStr
    If Len(DecryptL4) Mod 2 = 0 Then
        GRndXorVal = Asc(Mid$(DecryptL4, Len(DecryptL4) \ 2 - 1, 1))
        GRndBitShft = Asc(Mid$(DecryptL4, Len(DecryptL4) \ 2, 1))
        DecryptL4 = Left$(DecryptL4, Len(DecryptL4) \ 2 - 2) & Right$(DecryptL4, Len(DecryptL4) \ 2)
    Else
        GRndBitShft = Asc(Right$(DecryptL4, 1))
        GRndXorVal = Asc(Left$(DecryptL4, 1))
        DecryptL4 = Mid$(DecryptL4, 2, Len(DecryptL4) - 2)
    End If
    DecryptL4 = Replace(DecryptL4, "ÿ", "0")
    DecryptL4 = Replace(DecryptL4, "Ä", "1")
    DecryptL4 = Replace(DecryptL4, "Ñ", "2")
    DecryptL4 = Replace(DecryptL4, "æ", "3")
    DecryptL4 = Replace(DecryptL4, "ò", "4")
    DecryptL4 = Replace(DecryptL4, "§", "5")
    DecryptL4 = Replace(DecryptL4, "¼", "6")
    DecryptL4 = Replace(DecryptL4, "ü", "7")
    DecryptL4 = Replace(DecryptL4, "µ", "8")
    DecryptL4 = Replace(DecryptL4, "¤", "9")
    DecryptL1 = Replace(DecryptL4, "%", "") 'take out all useless info - the % char
'reverse it back derr
    DecryptL1 = StrReverse(DecryptL1)
'now we need to "UN"-bitshift using the second number in the IP
    DecryptL2A = Split(UniqueID$, ".", -1, 1)
    DecryptL2l = DecryptL2A(0)
    DecryptL2r = DecryptL2A(3)
    For i = 1 To Mid$(DecryptL2l, 2, 1) 'bit shift w/ second number
        D2BitShft = Left$(DecryptL1, 1)
        If i = 1 Then
            D2BitShftP = Right$(DecryptL1, Len(DecryptL1) - 1) & D2BitShft
            Else
            D2BitShft = Left$(D2BitShftP, 1)
            D2BitShftP = Mid$(D2BitShftP, 2, Len(D2BitShftP) - 1) & D2BitShft
        End If
        DoEvents
    Next i
'CDEFGAB
    If Len(D2BitShftP) Mod 2 = 0 Then
        D2BitShftC = Left$(D2BitShftP, 1) & Right(D2BitShftP, 1)
        D3BitShft = Mid$(D2BitShftP, 2, Len(D2BitShftP) \ 2 - 2) & D2BitShftC & Mid$(D2BitShftP, Len(D2BitShftP) \ 2, Len(D2BitShftP) \ 2)
    Else
    D3BitShft = D2BitShftP
    
    End If
'D3BitShft = Left$(D3BitShft, Len(D3BitShft) - 1)
'test if works
    For i = 1 To GRndBitShft 'Left$(EncryptL2l, 1) 'bit shift w/ first
        If i = 1 Then
            D1BitShft = Left$(D3BitShft, 1)
            D1BitShftP = Mid$(D3BitShft, 2, Len(D3BitShft) - 1) & D1BitShft
            Else
            D1BitShft = Left$(D1BitShftP, 1)
            D1BitShftP = Mid$(D1BitShftP, 2, Len(D1BitShftP) - 1) & D1BitShft
        End If
        DoEvents
    Next i
'OK SPLIT TIME
Dim AscSplit() As String
Dim ALength
    AscSplit$ = Split(D1BitShftP, ";", -1, 1)

    For i = 0 To UBound(AscSplit) - 1
        If AscSplit(i) = "" Then
            Else
            If AscSplit(i) = "x" Then
                DecryptLX = DecryptLX & " "
                Else
                If Left$(AscSplit(i), 1) = "x" Or Right$(AscSplit(i), 1) = "x" Or Mid$(Asc(i), 2, 1) = "x" Then
                    DecryptLX = DecryptLX & " "
                    AscSplit(i) = Replace(AscSplit(i), "x", "1")
                    Else
                    If i Mod 2 = 0 Then
                        DecryptLX = DecryptLX & Chr((AscSplit(i) Xor GRndXorVal))
                    Else
                        If AscSplit(i) > 256 Then
                            DecryptLX = DecryptLX
                        Else
                            DecryptLX = DecryptLX & Chr((AscSplit(i) Xor DecryptL2A(0)))
                        End If
                    End If
                    End If
                End If
                ALength = ALength + Len(AscSplit(i))
            End If
        DoEvents
    Next i
    DecryptLoc = DecryptLX
Exit Function
InvalidKey:
    MsgBox "You have an invalid key! You cannot run this program.", vbCritical, "Invalid Key"
    CloseApplication
End Function

Public Function CloseApplication()
    WriteString "Miscellaneous", "FORMTRANSPARENCY", frmBrowser.sldTransparency.Value, "DATA\options.dat"
    
    Unload frmAbout
    Unload frmDeleteFiles
    Unload frmDNS
    Unload frmFavorites
    Unload frmOptions
    Unload frmSplash
    Unload frmTip
    Unload frmFilterByPass
    Unload frmBrowser
    
    Set frmAbout = Nothing
    Set frmBrowser = Nothing
    Set frmDeleteFiles = Nothing
    Set frmDNS = Nothing
    Set frmFavorites = Nothing
    Set frmOptions = Nothing
    Set frmFilterByPass = Nothing
    Set frmSplash = Nothing
    Set frmTip = Nothing
    
    End
 
End Function
Public Function LoadFilterList()
frmOptions.lstFilter.Clear
Dim TempStr As String, strFilters As String, splitFilters() As String
Dim i As Long
FileNumber = FreeFile

Open "DATA\filterkey.dat" For Input As #FileNumber
    Do While Not EOF(FileNumber)
        Input #FileNumber, TempStr$
        strFilters$ = strFilters$ & TempStr$
    DoEvents
    Loop
Close #FileNumber
splitFilters = Split(strFilters$, " ", -1, 1)

For i = 0 To UBound(splitFilters)
    frmOptions.lstFilter.AddItem splitFilters(i)
    DoEvents
Next i

End Function

Public Function LoadFilterSettings()
Dim TempStr As String, strFilterFile As String, splitFilterForDecrypt() As String
Dim strFilterSettings As String, ParseFilterSettings() As String, i As Long

    ValEnFilter = ReadString("Filter Settings", "ENABLEFILTER", "DATA\options.dat")
    strFilterPass$ = ReadString("Filter Settings", "PASSWORD", "DATA\options.dat")
    blnAutoFilter = ReadString("Filter Settings", "AUTOFILTER", "DATA\options.dat")
    
    DecryptKey ValEnFilter, ValEnFilter, "37.285.17.32"
    DecryptKey strFilterPass$, strFilterPass$, "68.158.53.60"
    DecryptKey blnAutoFilter, blnAutoFilter, "45.192.45.21"

    frmOptions.chkFilter.Value = ValEnFilter
    frmOptions.OptAuto(0).Value = blnAutoFilter
    
'blnLoginFilter = 1


End Function

Public Function FilterCheck(URL As String)
If frmBrowser.wb(curWB).Busy <> True Then
If frmOptions.chkFilter.Value = 1 Then
        For i = 1 To Len(URL)
            For j = 0 To frmOptions.lstFilter.ListCount - 1
                If InStr(i, URL, frmOptions.lstFilter.List(j)) Then
                    If frmOptions.OptAuto(0).Value = True Then
                        MsgBox "Website has been filtered out by Filter Control.", vbCritical, "Filter Control Alert"
                        xCancel = 1
                        Exit Function
                    Else
                        frmFilterByPass.lblAllow.Caption = "Allow: '" & URL & "' and add to list?"
                        frmFilterByPass.Show vbModal
                            If ByPassFilter <> 1 Then
                                xCancel = 1
                                MsgBox "Website has been filtered out by Filter Control.", vbCritical, "Filter Control Alert"
                            Else
                                'Do Until frmBrowser.wb(curWB).Busy <> True
                                xCancel = 0
                                'DoEvents
                                'Loop
                            Exit Function
                            End If
                    End If
                End If
                DoEvents
            Next j
        Next i
Else

End If
End If
End Function


