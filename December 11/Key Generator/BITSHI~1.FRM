VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   ScaleHeight     =   9885
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9960
      Top             =   9360
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4680
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   9480
      Width           =   2415
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3360
      Top             =   9000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtDecrypt 
      Height          =   8775
      Left            =   6240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   120
      Width           =   6495
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "Decrypt"
      Height          =   375
      Left            =   10080
      TabIndex        =   3
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set IP"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "Encrypt"
      Height          =   375
      Left            =   11400
      TabIndex        =   1
      Top             =   9000
      Width           =   1335
   End
   Begin VB.TextBox txtEncrypt 
      Height          =   8775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "BITSHI~1.frx":0000
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status: Waiting..."
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   9600
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   8280
      TabIndex        =   6
      Top             =   9000
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   9000
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Encryption Stuff
Private EncryptL1 As String
Private EncryptL3 As String
Private EncryptL4 As String
Private EncryptL2A
Private EncryptL2l As Integer
Private EncryptL2r As Integer
Private L2BitShft As String
Private L2BitShftP As String
Private EncryptA As String
Private EncryptCom As String
Private EncryptChar As String
Private L2BitShftC As String
Private EncryptL4R As String
Private GRndBitShft As String
Private GRndXorVal As String
'Decryption stuff
Private DecryptL1 As String
Private DecryptL4 As String
Private DecryptLX As String
Private D2BitShft As String
Private D2BitShftP As String
Private D2BitShftC As String
Private D3BitShft As String
Private RndBitShft As Long
Private RndXorVal As Integer
Private DecryptL2A
Private DecryptL2l As Integer
Private DecryptL2r As Integer
Private D1BitShft As String
Private D1BitShftP As String
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Dim EncryptTime1, EncryptTime2 As Long
Dim DecryptTime1, DecryptTime2 As Long
Dim i As Long

Private Sub cmdDecrypt_Click()
DecryptTime1 = timeGetTime

'time to decrypt YEAH!
    lblStatus.Caption = "Status: Decrypting File..."
    Call DecryptText(txtEncrypt, txtDecrypt)

End Sub
Private Sub cmdEncrypt_Click()
EncryptTime1 = timeGetTime

    lblStatus.Caption = "Status: Encrypting File..."
    Call EncryptText(txtEncrypt, txtEncrypt)

End Sub
Sub DecryptText(DecryptStr As TextBox, DecryptLoc As TextBox)

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
    DecryptL2A = Split("59.853.36.6", ".", -1, 1)
    DecryptL2l = DecryptL2A(0)
    DecryptL2r = DecryptL2A(3)
    For i = 1 To Mid$(DecryptL2l, 2, 1) 'bit shift w/ second number
        D2BitShft = Left$(DecryptL1, 1)
        If i = 1 Then
            D2BitShftP = Right(DecryptL1, Len(DecryptL1) - 1) & D2BitShft
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
Dim AscSplit
Dim ALength
    AscSplit = Split(D1BitShftP, ";", -1, 1)

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
                            DecryptLX = DecryptLX & Chr((AscSplit(i) Xor EncryptL2A(0)))
                        End If
                    End If
                    End If
                End If
                ALength = ALength + Len(AscSplit(i))
            End If
        DoEvents
    Next i
    DecryptLoc = DecryptLX

DecryptTime2 = timeGetTime

    lblStatus.Caption = "Status: File Decrypted in " & (DecryptTime2 - DecryptTime1) \ 100 & " seconds."
End Sub
Sub EncryptText(EncryptStr As TextBox, EncryptLoc As TextBox)


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
    EncryptL2A = Split("59.853.36.6", ".", -1, 1)
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
    Label2.Caption = Len(EncryptL1)
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
    
EncryptTime2 = timeGetTime

    lblStatus.Caption = "Status: File Encrypted in " & (EncryptTime2 - EncryptTime1) \ 1000 & " seconds."

End Sub
Private Sub Command1_Click()

    Winsock1.RemoteHost = "raven"
    Winsock1.RemotePort = "445"
    Winsock1.Connect

End Sub
Private Sub Winsock1_Click()


End Sub

Private Sub txtEncrypt_Change()

    Label1.Caption = Len(txtEncrypt.Text)

End Sub
