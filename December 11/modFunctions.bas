Attribute VB_Name = "modFunctions"
'MENU API FOR ICONS
Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wid As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
End Type
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bypos As Long, lpcMenuItemInfo As MENUITEMINFO) As Long
Public Const MF_BITMAP = &H4&
Public Const MFT_BITMAP = MF_BITMAP
Public Const MIIM_TYPE = &H10
'WHOIS STUFF
Public sckServer As String * 15
Public sckHandle As String, sckPort As Long

'AUTO COMPLETE STUFF
Public Const CB_FINDSTRING = &H14C
Private Const CB_SHOWDROPDOWN = &H14F
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Function sckHandleInfo()
        Select Case frmWhois.lstServers.ListIndex
            Case 0
            sckServer$ = "64.124.14.21"
            sckHandle$ = "whois.alldomains.com"
            sckPort = "43"
            
            Case 1
            sckServer$ = "202.12.29.13"
            sckHandle$ = "whois.apnic.net"
            sckPort = "43"
            
            Case 2
            sckServer$ = "192.149.252.44"
            sckHandle$ = "whois.arin.net"
            sckPort = "43"
            
            Case 3
            sckServer$ = "198.41.0.6"
            sckHandle$ = "whois.internic.net"
            sckPort = "43"
            
            Case 4
            sckServer$ = "216.55.191.55"
            sckHandle$ = "whois.names4ever.net"
            sckPort = "43"
            
            Case 5
            sckServer$ = "193.0.0.135"
            sckHandle$ = "whois.ripe.net"
            sckPort = "43"
            
        End Select
End Function
Function AutoComplete(cbCombo As ComboBox, sKeyAscii As Integer, bUpperCase As Boolean) As Integer
   ' Dim lngFind As Long, intPos As Integer, intLength As Integer
   ' Dim tStr As String


    With cbCombo
        

           ' Dim abc As Long
           ' abc = SendMessage(.hwnd, &H14F, -1, ByVal .Text)

        
        If sKeyAscii = 8 Then
            If .SelStart = 0 Then Exit Function
            .SelStart = .SelStart - 1
            .SelLength = 32000
            .SelText = ""
        Else
           intPos = .SelStart '// save intial cursor position
            tStr = .Text '// save String


            If bUpperCase = True Then
               .SelText = UCase(Chr(sKeyAscii)) '// change string. (uppercase only)
            Else
                .SelText = Chr(sKeyAscii) '// change string. (leave case alone)
            End If
        End If
        
        
        lngFind = SendMessage(.hwnd, CB_FINDSTRING, 0, ByVal .Text) '// Find string in combobox

      '  If lngFind = -1 Then '// if String Not found
      '      .Text = tStr '// Set old String (used For boxes that require charachter monitoring
      '      .SelStart = intPos '// Set cursor position
      '      .SelLength = (Len(.Text) - intPos) '// Set selected length
      '      AutoComplete = 0 '// return 0 value to KeyAscii
      '      Exit Function
            
        'Else '// If String found
        If lngFind <> -1 Then
            intPos = .SelStart '// save cursor position
            intLength = Len(.List(lngFind)) - Len(.Text) '// save remaining highlighted text length
            .SelText = .SelText & Right(.List(lngFind), intLength) '// change new text in String
            .Text = .List(lngFind) '// Use this inst
            '     ead of the above .Seltext line to set th
            '     e text typed to the exact case of the it
            '     em selected in the combo box.
            .SelStart = intPos '// Set cursor position
            .SelLength = intLength '// Set selected length
        End If
    End With
    
End Function

