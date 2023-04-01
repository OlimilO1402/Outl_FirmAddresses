Attribute VB_Name = "Module1"
Option Explicit
'Firm also has external companies
Public Firm As Firm

Sub Test()
    Set Firm = MNew.Firm(Firm)
    Dim ectc As FirmContact: Set ectc = MNew.FirmContact(Nothing, "", "", "", "name@firm.com")
    If Firm.Contacts.Contains(ectc) Then
        Set ectc = Firm.Contacts.Item(ectc.Key)
        Dim cmpy As FirmCompany: Set cmpy = ectc.Company
        MsgBox "OK " & ectc.Name & " is Member of " & ectc.Company.Abbrev
    End If
End Sub

Sub Telefonliste()
    'Makro Telefonliste
    'speichert Firm-Telefonlisten in einer Datei je Firm-Büro
    'zuerst alle Addressen durchlaufen
    'und die einzelnen Tels dem jeweiligen Office zuordnen
    'Die Addressen-Liste muss nur einmal durchlaufen werden
    'bei jedem Schritt wird geprüft
    '  ob es den jeweilige Standort schon gibt
    '  ja: die Telefonnummer zu dem Standort hinzufügen
    '  ne: den Standort raussuchen über den key und dann die Tel hinzufügen
    Set Firm = MNew.Firm(Firm)
    
    'jetzt in Dateien schreiben
    Firm.ToFiles Environ("USERPROFILE")
    
    'in Excel Öffnen?
    Dim ret As VbMsgBoxResult
    ret = MsgBox("Successfully stored a total of " & Firm.Contacts.Count & " entries in " & Firm.Offices.Count & " in telephone-lists. " & vbCrLf & _
                 "Do you like to open the files in Excel?", vbQuestion Or vbOKCancel)
    If ret = vbOK Then
        Firm.ToExcel
    End If
End Sub

Public Function Str0(s As String, ByVal length As Byte) As String
    Dim d As Long: d = length - Len(s)
    If d > 0 Then Str0 = String(d, "0") & s Else Str0 = s
End Function

Public Function ValidDir(p As String) As String
    ValidDir = ReplaceAll(p, "\/:*?""<>|", "_")
End Function

Public Function ReplaceAll(s As String, find As String, rep As String) As String
    'ersetzt alle Vorkommen von "find" in "s" durch "rep"
    If Len(s) = 0 Then Exit Function
    If Len(find) = 0 Then Exit Function
    ReplaceAll = s
    Dim i As Long
    Dim c As String
    For i = 1 To Len(find)
        c = Mid(find, i, 1)
        ReplaceAll = Replace(ReplaceAll, c, rep)
    Next
End Function

Public Function RemoveChars(chars As String, Value As String) As String
    'löscht alle Vorkommen von chars in Value
    Dim s As String: s = Value
    Dim c As String
    Dim i As Long
    For i = 1 To Len(chars)
        c = Mid(chars, i, 1)
        If InStr(1, s, c) Then
            s = Replace(s, c, " ")
        End If
    Next
    RemoveChars = s
End Function

Public Function FileExists(ByVal FileName As String) As Boolean
    On Error Resume Next
    FileExists = Not CBool(GetAttr(FileName) And (vbDirectory Or vbVolume))
    On Error GoTo 0
End Function

Public Function DirExists(ByVal DirectoryName As String) As Boolean
    On Error Resume Next
    DirExists = CBool(GetAttr(DirectoryName) And vbDirectory)
    On Error GoTo 0
End Function

Public Function ParseEmailAddress(ByVal s As String) As String
    'wenn eine Emailaddresse im String s enthalten ist dann wird das "@"-Zeichen darin gefunden
    'dann wird rund um das @-Zeichen die Emailaddresse herausgelesen
    s = ReplaceAll(s, "(){}[]<>", " ")
    Dim pos_at As Long: pos_at = InStr(1, s, "@")
    If pos_at = 0 Then Exit Function
    Dim pos_beg As Long: pos_beg = InStrRev(s, " ", pos_at) + 1
    If pos_beg = 0 Then pos_beg = 1
    Dim pos_end As Long: pos_end = InStr(pos_at, s, " ")
    'If pos_end > 0 Then pos_end = pos_end '- 1
    If pos_end = 0 Then pos_end = Len(s) + 1
    ParseEmailAddress = Mid(s, pos_beg, pos_end - pos_beg)
End Function

'Outlook-related
Public Function AddressEntryUserType_ToStr(aeut As OlAddressEntryUserType) As String
    Dim s As String: s = CStr(aeut) & " "
    If aeut And OlAddressEntryUserType.olExchangeAgentAddressEntry Then _
        s = s & "ExchangeAgentAddressEntry "
    If aeut And OlAddressEntryUserType.olExchangeDistributionListAddressEntry Then _
        s = s & "ExchangeDistributionListAddressEntry "
    If aeut And OlAddressEntryUserType.olExchangeOrganizationAddressEntry Then _
        s = s & "ExchangeOrganizationAddressEntry "
    If aeut And OlAddressEntryUserType.olExchangePublicFolderAddressEntry Then _
        s = s & "ExchangePublicFolderAddressEntry "
    If aeut And OlAddressEntryUserType.olExchangeRemoteUserAddressEntry Then _
        s = s & "ExchangeRemoteUserAddressEntry "
    If aeut And OlAddressEntryUserType.olExchangeUserAddressEntry Then _
        s = s & "ExchangeUserAddressEntry "
    If aeut And OlAddressEntryUserType.olLdapAddressEntry Then _
        s = s & "LdapAddressEntry "
    If aeut And OlAddressEntryUserType.olOutlookContactAddressEntry Then _
        s = s & "OutlookContactAddressEntry "
    If aeut And OlAddressEntryUserType.olOutlookDistributionListAddressEntry Then _
        s = s & "OutlookDistributionListAddressEntry "
    If aeut And OlAddressEntryUserType.olSmtpAddressEntry Then _
        s = s & "SmtpAddressEntry "
    AddressEntryUserType_ToStr = s
End Function
    


