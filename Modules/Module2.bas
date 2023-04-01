Attribute VB_Name = "Module2"
Option Explicit
'eine Datei speichern in der alle "FirmenKürzel", "Nachname", "Vorname", "Emailadresse"
'Infos anschließend aus Datei holen

'Sub TestLike()
'                                                            'key = "EAT"
'    Dim nam As String: nam = "EQOS Energie Österreich GmbH"
'    Dim key As String: key = Trim(LCase(nam))
'
'    If key Like "?Österreich?" Then
'        MsgBox "OK key is like " & key
'    End If
'
'End Sub
Sub AusgewaehlteMailsInOrdnerKopieren()
Try: On Error GoTo Catch
    If Firm Is Nothing Then
        Set Firm = MNew.Firm(Firm)
    End If
    
    'alle ausgewählten Mails sammeln
    Dim SelectedMails As Collection: Set SelectedMails = GetSelectedMails 'Of Outlook.MailItem
    Debug.Print SelectedMails.Count
    If SelectedMails.Count = 0 Then
        MsgBox "Keine Emails gewählt. Bitte zuerst Emails auswählen!"
        Exit Sub
    End If
    'Else
    'Dim s As String: s =
    'dim
    'ParseSelectedMails (SelectedMails)
    'Debug.Print s
    
    'Exit Sub
        'SelectedMailsToFile SelectedMails
    'End If
    'Dim sl As String: sl = ParseSelectedMails(SelectedMails)
    'Debug.Print sl
    'Exit Sub
    'Dim FileNames     As Collection: Set FileNames = ParseSelectedMails(SelectedMails) 'GetProperFileNames(SelectedMails)
    Dim FileNames()   As String: Call ParseSelectedMails(SelectedMails, FileNames())    'GetProperFileNames(SelectedMails)
    Dim FNm           As String:               FNm = FileNames(1)
    
    'Debug.Print Fnm
    
    'OK dann eben keinen SaveFileDialog
    'Mist wie heißt der SaveFileDialog unter .Net?
    'Dim SFD As New SaveFileDialog
    'application.
    'SFD.Filter = "*.msg"
    'SFD.FileName = FNm & ".msg"
    'If SFD.ShowDialog() = vbOK Then
        'Debug.Print SFD.FileName
    Dim tmpFn As String
    Dim c As Long
        Dim path As String
        
        'ACHTUNG !!!! HIER IMMER DEN PFAD ANPASSEN!!!!
        path = "\\BIBFS05\Daten\daten\Engineering\FTZ-E\KUNDEN\TenneT\"
        path = path & "380kV-Ltg. Süderdonn-Heide West, LH-13-0319\"
        'path = "\\bibfs05\Daten\daten\Engineering\FTZ-E\KUNDEN\NetzeBW\110kV-Ltg. A0408 Probabilistische Überrechnung\Schriftverkehr\"
        
        'path = path & "380kV-Ltg. Audorf-Flensburg\Schriftverkehr\Pfahlgründungen\Posteingang\"
        'path = path & "380kV-Ltg. Audorf-Flensburg\Schriftverkehr\Pfahlgründungen\Gesendet\"
        'path = path & "Schriftverkehr\Pfahlgründungen\2018\Q3\Posteingang\08\"
        path = path & "Schriftverkehr\Pfahlgründungen\2018\Q3\Gesendet\07\"
        Dim eml As Outlook.MailItem
        Dim i As Long
        For i = 1 To SelectedMails.Count
            'Debug.Print SelectedMails.Count
            Set eml = SelectedMails(i)
            FNm = FileNames(i)
            If Len(FNm) > 77 Then
                FNm = Trim(Left(FNm, 77))
                'Debug.Print FNm
            End If
            'FNm = path & FNm & ".msg"
            tmpFn = path & FNm & ".msg"
            If FileExists(tmpFn) Then
                Do Until Not FileExists(tmpFn)
                    c = c + 1
                    tmpFn = path & FNm & "(" & CStr(c) & ")" & ".msg"
                Loop
            Else
                c = 0
            End If
            FNm = tmpFn
            'FNm = path & FNm
            'Debug.Print FNm
            SaveEmail eml, FNm
        Next
    'Else
    '    Debug.Print "kein Dialog oder Abbrechen?"
    'End If
    'Dim Fld As Outlook.Folder
    'fld.
    'CurrMail.SaveAs FNm

    'einen Speichern unter Dialog anzeigen. den Dateiname erzeugen
    Exit Sub
Catch:
    Debug.Print "err " & Err.Number & Err.Description
End Sub

Private Sub SaveEmail(eml As Outlook.MailItem, FNm As String)
Try: On Error GoTo Catch
    eml.SaveAs FNm
    Exit Sub
Catch:
    Debug.Print "Err Fehler"
    Debug.Print "Err " & Err.Number & " " & Err.Description
    Debug.Print "Err " & FNm
End Sub

Public Sub ParseSelectedMails(SelectedMails As Collection, ByRef list_fn_out() As String) 'As Collection 'Of Outlook.MailItem
'Public Function ParseSelectedMails(SelectedMails As Collection) As String 'Of Outlook.MailItem
    Dim s As String
    Dim sl As String
    Dim list As New Collection
    Dim Email As Outlook.MailItem
    Dim ec As FirmContact
    Dim sea As String
    Dim recps As Recipients
    Dim recp  As Recipient
    Dim PR_SMTP_ADDRESS As String: PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
    Dim wrongchars As String: wrongchars = "\/:*?""<>|"
    Dim i As Long
    If SelectedMails.Count > 0 Then
        ReDim list_fn_out(1 To SelectedMails.Count)
        Debug.Print SelectedMails.Count
        For i = 1 To SelectedMails.Count
            
            Set Email = SelectedMails.Item(i)
            If Not Email Is Nothing Then
                'entweder Empfänger oder Sender?
                Set recps = Email.Recipients
                If Not recps Is Nothing Then
                    If recps.Count > 0 Then
                        Set recp = recps.Item(1)
                        Dim pa As PropertyAccessor: Set pa = recp.PropertyAccessor
                        sea = pa.GetProperty(PR_SMTP_ADDRESS) ' recp.address
                        'Debug.Print sea
                        If Firm.Contacts.ContainsKey(sea) Then
                            Set ec = Firm.Contacts.Item(sea)
                        End If
                        If Not ec Is Nothing Then
                            Dim dat    As Date:  dat = Email.ReceivedTime
                            Dim subj As String: subj = Email.Subject
                            subj = IIf(Len(subj), subj, "(Kein Betreff)")
                            subj = RemoveChars(wrongchars, subj)
                            s = Year(dat) & "-" & Month(dat) & "-" & Day(dat) & "-a" & ec.AbteilungAbbrev & " " & subj '& vbCrLf
                            'sl = sl & s & vbCrLf
                            'list.Add s
                            list_fn_out(i) = s
                        End If
                    End If
                End If
            End If
        Next
    End If
    'ParseSelectedMails = sl
    'Set ParseSelectedMails = list
End Sub

Public Sub SelectedMailsToFile(col As Collection) 'Of Outlook.MailItem
    Dim Email As Outlook.MailItem
    Dim FNr As Integer: FNr = FreeFile
    Dim FNm As String: FNm = "C:\"
    For Each Email In SelectedMails
        GetName (Email)
    Next
End Sub

Public Function GetSelectedMails() As Collection 'Of Outlook.MailItem
    'welches ist die aktuelle Mail in Outlook?
    If TypeOf Application.ActiveWindow Is Outlook.Explorer Then
        Dim Explorer As Outlook.Explorer:  Set Explorer = Application.ActiveExplorer ' Application.ActiveWindow
        Dim m As Outlook.MailItem
        Dim i As Long
        Set GetSelectedMails = New Collection
        Dim sel As Selection: Set sel = Explorer.Selection
        On Error Resume Next
        'Debug.Print Explorer.Selection.Count
        'Dim i As Long
        'OK jetzt nicht meher über For Each, das war ja Mist,
        'vielleicht funzt es jetzt besser?
        Dim obj As Object
        For i = 1 To sel.Count
        'For i = 1 To Explorer.Selection.Count ' m In Explorer.Selection.
            Set obj = sel.Item(i)
            If TypeOf obj Is Outlook.MailItem Then
                Set m = obj
            End If
            If Not m Is Nothing Then
                GetSelectedMails.Add m
            End If
        Next
    Else
        Dim Inspectr As Outlook.Inspector: Set Inspectr = Application.ActiveInspector
        Inspectr.Activate
        If Inspectr.IsWordMail Then
            Set GetSelectedMails = New Collection
            GetSelectedMails.Add Inspectr.CurrentItem
            'Set GetCurrentMail = Inspectr.CurrentItem
        End If
    End If
End Function

Public Function GetProperFileName(Email As Outlook.MailItem, Optional withTime As Boolean = False) As String
    Dim FirstRecipient As Recipient:   Set FirstRecipient = Email.Recipients.Item(1)
    Dim rec_addrentry  As AddressEntry: Set rec_addrentry = FirstRecipient.AddressEntry
    Dim rec_contact    As ContactItem
    Dim an_Name As String
    Dim an_Eadr As String
    Dim an_Firm As String
    an_Name = FirstRecipient.Name
    an_Eadr = FirstRecipient.address

Try: On Error Resume Next
    
    If rec_addrentry Is Nothing Then
        'Outlook.
    End If
    
    If Not rec_addrentry Is Nothing Then
        'in der Email kein companyname, email anhand Emailadresse erst mit Contaktliste abchecken
        
        Set rec_contact = rec_addrentry.GetContact
        If Not rec_contact Is Nothing Then
            an_Firm = rec_contact.CompanyName
        End If
    End If
    'Debug.Print an_Eadr
    'der gleiche Name wie CurrMail.Recipients.Item(1).Name
    Dim vonName As String: vonName = Email.SenderName
    Dim vonEadr As String: vonEadr = Email.SenderEmailAddress
    Dim vonFirm As String: vonFirm = Email.sender.GetContact.CompanyName
    'Dim dat As Date: dat = CurrMail.CreationTime 'oder ReceivedTime?
    Dim dat     As Date:       dat = Email.ReceivedTime
    'GetProperFileName = Year(dat) & "-" & Month(dat) & "-" & Day(dat) & " an " & an_Name & " von " & vonName & " btr " & CurrMail.Subject
    Dim s As String
    s = Year(dat) & "-" & Month(dat) & "-" & Day(dat)
    If withTime Then
        s = s & "-" & Hour(dat) & "-" & Minute(dat) & "-" & Second(dat)
    End If
    s = s & "_a" & GetName(an_Name, an_Firm, an_Eadr)
    s = s & "_v" & GetName(vonName, vonFirm, vonEadr) ', eMail.sender.GetContact.CompanyName)
    'OK in Subject könnten noch Zeichen drin sein, die als Dateiname ungültig sind
    'also vorher rausparsen von "\", "/", ":", "*", "?", """", "<", ">", "|"
    '"\/:*?""<>|"
    Dim wrongchars As String: wrongchars = "\/:*?""<>|"
    Dim subj As String: subj = Email.Subject: subj = IIf(Len(subj), subj, "(Kein Betreff)")
    s = s & " - " & RemoveChars(wrongchars, subj)
    GetProperFileName = s
    'Debug.Print s
End Function

Public Function RemoveChars(chars As String, Value As String) As String
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
'Public Function GetName(aName As String, Optional FirmName As String, Optional Emailaddress As String) As String
'    If Len(aName) = 0 Then Exit Function
'    Dim sep As String: sep = IIf(InStr(1, aName, ","), ",", IIf(InStr(1, aName, " "), " ", " "))
'    Dim sa() As String: sa = Split(aName, sep)
'    Dim s    As String:  s = Replace(Trim(sa(0)), "'", "")
'    Select Case LCase(s)
'    Case "meyer", "haller", "winderlich", "oldenburg", "glöggler", "karahasanovic", "beitz", "emmenlauer", "bornemann", "vukobrat", "krepp", "besier", "wohlhueter", "isy-support", "materne", "gaertner"
'                        s = "EDE"
'    Case "mauracher", "rapp", "mauracher bernhard", "isola"
'                        s = "EAT"
'    Case "linz", "olaf.linz@tennet.eu", "rohrmoser", "buerger", "hansen", "zabold", "oesterlink", "boettger", "weike"
'                        s = "TEN"
'    Case "eggers":      s = "OMX"
'    Case "frenzel", "ventker"
'                        s = "SPZ"
'    Case "janssen", "janßen"
'                        s = "IBJ"
'    Case "janus", "richwien"
'                        s = "RiP"
'    Case "vierkant", "scholz kerstin"
'                        s = "BuP"
'    Case "balázsovics", "schüttné balázsovics mónika", "schüttné balázsovics", "schüttné", "rauch lubos", "katona zoltán"
'                        s = "OVI"
'    Case "jurkova katerina"
'                        s = "EGE"
'    Case Else
'        If Len(FirmName) Then
'            Select Case FirmName
'            Case "TenneT"
'                        s = "TEN"
'            Case "EQOS Deutschland GmbH"
'                        s = "EDE"
'            Case "EQOS Austria GmbH"
'                        s = "EAT"
'            Case "Spitzke SE"
'                        s = "SPZ"
'            Case Else
'
'            End Select
'        Else
'            's bleibt so
'        End If
'    End Select
'
'    GetName = s
'End Function

'Public Function GetShortCompany(ByVal Emailaddress As String) As String
'    Emailaddress = Trim(LCase(Emailaddress))
'    Dim scn As String
'    Select Case True
'    Case CBool(InStr(1, Emailaddress, "@tennet.eu") > 1)
'        scn = "TEN"
'    Case CBool(InStr(1, Emailaddress, "@eqos-energie.com"))
'        scn = "EDE"
'    Case CBool(InStr(1, Emailaddress, "@eqos-energie.com"))
'
'    End Select
'End Function
Public Function GetProperFileNames(Mails As Collection) As Collection 'Of string 'Collection Of Outlook.MailItem
    Set GetProperFileNames = New Collection
    Dim Email As Outlook.MailItem
    'For i = 1 To SelectedMails.Count
    For Each Email In Mails
        'Set eMail = SelectedMails(i)
        Dim FNm As String: FNm = GetProperFileName(Email)
        'so jetzt noch überprüfen ob evtl gleiche Dateinamen drin,
        'wenn ja dann nochmal Dateinamen, diesmal mit Zeitangabe
        If Contains(GetProperFileNames, FNm) Then
            FNm = GetProperFileName(Email, True)
            If Contains(GetProperFileNames, FNm) Then
                FNm = FNm & "(1)"
                'so und jetzt für sämtliche anderen auch!
                'd.h. erst mal sämtliche gleichen Dateinamen herausfinden.
                
            End If
        End If
        GetProperFileNames.Add FNm
    Next
End Function
Public Function Contains(col As Collection, Item As String) As Boolean
    On Error Resume Next
    If IsEmpty(col(Item)) Then: 'DoNothing
    Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

