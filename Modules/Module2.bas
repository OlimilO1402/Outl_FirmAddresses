Attribute VB_Name = "Module2"
Option Explicit
'eine Datei speichern in der alle "FirmenKürzel", "Nachname", "Vorname", "Emailadresse"
'Infos anschließend aus Datei holen

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
    
    Dim FileNames()   As String: Call ParseSelectedMails(SelectedMails, FileNames())    'GetProperFileNames(SelectedMails)
    Dim FNm           As String:               FNm = FileNames(1)
    
    Dim tmpFn As String
    Dim c As Long
        Dim Path As String
        
        'ACHTUNG !!!! HIER DEN PFAD ANPASSEN!!!!
        Path = "\\DiesistmeinPfad\DatenPfad\MeineFirma\Kunden\MeinKunde\"
        Path = Path & "MeinProjektPfad\"
        Path = Path & "Schriftverkehr\Projekt\Jahr\Quartal\Gesendet\Monat\"
        Dim eml As Outlook.MailItem
        Dim i As Long
        For i = 1 To SelectedMails.Count
            'Debug.Print SelectedMails.Count
            Set eml = SelectedMails(i)
            FNm = FileNames(i)
            If Len(FNm) > 77 Then
                FNm = Trim(Left(FNm, 77))
            End If
            tmpFn = Path & FNm & ".msg"
            If FileExists(tmpFn) Then
                Do Until Not FileExists(tmpFn)
                    c = c + 1
                    tmpFn = Path & FNm & "(" & CStr(c) & ")" & ".msg"
                Loop
            Else
                c = 0
            End If
            FNm = tmpFn
            SaveEmail eml, FNm
        Next
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
    Dim List As New Collection
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

Public Sub SelectedMailsToFile(Col As Collection) 'Of Outlook.MailItem
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
        Dim obj As Object
        For i = 1 To sel.Count
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

Public Function RemoveChars(Chars As String, Value As String) As String
    Dim s As String: s = Value
    Dim c As String
    Dim i As Long
    For i = 1 To Len(Chars)
        c = Mid(Chars, i, 1)
        If InStr(1, s, c) Then
            s = Replace(s, c, " ")
        End If
    Next
    RemoveChars = s
End Function

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

Public Function Contains(Col As Collection, Item As String) As Boolean
    On Error Resume Next
    If IsEmpty(Col(Item)) Then: 'DoNothing
    Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

