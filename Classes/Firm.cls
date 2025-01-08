VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Firm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'halt stopp hier sollte zuerst die Unterscheidung in LÃ¤nder
'
'Private m_CompanyList As New Collection 'As HashList(Of FirmCompany)
'Private m_OfficeList  As New Collection 'As HashList(Of FirmOffice)
'Private m_ContactList As New Collection 'As HashList(Of FirmContact)
Public XCompanys As List 'Of FirmExtCompany
Public Companys  As List 'Of FirmCompany
Public Offices   As List 'Of FirmOffice
Public Contacts  As List 'Of FirmContact
Public Path      As String
Public Prepfn    As String

Private Sub Class_Initialize()
    Set Companys = MNew.List(vbObject)  'Of FirmCompany
    Set XCompanys = MNew.List(vbObject) 'Of FirmExtCompany
    Set Offices = MNew.List(vbObject)   'Of FirmOffice
    Set Contacts = MNew.List(vbObject)  'Of FirmContact
End Sub

Public Sub ParseContacts(OutlContacts As Outlook.Items)
    Dim ContactItem As Outlook.ContactItem
    Dim ec As FirmContact
    Dim FNam As String
    Dim LNam As String
    Dim xCmp As FirmExtCompany
    For Each ContactItem In OutlContacts
        With ContactItem
            FNam = .FirstName
            LNam = .LastName
            Set xCmp = MNew.FirmExtCompany(.CompanyName)
            Set ec = MNew.FirmContact(xCmp, LNam & ", " & FNam, .BusinessTelephoneNumber, .MobileTelephoneNumber, .Email1Address)
            'Debug.Print LNam & ", " & FNam
            'If Not xCmp.Contacts Is Nothing Then
            '    xCmp.Contacts.Add ec
            'End If
        End With
        Contacts.Add ec
    Next
End Sub

Public Sub Parse(addresses As Outlook.AddressList)
    Dim address As Outlook.AddressEntry
    Dim eCmp    As FirmCompany
    Dim eOff    As FirmOffice
    Dim ectc    As FirmContact
    Dim cmpn As String
    Dim sopr As String
    Dim oloc As String
    Dim FNam As String
    Dim LNam As String
    Dim BTel As String
    Dim TMob As String
    Dim Emal As String
    Dim aeut As OlAddressEntryUserType
    Dim eu  As ExchangeUser
    Dim edl As ExchangeDistributionList
    For Each address In addresses.AddressEntries
        Set eu = address.GetExchangeUser
        If Not eu Is Nothing Then
            With eu
                cmpn = .CompanyName
                sopr = .StateOrProvince
                oloc = .OfficeLocation
                FNam = .FirstName
                LNam = .LastName
                BTel = .BusinessTelephoneNumber
                TMob = .MobileTelephoneNumber
                Emal = .PrimarySmtpAddress
            End With
            aeut = address.AddressEntryUserType
            'Debug.Print AddressEntryUserType_ToStr(aeut)
            Select Case aeut
            Case OlAddressEntryUserType.olExchangeUserAddressEntry
                If Not eu Is Nothing Then
                    'With eu
                        'Emal = .PrimarySmtpAddress
                    'End With
                End If
            Case OlAddressEntryUserType.olExchangeDistributionListAddressEntry
                 Set edl = address.GetExchangeDistributionList
                 If Not edl Is Nothing Then
                    With edl
                        Emal = .PrimarySmtpAddress
                    End With
                 End If
            End Select
            Set eCmp = Me.Companys.Add(MNew.FirmCompany(cmpn))
            Set eOff = Me.Offices.Add(MNew.FirmOffice(sopr, oloc))
            Set ectc = Me.Contacts.Add(MNew.FirmContact(eCmp, LNam & ", " & FNam, BTel, TMob, Emal))
            eCmp.Offices.Add eOff
            eCmp.Contacts.Add ectc
            eOff.Contacts.Add ectc
            'c = c + 1
        End If
    Next
End Sub

Public Sub ToFiles(Env_Userprofile As String)
    Dim n   As Date:     n = Now
    Dim d   As String:   d = Year(n) & "_" & Str0(Month(n), 2) & "_" & Str0(Day(n), 2)
    If Len(Env_Userprofile) = 0 Then Env_Userprofile = Environ("USERPROFILE")
    Path = Env_Userprofile & "\documents\TelefonListen\" & d & "\"
    If Not DirExists(Path) Then MkDir Path
    Prepfn = "TelefonListe_" & d & "_"
    Dim PFN As String: PFN = Path & Prepfn
    MFirmHelper.Offices_ToFilesBinary Offices, PFN
    MFirmHelper.Companys_ToFiles Companys, PFN
End Sub

Public Sub ToExcel()
    'entweder so
    If Not DirExists(Path) Then MsgBox "Der Pfad existiert nicht:" & vbCrLf & Path
    Dim pfns As String: pfns = MFirmHelper.Offices_ToPFNs(Offices)
    Shell "Excel.exe " & pfns, vbNormalFocus
    'oder so:
    '* Excel Instanz öffnen
    Dim Exl As New Excel.Application
    Dim wkb As Workbook: Set wkb = Exl.Workbooks.Open(Path & "\TelefonListe.xlsx")
    
    '* eine neue Datei / Workbook anlegen
    '* für jeden Text ein neues Worksheet anlegen und Text reinkopieren
    
End Sub
