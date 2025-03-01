Attribute VB_Name = "MNew"
Option Explicit
#If False Then
    Firm , FirmCompany, FirmExtCompany, FirmOffice, FirmContact, List, PathFileName
#End If

Public Function Firm(aFirm As Firm) As Firm
    If aFirm Is Nothing Then
        Set aFirm = New Firm
        Dim MapiNamespace As NameSpace: Set MapiNamespace = Outlook.GetNamespace("MAPI")
        'Call MapiNamespace.Logon(Nothing, Nothing, True, False)
        Dim GlobAddrList  As AddressList: Set GlobAddrList = MapiNamespace.GetGlobalAddressList
        aFirm.Parse GlobAddrList 'addresses
        
        Dim folMapi As Outlook.MAPIFolder: Set folMapi = MapiNamespace.GetDefaultFolder(olFolderContacts)
        Dim strContactFilter As String: strContactFilter = "[MessageClass] = 'IPM.Contact'"
        Dim itmAll  As Outlook.Items: Set itmAll = folMapi.Items
        Dim itmReal As Outlook.Items: Set itmReal = itmAll.Restrict(strContactFilter)
        'Dim itmContacts As Outlook.ContactItem
        aFirm.ParseContacts itmReal
    End If
    Set Firm = aFirm
End Function

Public Function FirmCompany(aName As String) As FirmCompany
    Set FirmCompany = New FirmCompany: FirmCompany.New_ aName
End Function

Public Function FirmExtCompany(aName As String) As FirmExtCompany
    Set FirmExtCompany = New FirmExtCompany: FirmExtCompany.New_ aName
End Function

Public Function FirmOffice(aLand As String, aName As String) As FirmOffice
    Set FirmOffice = New FirmOffice: FirmOffice.New_ aLand, aName
End Function

Public Function FirmContact(aCompany As FirmCompany, aName As String, aTel As String, aMob As String, aEmail As String) As FirmContact
    Set FirmContact = New FirmContact: FirmContact.New_ aCompany, aName, aTel, aMob, aEmail
End Function

Public Function List(Of_T As EDataType, _
                     Optional ArrColStrTypList, _
                     Optional ByVal IsHashed As Boolean = False, _
                     Optional ByVal Capacity As Long = 32, _
                     Optional ByVal GrowRate As Single = 2, _
                     Optional ByVal GrowSize As Long = 0) As List
    Set List = New List: List.New_ Of_T, ArrColStrTypList, IsHashed, Capacity, GrowRate, GrowSize
    'https://github.com/OlimilO1402/List_GenericNLinq/blob/main/Classes/List.cls
End Function

Public Function PathFileName(ByVal aPathFileName As String, _
                    Optional ByVal aFileName As String, _
                    Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathFileName, aFileName, aExt
    'https://github.com/OlimilO1402/IO_PathFileName/blob/main/Classes/PathFileName.cls
End Function


