Attribute VB_Name = "MFirmHelper"
Option Explicit

Public Function Offices_ToFilesAppend(Offices As List, ByVal Appendable_FNr As Integer) 'Offices As List Of FirmOffice
    Dim PFN As String
    Dim offi As FirmOffice
    Dim i As Long
    For i = 0 To Offices.Count - 1
        Set offi = Offices.Item(i)
        'pfn = prepfn & offi.Name & ".txt"
        offi.ToFileAppend Appendable_FNr
    Next
End Function

Public Function Offices_ToFilesBinary(Offices As List, Prepfn As String) 'Offices As List Of Office
    Dim PFN As String
    Dim offi As FirmOffice
    Dim i As Long
    For i = 0 To Offices.Count - 1
        Set offi = Offices.Item(i)
        PFN = Prepfn & offi.Name & ".txt"
        offi.ToFileBinary PFN
    Next
End Function

Public Function Offices_ToPFNs(Offices As List) 'Offices As List Of Office
    Dim s As String
    Dim offi As FirmOffice
    Dim i As Long
    For i = 0 To Offices.Count - 1
        Set offi = Offices.Item(i)
        s = s & """" & offi.PFN & """" & " " 'vbCrLf
    Next
    Offices_ToPFNs = s
End Function

Public Function Contacts_ToStr(Contacts As List) As String 'Contacts as List Of FirmContact
    Dim s As String
    Dim ctc As FirmContact
    Dim i As Long
    For i = 0 To Contacts.Count - 1
        Set ctc = Contacts.Item(i)
        s = s & ctc.ToStr & vbCrLf
    Next
    Contacts_ToStr = s
End Function

Public Function Companys_ToFiles(Companys As List, PFN As String) 'Companys As List Of Company
    Dim PFN As String
    Dim cmp As FirmCompany
    Dim i As Long
    For i = 0 To Companys.Count - 1
        Set cmp = Companys.Item(i)
        PFN = Prepfn & cmp.Key & ".txt"
        'Debug.Print pfn
        cmp.ToFile PFN
    Next
End Function

