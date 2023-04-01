VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FirmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_PFN As String
Public Name As String
Public Offices  As List 'Of FirmOffice  'As HashList(Of FirmOffice)
Public Contacts As List 'Of FirmContacts 'As HashList(Of FirmContact)

Private Sub Class_Initialize()
    Set Offices = MNew.List(vbObject) 'Of FirmOffice
    Set Contacts = MNew.List(vbObject) 'Of FirmContacts
End Sub

Friend Sub New_(aName As String)
    Me.Name = aName
End Sub

Public Property Get Abbrev() As String
    Dim nam As String: nam = Trim(LCase(Me.Name))
    Dim s As String
    If InStr(1, nam, "eqos") > 0 Then s = "E"
    If InStr(1, nam, "alp") > 0 Then
        s = "A"
        If InStr(1, nam, "green") > 0 Then s = s & "G"
        If InStr(1, nam, "energ") > 0 Then s = s & "E"
    End If
    If InStr(1, nam, "öster") > 0 Then s = s & "AT"
    If InStr(1, nam, "deuts") > 0 Then s = s & "DE"
    If InStr(1, nam, "holdi") > 0 Then s = s & "HO"
    If InStr(1, nam, "ceska") > 0 Then s = s & "CZ"
    If InStr(1, nam, "luxem") > 0 Then s = s & "LU"
    If InStr(1, nam, "pol") > 0 Then s = s & "PO"
    If InStr(1, nam, "beteil") > 0 Then s = s & "Bet"
    If Len(s) = 0 Then s = "EDE"
    Abbrev = s
End Property
    
Public Property Get Key() As String
    Key = Me.Abbrev
End Property

Friend Property Get pfn() As String
    pfn = m_PFN
End Property

Friend Sub ToFile(aFileName As String)
    m_PFN = aFileName
Try: On Error GoTo Finally
    Dim FNr As Integer: FNr = FreeFile
    Open pfn For Append As FNr 'Access Write
    Print #FNr, Me.Key & vbTab & Me.Name & vbCrLf
    Print #FNr, "Name, Vorname" & vbTab & "Tel" & vbTab & "Mobil" & vbCrLf
    Dim offi As FirmOffice
    Offices.ToFilesAppend FNr
'    For Each offi In m_OfficeList
'        offi.ToFileAppend FNr
'    Next
Finally:
    Close FNr
    If Err Then
        Debug.Print Err.Description
    End If
End Sub

