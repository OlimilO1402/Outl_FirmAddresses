VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FirmExtCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Klasse repräsentiert eine allgemeine Firma, die ebenso "Teil" aber ein externer Teil der FIRMA ist, und allgemein zu Kunden oder Subunternehmen zählt
'Klasse repräsentiert die Firma eines Kunden oder Subunternehmens der FIRMA. Offices werden hier nicht gebraucht
Implements FirmCompany
Public pfn      As String
Public Name     As String
Public m_Abbrev As String
Public Contacts As list 'Of FirmContact

Friend Sub New_(aName As String) ', aAbbrev As String)
    Me.Name = aName
    'm_Abbrev = aAbbrev
    'If Len(m_Abbrev) = 0 Then
       m_Abbrev = GetAbbrev
    'End If
End Sub

Private Function GetAbbrev() As String
    Dim s As String
    If Len(m_Abbrev) = 0 Then
        If NameContains("tennet") Then
            s = "TEN"
        ElseIf NameContains("terracon") Then
            s = "TRC"
        ElseIf NameContains("strukton") Then
            s = "STR"
        ElseIf NameContains("lincon") Then
            s = "LIN"
        ElseIf NameContains("spitzke") Then
            s = "SPZ"
        ElseIf NameContains("ege") Then
            s = "EGE"
        ElseIf NameContains("ovit") Then
            s = "OVI"
        ElseIf NameContains("transnetbw") Then
            s = "TRA"
        ElseIf NameContains("netzebw") Then
            s = "NET"
        ElseIf NameContains("omexom") Then
            s = "OMX"
        ElseIf NameContains("janssen") Or NameContains("janßen") Then
            s = "IBJ"
        ElseIf NameContains("rizkallah") Then
            s = "RiP"
        ElseIf NameContains("buchholz") Then
            s = "BuP"
        ElseIf NameContains("dc-software") Then
            s = "DCS"
        ElseIf NameContains("dmt ingenieure") Then
            s = "DMT"
        ElseIf NameContains("sag") Then
            s = "SAG"
        ElseIf NameContains("spie") Then
            s = "SAG"
        ElseIf NameContains("imp") Then
            s = "IMP"
        ElseIf NameContains("franki") Then
            s = "FRA"
        End If
    End If
    GetAbbrev = s
End Function
Public Property Get Abbrev() As String
    Abbrev = m_Abbrev
End Property
Private Function NameContains(s As String) As Boolean
    NameContains = InStr(1, Me.Name, s, vbTextCompare)
End Function

Public Property Get Key() As String
    Key = m_Abbrev
End Property

Private Property Get FirmCompany_Abbrev() As String
    FirmCompany_Abbrev = Me.Abbrev
End Property

Private Property Set FirmCompany_Contacts(ByVal Value As list) 'FirmContacts)
    Set Me.Contacts = Value
End Property
Private Property Get FirmCompany_Contacts() As FirmContacts
    Set FirmCompany_Contacts = Me.Contacts
End Property

Private Property Get FirmCompany_Key() As String
    FirmCompany_Key = Me.Key
End Property

Private Property Let FirmCompany_Name(ByVal RHS As String)
    Me.Name = RHS
End Property
Private Property Get FirmCompany_Name() As String
    FirmCompany_Name = Me.Name
End Property

Private Property Set FirmCompany_Offices(ByVal Value As list) 'RHS As FirmOffices)
    '
End Property
Private Property Get FirmCompany_Offices() As FirmOffices
    Set FirmCompany_Offices = Nothing
End Property
