VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FirmOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public pfn  As String
Public Land As String 'z.B. "Bayern", "Baden-Württemberg", "Oberösterreich"
Public Name As String 'z.B. "BIB (Biberach)"
'Public pfn  As String 'der Pfad der Datei
'Private m_Tels() As FirmContact 'New Collection 'As List(Of FirmContact)
'Private m_Tels As New Collection 'As List(Of FirmContact)
Public Contacts As List 'Of FirmContacts 'As HashList(Of FirmContact)

Private Sub Class_Initialize()
    Set Contacts = MNew.List(vbObject) ' FirmContacts
End Sub

Friend Sub New_(ByVal aLand As String, ByVal aName As String)
    If Len(aLand) = 0 Then aLand = "Deutschland"
    Me.Land = ValidDir(aLand)
    If Len(aName) = 0 Then aName = "DefaultOffice"
    Me.Name = ValidDir(aName)
End Sub
'Public Property Get Land() As String
'    Land = m_Land
'End Property
'Public Property Get Name() As String
'    Name = m_Name
'End Property
Public Property Get Key() As String
    'key = m_Land & ", " & m_Name
    Key = Me.Name
End Property
'Public Function AddTel(aTel As FirmContact) As FirmContact
'    'Dim u As Long: u = UBound(m_Tels)
'    'Set m_Tels(u) = aTel
'    'ReDim Preserve m_Tels(0 To u + 1)
'    'Set AddTel = aTel
'    m_Tels.Add aTel
'End Function
'
'Public Property Get Item(ByVal Index As Integer) As FirmContact
'    Set Item = m_Tels(Index)
'End Property
'Public Property Get Count() As Long
'    'Count = UBound(m_Tels)
'    Count = m_Tels.Count
'End Property
'
Public Sub ToFileBinary(ByVal aFileName As String)
    Me.pfn = aFileName
Try: On Error GoTo Finally
    Dim FNr As Integer: FNr = FreeFile
    Open pfn For Binary Access Write As FNr
    Put FNr, , Me.Name & vbCrLf
    Put FNr, , "Name, Vorname" & vbTab & "Telefon" & vbTab & "Mobil" & vbTab & "Emailadresse" & vbCrLf
    Put FNr, , Me.Contacts.ToStr
Finally:
    Close FNr
    If Err Then
        Debug.Print Err.Description
    End If
End Sub
Public Sub ToFileAppend(ByVal Appendable_FNr As Integer)
    Print #Appendable_FNr, Me.ToStr
End Sub

Public Function ToStr() As String
    Dim s As String
    s = Me.Name & vbCrLf
    s = s & Me.Contacts.ToStr
    ToStr = s
End Function

