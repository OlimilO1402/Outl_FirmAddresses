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
Public PFN  As String
Public Land As String 'z.B. "Bayern", "Baden-Württemberg", "Oberösterreich"
Public Name As String 'z.B. M (München)
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

Public Property Get Key() As String
    'key = m_Land & ", " & m_Name
    Key = Me.Name
End Property

Public Sub ToFileBinary(ByVal aFileName As String)
    Me.PFN = aFileName
Try: On Error GoTo Finally
    Dim FNr As Integer: FNr = FreeFile
    Open PFN For Binary Access Write As FNr
    Put FNr, , Me.Name & vbCrLf
    Put FNr, , "Name, Vorname" & vbTab & "Telefon" & vbTab & "Mobil" & vbTab & "Emailadresse" & vbCrLf
    Put FNr, , MFirmHelper.Contacts_ToStr(Contacts)
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
    s = s & MFirmHelper.Contacts_ToStr(Contacts)
    ToStr = s
End Function

