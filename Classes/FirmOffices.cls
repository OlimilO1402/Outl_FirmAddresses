VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FirmOffices_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_List As Collection '(Of FirmContact)

Private Sub Class_Initialize()
    Set m_List = New Collection
End Sub

Public Function Add(aNewOffice As FirmOffice) As FirmOffice
    Set Add = aNewOffice
    Dim Key As String: Key = Add.Key
    If ContainsKey(Key) Then
        Set Add = m_List.Item(Key)
    Else
        m_List.Add Add, Key
    End If
End Function

Private Function ContainsKey(Key As String) As Boolean
Try: On Error Resume Next
    If IsEmpty(m_List(Key)) Then:  'DoNothing
    ContainsKey = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Property Get Item(Key As String) As FirmOffice
    Set Item = m_List.Item(Key)
End Property

Public Property Get Contains(aOffice As FirmOffice) As Boolean
    Contains = ContainsKey(aOffice.Key)
End Property

Public Property Get Count() As Long
    Count = m_List.Count
End Property

Public Property Get List() As Collection
    Set List = m_List
End Property

Public Function ToFilesBinary(Prepfn As String)
    Dim pfn As String
    Dim offi As FirmOffice
    For Each offi In m_List
        pfn = Prepfn & offi.Name & ".txt"
        offi.ToFileBinary pfn
    Next
End Function

Public Function ToFilesAppend(Appendable_FNr As Integer)
    Dim pfn As String
    Dim offi As FirmOffice
    For Each offi In m_List
        'pfn = prepfn & offi.Name & ".txt"
        offi.ToFileAppend Appendable_FNr
    Next
End Function

Public Function ToPFNs() As String
    Dim s As String
    Dim offi As FirmOffice
    For Each offi In m_List
        s = s & """" & offi.pfn & """" & " " 'vbCrLf
    Next
    ToPFNs = s
End Function
