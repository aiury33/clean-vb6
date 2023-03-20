VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompanyRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mID As Long
Private formatCur As New CurrencyFormat

Public Property Let SetID(id As Long)
    
    mID = id
    
End Property

Public Function FindCompany() As Company
    
    Dim rsCompany As ADODB.Recordset
    Dim comp As New Company
    
    Set rsCompany = FindCompanyInDatabase()
    
    If Not rsCompany.EOF Then
        
        FillCompany comp, rsCompany
                
    End If
    
    Set FindCompany = comp
    
End Function

Private Sub FillCompany(comp As Company, rsCompany As ADODB.Recordset)

    comp.SetID = rsCompany!IDCompany
    comp.SetName = rsCompany!name
    comp.SetWallet = rsCompany!wallet

End Sub

Private Function FindCompanyInDatabase() As ADODB.Recordset
    
    Dim finder As New DatabaseFinder
    Dim query As String
    
    query = GetSelectQuery
    FilterID query
    
    Set FindCompanyInDatabase = finder.Find(query)
    
End Function

Private Function GetSelectQuery() As String
    
    GetSelectQuery = "SELECT TOP 1 * FROM Company"
    
End Function

Private Sub FilterID(query As String)
    
    If mID = 0 Then Exit Sub
    
    query = query & vbCr & "WHERE IDCompany = " & mID
    
End Sub

Public Function GetUpdateAdjustCompanyWallet(value As Currency) As String
    
    Dim query As String
    
    query = "UPDATE Company SET Wallet = Wallet + (" & formatCur.FormatCurrencySQL(value) & ")"
    FilterID query
    
    query = query & ";"
    
    GetUpdateAdjustCompanyWallet = query
    
End Function
