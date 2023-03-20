VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProductRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mID As Long
Private formatFloat As New FloatFormat

Public Property Let SetID(id As Long)
    
    mID = id
    
End Property

Public Function FindProduct() As Product
    
    Dim rsProduct As ADODB.Recordset
    Dim item As New Product
    
    Set rsProduct = FindProductInDatabase()
    
    If Not rsProduct.EOF Then
        
        FillProduct item, rsProduct
                
    End If
    
    Set FindProduct = item
    
End Function

Private Sub FillProduct(item As Product, rsProduct As ADODB.Recordset)

    item.SetID = rsProduct!IDProduct
    item.SetCompanyID = rsProduct!IDCompany
    item.SetName = rsProduct!name
    item.SetCost = rsProduct!cost
    item.SetPrice = rsProduct!price
    item.SetQuantity = rsProduct!quantity

End Sub

Private Function FindProductInDatabase() As ADODB.Recordset
    
    Dim finder As New DatabaseFinder
    Dim query As String
    
    query = GetSelectQuery
    FilterID query
    
    Set FindProductInDatabase = finder.Find(query)
    
End Function

Private Function GetSelectQuery() As String
    
    GetSelectQuery = "SELECT TOP 1 * FROM Product"
    
End Function

Private Sub FilterID(query As String)
    
    If mID = 0 Then Exit Sub
    
    query = query & vbCr & "WHERE IDProduct = " & mID
    
End Sub

Public Function GetUpdateAdjustQuantityQuery(quantity As Double) As String
    
    Dim query As String
    
    query = "UPDATE Product SET Quantity = Quantity + (" & formatFloat.FormatFloatSQL(quantity) & ")"
    FilterID query
    
    query = query & ";"
    
    GetUpdateAdjustQuantityQuery = query
    
End Function
