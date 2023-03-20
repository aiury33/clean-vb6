VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StorageMaintainence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mProduct As Product
Private mOperation As StorageMaintainenceOperations
Private Const FIXED_MOVEMENT_QUANTITY = 1

Public Enum StorageMaintainenceOperations
    Buying = 0
    Selling = 1
End Enum

Public Property Let SetProduct(prod As Product)
    
    Set mProduct = prod
    
End Property

Public Property Let SetOperation(operation As StorageMaintainenceOperations)
    
    mOperation = operation
    
End Property

Public Sub Alter()
    
    Dim inserter As New DatabaseInserter
    
    bulkQuery = GetBulkQueryToAlterStorage
    
    inserter.Insert bulkQuery
    
End Sub

Private Function GetBulkQueryToAlterStorage() As String

    GetBulkQueryToAlterStorage = GetQueryInsertMovement(CreateMovement) & vbCr & _
                                 GetQueryAdjustProductQuantity & vbCr & _
                                 GetQueryAdjustCompanyWallet & vbCr

End Function

Private Property Get QuantityToAdjust() As Double

    QuantityToAdjust = IIf(mOperation = Buying, FIXED_MOVEMENT_QUANTITY, -FIXED_MOVEMENT_QUANTITY)
    
End Property

Private Property Get BaseValue() As Currency

    BaseValue = IIf(mOperation = Buying, -mProduct.GetCost, mProduct.GetPrice)
    
End Property

Private Function CreateMovement() As Movement
    
    Dim newMovement As New Movement
    
    newMovement.SetProductID = mProduct.GetID
    newMovement.SetCost = mProduct.GetCost
    newMovement.SetPrice = mProduct.GetPrice
    newMovement.SetQuantity = QuantityToAdjust
    
    Set CreateMovement = newMovement
    
End Function

Private Function GetQueryInsertMovement(newMovement As Movement) As String
    
    Dim repository As New MovementRepository
    
    GetQueryInsertMovement = repository.GetInsertQuery(newMovement)
    
End Function

Private Function GetQueryAdjustProductQuantity() As String
    
    Dim repository As New ProductRepository
    
    repository.SetID = mProduct.GetID
    GetQueryAdjustProductQuantity = repository.GetUpdateAdjustQuantityQuery(QuantityToAdjust)
    
End Function

Private Function GetQueryAdjustCompanyWallet() As String
    
    Dim repository As New CompanyRepository
    
    repository.SetID = mProduct.GetCompanyID
    GetQueryAdjustCompanyWallet = repository.GetUpdateAdjustCompanyWallet(FIXED_MOVEMENT_QUANTITY * BaseValue)
    
End Function
