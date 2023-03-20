VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Movement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mProductId As Long
Private mCost As Currency
Private mPrice As Currency
Private mQuantity As Double

Public Property Get GetProductID() As Long
    
    GetProductID = mProductId
    
End Property

Public Property Let SetProductID(id As Long)
    
    mProductId = id
    
End Property

Public Property Get GetCost() As Currency
    
    GetCost = mCost
    
End Property

Public Property Let SetCost(cost As Currency)
    
    mCost = cost
    
End Property

Public Property Get GetPrice() As Currency
    
    GetPrice = mPrice
    
End Property

Public Property Let SetPrice(price As Currency)
    
    mPrice = price
    
End Property

Public Property Get GetQuantity() As Double
    
    GetQuantity = mQuantity
    
End Property

Public Property Let SetQuantity(quantity As Double)
    
    mQuantity = quantity
    
End Property

