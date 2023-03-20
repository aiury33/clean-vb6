VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mID As Long
Private mCompanyID As Long
Private mName As String
Private mCost As Currency
Private mPrice As Currency
Private mQuantity As Double

Public Property Get GetID() As Long
    
    GetID = mID
    
End Property

Public Property Let SetID(id As Long)
    
    mID = id
    
End Property

Public Property Get GetCompanyID() As Long
    
    GetCompanyID = mCompanyID
    
End Property

Public Property Let SetCompanyID(id As Long)
    
    mCompanyID = id
    
End Property

Public Property Get GetName() As String
    
    GetName = mName
    
End Property

Public Property Let SetName(name As String)
    
    mName = name
    
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
