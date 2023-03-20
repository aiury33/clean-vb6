VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mID As Long
Private mName As String
Private mWallet As Currency

Public Property Get GetID() As Long
    
    GetID = mID
    
End Property

Public Property Let SetID(id As Long)
    
    mID = id
    
End Property

Public Property Get GetName() As String
    
    GetName = mName
    
End Property

Public Property Let SetName(name As String)
    
    mName = name
    
End Property

Public Property Get GetWallet() As Currency
    
    GetWallet = mWallet
    
End Property

Public Property Let SetWallet(wallet As Currency)
    
    mWallet = wallet
    
End Property
