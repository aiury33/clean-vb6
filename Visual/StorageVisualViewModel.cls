VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StorageVisualViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private companyController As New CompanyRepository
Private productController As New ProductRepository

Private storageCompany As Company
Private selectedProduct As Product

Private maintainStorage As New StorageMaintainence

Private curF As New CurrencyFormat

Public Property Get GetCompanyNameLabel() As String
    
    GetCompanyNameLabel = "Company: " & storageCompany.GetName
    
End Property

Public Property Get GetWalletLabel() As String
    
    GetWalletLabel = "Wallet: " & curF.FormatCurrencyDollarDefault(storageCompany.GetWallet)
    
End Property

Public Property Get GetProductNameLabel() As String
    
    GetProductNameLabel = "Product: " & selectedProduct.GetName
    
End Property

Public Property Get GetProductQuantityLabel() As String
    
    GetProductQuantityLabel = "In storage: " & selectedProduct.GetQuantity
    
End Property

Public Property Get GetProductCostLabel() As String
    
    GetProductCostLabel = "Cost: " & curF.FormatCurrencyDollarDefault(selectedProduct.GetCost)
    
End Property

Public Property Get GetProductPriceLabel() As String
    
    GetProductPriceLabel = "Price: " & curF.FormatCurrencyDollarDefault(selectedProduct.GetPrice)
    
End Property

Private Sub Class_Initialize()
    
    SearchData
    
    maintainStorage.SetProduct = selectedProduct
    
End Sub

Private Sub SearchData()

    Set storageCompany = companyController.FindCompany()
    Set selectedProduct = productController.FindProduct()
    
End Sub

Public Sub BuyProduct()

    maintainStorage.SetOperation = Buying
    maintainStorage.Alter
    SearchData
    
End Sub

Public Sub SellProduct()

    maintainStorage.SetOperation = Selling
    maintainStorage.Alter
    SearchData
    
End Sub
