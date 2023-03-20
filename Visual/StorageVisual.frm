VERSION 5.00
Begin VB.Form StorageVisual 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Storage"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton SellButton 
      Caption         =   "Sell"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton BuyButton 
      Caption         =   "Buy"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label PriceLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Price: $0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label CostLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cost: $0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label StorageQuantityLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Storage: 0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   9135
   End
   Begin VB.Label WalletLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Wallet : $0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   9135
   End
   Begin VB.Label CompanyNameLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Market"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9135
   End
   Begin VB.Label ProductNameLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Potato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   9135
   End
End
Attribute VB_Name = "StorageVisual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private viewModel As New StorageVisualViewModel

Private Sub Form_Load()

    FillLabels
    
End Sub

Private Sub BuyButton_Click()
    
    viewModel.BuyProduct
    
    FillLabels
    
End Sub

Private Sub SellButton_Click()

    viewModel.SellProduct
    
    FillLabels
    
End Sub

Private Sub FillLabels()
    
    FillCompanyLabels
    FillProductLabels
    
End Sub

Private Sub FillCompanyLabels()
    
    CompanyNameLabel.Caption = viewModel.GetCompanyNameLabel
    WalletLabel.Caption = viewModel.GetWalletLabel
    
End Sub

Private Sub FillProductLabels()
    
    ProductNameLabel.Caption = viewModel.GetProductNameLabel
    StorageQuantityLabel.Caption = viewModel.GetProductQuantityLabel
    CostLabel.Caption = viewModel.GetProductCostLabel
    PriceLabel.Caption = viewModel.GetProductPriceLabel
    
End Sub
