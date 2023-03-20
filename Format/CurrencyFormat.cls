VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CurrencyFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function FormatCurrencyDollarDefault(value As Currency) As String

    FormatCurrencyDollarDefault = Format(value, "$0.00")
    
End Function

Public Function FormatCurrencySQL(value As Currency) As String

    FormatCurrencySQL = Format(value, "0.00")
    
End Function

