VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DatabaseFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Find(ByVal query As String) As ADODB.Recordset

    Dim cn As ADODB.Connection
    Dim rs As New ADODB.Recordset

    Set cn = Database.GetConnection()
    
    rs.Open query, cn, adOpenStatic, adLockReadOnly
    
    Set Find = rs
    
End Function

