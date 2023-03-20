VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DatabaseInserter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub Insert(ByVal query As String)

    Dim cn As ADODB.Connection

    Set cn = Database.GetConnection()
    
    On Error GoTo err
    
    cn.BeginTrans
    cn.Execute query
    cn.CommitTrans
    
    Exit Sub
err:
    cn.RollbackTrans
End Sub


