VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DatabaseConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mServer As String
Private mLogin As String
Private mPassword As String

Public Property Let SetServer(server As String)
    
    mServer = server
    
End Property

Public Property Let SetLogin(login As String)
    
    mLogin = login
    
End Property

Public Property Let SetPassword(password As String)
    
    mPassword = password
    
End Property

Public Function Connected() As Boolean

    Dim cn As ADODB.Connection
    
    On Error GoTo err
    
    Set cn = GetConnection()
    
    Connected = True
    
err:
End Function

Public Function GetConnection() As ADODB.Connection

    Dim cn As New ADODB.Connection

    cn.ConnectionString = GetConnectionString()
    cn.Open
    
    Set GetConnection = cn
    
End Function

Private Function GetConnectionString() As String

    GetConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & mLogin & ";Password=" & mPassword & ";Initial Catalog=clean;Data Source=" & mServer

End Function
