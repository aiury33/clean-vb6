VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StartVisualViewModel"
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

Public Function ConnectToDatabase() As Boolean
    
    SetLoginData
    
    ConnectToDatabase = ValidateConnection
    
End Function

Private Sub SetLoginData()

    Database.SetServer = mServer
    Database.SetLogin = mLogin
    Database.SetPassword = mPassword

End Sub

Private Function ValidateConnection() As Boolean

    ValidateConnection = Database.Connected
    
End Function

