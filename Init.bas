Attribute VB_Name = "Init"
Public Database As New DatabaseConnection

Private Sub Main()
    
    ConnectToDatabase
    
    LoadStorage
    
End Sub

Private Sub ConnectToDatabase()

    StartVisual.Show vbModal
    
End Sub

Private Sub LoadStorage()
    
    If Database.ItsPossibleToConnect Then
    
        StorageVisual.Show
        
    End If
    
End Sub
