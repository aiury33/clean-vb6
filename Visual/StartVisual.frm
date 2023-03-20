VERSION 5.00
Begin VB.Form StartVisual 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect to database"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox ServerText 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton ConnectButton 
      Appearance      =   0  'Flat
      Caption         =   "Connect"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox PasswordText 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2880
      Width           =   3615
   End
   Begin VB.TextBox LoginText 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label ServerLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Server"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label PasswordLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Password"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label LoginLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Login"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "StartVisual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private viewModel As New StartVisualViewModel

Private Sub ConnectButton_Click()
    
    ConnectToDatabase
    
End Sub

Private Sub ConnectToDatabase()

    Dim Connected As Boolean
    
    SetLoginData
    
    Connected = viewModel.ConnectToDatabase
    
    ShowMessageByStatus Connected
    
    ExitWhenSuccess Connected

End Sub

Private Sub SetLoginData()

    viewModel.SetServer = ServerText.Text
    viewModel.SetLogin = LoginText.Text
    viewModel.SetPassword = PasswordText.Text

End Sub

Private Sub ExitWhenSuccess(Connected As Boolean)

    If Connected Then
    
        Unload Me
        
    End If
    
End Sub

Private Sub ShowMessageByStatus(Connected As Boolean)
    
    Dim messageToShow As String
    
    If Connected Then
    
        messageToShow = "Success."
        
    Else

        messageToShow = "Failed."
        
    End If
    
    MsgBox messageToShow
    
End Sub
