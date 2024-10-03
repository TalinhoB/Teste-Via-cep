VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConexao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MyName As String = "clsConexao"

Public Sub Conectar()
    Dim VLStr_StringConx    As String
    
    On Local Error GoTo Erro
    
    If ConexaoAtiva Then
        Call Desconectar
    End If
    
    Set VGCnx_DBConect = New ADODB.Connection
    
    VLStr_StringConx = "DSN=PostgreSQL35W"
    
    VGCnx_DBConect.ConnectionString = VLStr_StringConx
    VGCnx_DBConect.Open
    
    GoTo Fim
Erro:
    Err.Raise Err.Number, "Classe Conectar - Erro no método Conectar", "Erro: " & Err.Description & " na conexão "
Fim:

End Sub

Private Property Get ConexaoAtiva() As Boolean
    If Not VGCnx_DBConect Is Nothing Then
        ConexaoAtiva = (VGCnx_DBConect.State = adStateOpen)
    Else
        ConexaoAtiva = False
    End If
End Property
Private Sub Desconectar()
    If Not VGCnx_DBConect Is Nothing Then
        VGCnx_DBConect.Close
        Set VGCnx_DBConect = Nothing
    End If
End Sub


