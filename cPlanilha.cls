VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPlanilha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pTitulo     As String
Private pArquivo    As String
Private pFrame      As Frame
Private pGrid       As VSFlexGrid
Private pColuna     As Dictionary
Private pTela       As Form

Public Property Get Arquivo() As String
    Arquivo = pArquivo
End Property
Public Property Let Arquivo(Value As String)
    pArquivo = Value
End Property

Public Function Salvar(Optional Abrir As Boolean = False, Optional Arquivo As String) As Boolean

    If Arquivo = vbNullString Then
        DialogoArquivo
    Else
        pArquivo = Arquivo
    End If
    
    Salvar = False
    If pArquivo = vbNullString Then Exit Function

    On Error GoTo Nao_Salvo
    
    pGrid.SaveGrid pArquivo, flexFileExcel
    
    If Abrir Then
        If GFcn_MsgBox("Deseja abrir o arquivo " & vbCrLf & vbCrLf & pArquivo & "?", vbYesNo + vbInformation) = vbYes Then
            ShellExecute pTela.hwnd, "Open", pArquivo, vbNullString, "C:\", 1
        End If
    End If
    
    Salvar = True
Nao_Salvo:

End Function

Private Sub Class_Initialize()
    Set pGrid = Nothing
End Sub

Private Sub Class_Terminate()
    pTela.Controls.Remove ("GridPlanilha")
    pTela.Controls.Remove ("FrameGrid")
    Set pColuna = Nothing
    Set pGrid = Nothing
End Sub

Public Sub DialogoArquivo()
    Dim FileDialog As MSComDlg.CommonDialog
    
    Set FileDialog = pTela.Controls.Add("MSComDlg.CommonDialog", "Dialogo", pTela)
    
    On Error GoTo ErrorHandler
    With FileDialog
        If pArquivo <> vbNullString Then .FileName = pArquivo
        .Filter = "Planilha do Excel (*.xls)|*.xls|Planilha do Excel(*.xlsx)|*.xlsx"
        .CancelError = True
        .ShowSave
        
        pArquivo = .FileName
    End With
      
ErrorHandler:
    If Err.Number = 32755 Then pArquivo = Empty

fim:
    pTela.Controls.Remove ("Dialogo")
    Set FileDialog = Nothing
End Sub

Public Function IniciarCabecalho(Tela As Form, Optional Titulo As String) As cPlanilha
    Dim Licenca As String
    
    Set pTela = Tela
    
    If pGrid Is Nothing Then
        Set pFrame = pTela.Controls.Add("VB.Frame", "FrameGrid", pTela)
        Set pGrid = pTela.Controls.Add("VSFlexGrid8.VSFlexGridADO", "GridPlanilha", pFrame)
        pFrame.Visible = False
    End If
    
    Set pColuna = New Dictionary
    
    pTitulo = Titulo
    
    With pGrid
        .FixedCols = 0
        .FixedRows = 0
        .Rows = 2
        .Cols = 0
        .Font.Name = "Calibri"
        .Font.Size = 12
    End With
    
    Set IniciarCabecalho = Me
End Function

Public Function Cabecalho( _
    Coluna As String, _
    Optional Descricao As String = "", _
    Optional Tamanho As Long = 20, _
    Optional FormatType As DataTypeSettings = flexDTString _
) As cPlanilha
    Dim Propriedades(0 To 2) As Variant
    
    Propriedades(0) = IIf(Descricao = vbNullString, Coluna, Descricao)
    Propriedades(1) = Tamanho * 100
    Propriedades(2) = FormatType
    pColuna.Add Coluna, Propriedades
    Set Cabecalho = Me
End Function

Public Function IniciarPlanilha() As cPlanilha
    Dim Coluna  As Variant
    Dim Indice  As Long
    
    With pGrid
        .Cols = pColuna.Count
        For Each Coluna In pColuna.Keys
            .Col = Indice
            .ColDataType(Indice) = pColuna(Coluna)(2)
            .Row = 0
            .CellBackColor = vbYellow
            .Row = 1
            .CellBackColor = vbYellow
            .Text = pColuna(Coluna)(0)
            .ColWidth(Indice) = pColuna(Coluna)(1)
            
            pColuna(Coluna) = Indice
            
            Indice = Indice + 1
        Next
        .Rows = 3
        .Row = 2
        .TextMatrix(0, 0) = pTitulo
    End With
    
    Set IniciarPlanilha = Me
End Function

Public Function Valor(Coluna As String, Conteudo As Variant) As cPlanilha
    pGrid.TextMatrix(pGrid.Row, pColuna(Coluna)) = CStr(Conteudo)
    Set Valor = Me
End Function

Public Function NovaLinha() As cPlanilha
    pGrid.Rows = pGrid.Rows + 1
    pGrid.Row = pGrid.Rows - 1
    Set NovaLinha = Me
End Function

