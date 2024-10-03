VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExportacao 
   BackColor       =   &H00FDDEC6&
   Caption         =   "Exportar dados"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   4440
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_ArqTxt 
      BackColor       =   &H00FDDEC6&
      Caption         =   "Exportar para bloco de notas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   675
      Width           =   4005
   End
   Begin MSComctlLib.ProgressBar pgs_Status 
      Height          =   240
      Left            =   150
      TabIndex        =   1
      Top             =   1275
      Visible         =   0   'False
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmd_Planilha 
      BackColor       =   &H00FDDEC6&
      Caption         =   "Exportar para excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   150
      Width           =   4005
   End
End
Attribute VB_Name = "frmExportacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ArqTxt_Click()
    Call MPrc_ExportaTxt
End Sub

Private Sub cmd_Planilha_Click()
    Call MPrc_ExportaPlanilha
End Sub

Private Sub MPrc_ExportaPlanilha()
    Dim VLObj_Aplicacao As Excel.Application
    Dim VLObj_Work      As Excel.Workbook
    Dim VLObj_Planilha  As Excel.Worksheet
    Dim VLRst_Tabela    As ADODB.Recordset
    
    Dim VLStr_Sql       As String
    Dim VLLng_TotRow    As Long
    Dim VLInt_Cols      As Integer
    Dim VLInt_Linhas    As Integer
    
    On Local Error GoTo Erro
    
    VLStr_Sql = MFcn_MontaQuery
    Set VLRst_Tabela = VGCnx_DBConect.Execute(VLStr_Sql, VLLng_TotRow, adCmdText)
    
    If VLLng_TotRow > 0 Then
    
        Set VLObj_Aplicacao = New Excel.Application
        Set VLObj_Work = VLObj_Aplicacao.Workbooks.Add
        Set VLObj_Planilha = VLObj_Work.Sheets(1)
        
        With pgs_Status
            .Visible = True
            .Min = 0
            .Max = VLLng_TotRow
            .Value = 0
        End With
        
        If Not VLRst_Tabela.EOF Then
    
            For VLInt_Cols = 1 To VLRst_Tabela.Fields.Count
                VLObj_Planilha.Cells(1, VLInt_Cols).Value = VLRst_Tabela.Fields(VLInt_Cols - 1).Name
            Next
    
            VLRst_Tabela.MoveFirst
            
            VLInt_Linhas = 2
            
            Do While Not VLRst_Tabela.EOF
                For VLInt_Cols = 1 To VLRst_Tabela.Fields.Count
                    VLObj_Planilha.Cells(VLInt_Linhas, VLInt_Cols).Value = VLRst_Tabela.Fields(VLInt_Cols - 1).Value
                Next VLInt_Cols
                
                VLInt_Linhas = VLInt_Linhas + 1
                
                pgs_Status.Value = pgs_Status.Value + 1
                
                VLRst_Tabela.MoveNext
            Loop
        End If
    
    Else
        GFcn_MsgBox "Nenhum registro foi encontrado na seleção dos dados!"
        Exit Sub
    End If
    
    With pgs_Status
        .Visible = False
        .Value = 0
    End With
    
    GoTo Fim
Erro:
    Set VLObj_Planilha = Nothing
    Set VLObj_Work = Nothing
    Set VLObj_Aplicacao = Nothing
    
    GFcn_MsgBox "Erro ao exportar planilha!"
    
    Exit Sub
Fim:
    VLObj_Aplicacao.Visible = True
    
    Set VLObj_Planilha = Nothing
    Set VLObj_Work = Nothing
    Set VLObj_Aplicacao = Nothing
End Sub

Private Function MFcn_MontaQuery() As String
    Dim VLStr_Sql   As String
    
    VLStr_Sql = "select * from CEP"
    
    MFcn_MontaQuery = VLStr_Sql
End Function

Private Sub cmd_Planilha_GotFocus()
    cmd_Planilha.BackColor = vbYellow
End Sub

Private Sub cmd_Planilha_LostFocus()
    cmd_Planilha.BackColor = &HFDDEC6
End Sub
Private Sub cmd_ArqTxt_GotFocus()
    cmd_ArqTxt.BackColor = vbYellow
End Sub

Private Sub cmd_ArqTxt_LostFocus()
    cmd_ArqTxt.BackColor = &HFDDEC6
End Sub
Private Sub MPrc_ExportaTxt()
    Dim VLRst_Tabela    As ADODB.Recordset
    
    Dim VLStr_Sql       As String
    Dim VLStr_Texto     As String
    
    Dim VLLng_TotRow    As Long
    Dim VLInt_Cont      As Integer
    Dim VLInt_NumArq    As Integer
    
    On Local Error GoTo Erro
    
    VLStr_Sql = MFcn_MontaQuery
    Set VLRst_Tabela = VGCnx_DBConect.Execute(VLStr_Sql, VLLng_TotRow, adCmdText)
    
    VLInt_NumArq = FreeFile
    Open App.Path & "\ArquivoCEP.txt" For Output As VLInt_NumArq
    
    If VLLng_TotRow > 0 Then
    
        With pgs_Status
            .Visible = True
            .Min = 0
            .Max = VLLng_TotRow
            .Value = 0
        End With
        
        VLStr_Texto = ""
        
        For VLInt_Cont = 0 To VLRst_Tabela.Fields.Count - 1
            VLStr_Texto = VLStr_Texto & VLRst_Tabela.Fields(VLInt_Cont).Name
            
            If VLInt_Cont < VLRst_Tabela.Fields.Count - 1 Then
                VLStr_Texto = VLStr_Texto & vbTab
            End If
        Next VLInt_Cont
        
        Print #VLInt_NumArq, VLStr_Texto
        
        VLRst_Tabela.MoveFirst
        
        Do While Not VLRst_Tabela.EOF
            VLStr_Texto = ""
            
            For VLInt_Cont = 0 To VLRst_Tabela.Fields.Count - 1
                VLStr_Texto = VLStr_Texto & Trim$(VLRst_Tabela.Fields(VLInt_Cont).Value)
                
                If VLInt_Cont < VLRst_Tabela.Fields.Count - 1 Then
                    VLStr_Texto = VLStr_Texto & vbTab
                End If
            Next VLInt_Cont
            
            Print #VLInt_NumArq, VLStr_Texto
            
            pgs_Status.Value = pgs_Status.Value + 1
            
            VLRst_Tabela.MoveNext
        Loop
    Else
        GFcn_MsgBox "Nenhum registro foi encontrado na seleção dos dados!"
        Exit Sub
    End If
    
    Close #VLInt_NumArq
    
    With pgs_Status
        .Visible = False
        .Value = 0
    End With
    
    GoTo Fim
Erro:
    GFcn_MsgBox "Erro ao exportar arquivo de texto!"
    Exit Sub
Fim:
    GFcn_MsgBox "Arquivo gerado com sucesso!" & vbCrLf & vbCrLf & "Se encontra no diretório: " & App.Path & "\ArquivoCEP.txt"
End Sub
