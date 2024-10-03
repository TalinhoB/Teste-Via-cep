VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCadastrarCEP 
   BackColor       =   &H00FDDEC6&
   Caption         =   "CEP"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7770
   Icon            =   "frmCEP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   7770
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_Limpar 
      BackColor       =   &H00FDDEC6&
      Caption         =   "&Limpar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   900
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3375
      Width           =   855
   End
   Begin VB.CommandButton cmd_Gravar 
      BackColor       =   &H00FDDEC6&
      Caption         =   "&Gravar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3375
      Width           =   855
   End
   Begin VB.Frame fra_Dados 
      BackColor       =   &H00FDDEC6&
      Caption         =   "Seleção"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3240
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   7590
      Begin VB.TextBox txt_CodigoGIA 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6000
         TabIndex        =   25
         Top             =   2775
         Width           =   1275
      End
      Begin VB.TextBox txt_CodigoSIAFI 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3450
         TabIndex        =   23
         Top             =   2775
         Width           =   1275
      End
      Begin VB.TextBox txt_DDDRegiao 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1650
         TabIndex        =   21
         Top             =   2775
         Width           =   540
      End
      Begin VB.TextBox txt_CodIBGE 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5175
         TabIndex        =   19
         Top             =   2400
         Width           =   2115
      End
      Begin VB.TextBox txt_Regiao 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1650
         TabIndex        =   17
         Top             =   2400
         Width           =   2115
      End
      Begin VB.TextBox txt_Estado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3225
         TabIndex        =   15
         Top             =   2025
         Width           =   2115
      End
      Begin VB.TextBox txt_UF 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1650
         TabIndex        =   14
         Top             =   2025
         Width           =   540
      End
      Begin VB.TextBox txt_Localidade 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2775
         TabIndex        =   12
         Top             =   480
         Width           =   4515
      End
      Begin VB.TextBox txt_Bairro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1650
         TabIndex        =   11
         Top             =   1650
         Width           =   3690
      End
      Begin VB.TextBox txt_Unidade 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5175
         TabIndex        =   9
         Top             =   1275
         Width           =   2115
      End
      Begin VB.TextBox txt_Complemento 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1650
         TabIndex        =   7
         Top             =   1275
         Width           =   2565
      End
      Begin VB.TextBox txt_Logradouro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1650
         TabIndex        =   4
         Top             =   900
         Width           =   5640
      End
      Begin MSMask.MaskEdBox msk_CepCodigo 
         Height          =   315
         Left            =   1650
         TabIndex        =   0
         Top             =   480
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "#####-###"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_Titulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Código GIA:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   4950
         TabIndex        =   26
         Top             =   2850
         Width           =   1140
      End
      Begin VB.Label lbl_Titulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Código SIAFI:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   2325
         TabIndex        =   24
         Top             =   2850
         Width           =   1140
      End
      Begin VB.Label lbl_Titulo 
         BackStyle       =   0  'Transparent
         Caption         =   "DDD:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   225
         TabIndex        =   22
         Top             =   2850
         Width           =   915
      End
      Begin VB.Label lbl_Titulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Código IBGE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   3900
         TabIndex        =   20
         Top             =   2475
         Width           =   1215
      End
      Begin VB.Label lbl_Titulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Região:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   225
         TabIndex        =   18
         Top             =   2475
         Width           =   915
      End
      Begin VB.Label lbl_Titulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   2475
         TabIndex        =   16
         Top             =   2100
         Width           =   915
      End
      Begin VB.Label lbl_Titulo 
         BackStyle       =   0  'Transparent
         Caption         =   "UF:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   225
         TabIndex        =   13
         Top             =   2100
         Width           =   1365
      End
      Begin VB.Label lbl_Titulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   225
         TabIndex        =   10
         Top             =   1725
         Width           =   1365
      End
      Begin VB.Label lbl_Titulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidade:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   4350
         TabIndex        =   8
         Top             =   1350
         Width           =   1365
      End
      Begin VB.Label lbl_Titulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   225
         TabIndex        =   6
         Top             =   1350
         Width           =   1365
      End
      Begin VB.Label lbl_Titulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Logradouro:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   225
         TabIndex        =   5
         Top             =   975
         Width           =   1065
      End
      Begin VB.Label lbl_Titulo 
         BackStyle       =   0  'Transparent
         Caption         =   "CEP:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   2
         Top             =   525
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmCadastrarCEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim VMStr_CEPCodigo As String
Dim VMStr_MaskInit  As String

Dim VMObj_Conexao   As clsConexao

Private Sub Command1_Click()

End Sub

Private Sub cmd_Gravar_Click()
    If msk_CepCodigo.Text = vbNullString And InStr(VMStr_MaskInit, msk_CepCodigo.Text) = 0 Then
        GFcn_MsgBox "Busque por um CEP antes de gravar!"
        Exit Sub
    Else
        Call MPrc_GravarDados
    End If
End Sub

Private Sub cmd_Gravar_GotFocus()
    cmd_Gravar.BackColor = vbYellow
End Sub

Private Sub cmd_Gravar_LostFocus()
    cmd_Gravar.BackColor = &HFDDEC6
End Sub

Private Sub cmd_Limpar_Click()
    Call MPrc_LimparCampos
End Sub

Private Sub cmd_Limpar_GotFocus()
    cmd_Limpar.BackColor = vbYellow
End Sub

Private Sub cmd_Limpar_LostFocus()
    cmd_Limpar.BackColor = &HFDDEC6
End Sub

Private Sub Form_Load()
    VMStr_MaskInit = msk_CepCodigo.Text
End Sub

Private Sub msk_CepCodigo_GotFocus()
    msk_CepCodigo.BackColor = vbYellow
End Sub

Private Sub msk_CepCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeysTab
    End If
End Sub

Private Sub msk_CepCodigo_LostFocus()
    msk_CepCodigo.BackColor = vbWhite
    
    VMStr_CEPCodigo = Format(Replace(msk_CepCodigo.Text, "-", ""), "00000000")
    
    If VMStr_CEPCodigo <> vbNullString And InStr(VMStr_MaskInit, msk_CepCodigo.Text) = 0 Then
        Call ConsultarCEP(VMStr_CEPCodigo)
    End If
End Sub
Private Sub ConsultarCEP(PPStr_CepCodigo As String)
    Dim VLObj_Requisicao    As Object
    Dim VLObj_RespostaRqs   As Object
    
    Dim VLStr_UrlAPI        As String
    Dim VLStr_Resposta      As String
    
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    
    VLStr_UrlAPI = "https://viacep.com.br/ws/" & PPStr_CepCodigo & "/json/"
    Set VLObj_Requisicao = CreateObject("MSXML2.XMLHTTP")

    With VLObj_Requisicao
        .Open "GET", VLStr_UrlAPI, False
        .send

        If .Status = 200 Then
            VLStr_Resposta = .responseText

            Set VLObj_RespostaRqs = JsonConverter.ParseJson(VLStr_Resposta)

            txt_Logradouro.Text = VLObj_RespostaRqs("logradouro")
            txt_Localidade.Text = VLObj_RespostaRqs("localidade")
            txt_UF.Text = VLObj_RespostaRqs("uf")
            txt_Complemento.Text = VLObj_RespostaRqs("complemento")
            txt_Bairro.Text = VLObj_RespostaRqs("bairro")
            txt_Unidade.Text = VLObj_RespostaRqs("unidade")
            txt_Estado.Text = VLObj_RespostaRqs("estado")
            txt_CodIBGE.Text = VLObj_RespostaRqs("ibge")
            txt_CodigoGIA.Text = VLObj_RespostaRqs("gia")
            txt_CodigoSIAFI.Text = VLObj_RespostaRqs("siafi")
            txt_DDDRegiao.Text = VLObj_RespostaRqs("ddd")
            txt_Regiao.Text = VLObj_RespostaRqs("regiao")
            
        Else
            GFcn_MsgBox "CEP Não encontrado!"
            msk_CepCodigo.SetFocus
            Exit Sub
        End If
    End With
    
    Screen.MousePointer = vbDefault
    Me.Enabled = True
End Sub
Private Sub MPrc_GravarDados()
    Dim VLStr_Sql   As String
    Dim VLStr_Cols  As String
    
    On Local Error GoTo Erro
    
    If MFcn_ExisteRegistro Then Exit Sub
    
    VGCnx_DBConect.BeginTrans
    
    VLStr_Cols = "cep,logradouro,complemento,unidade,bairro,localidade,uf,estado,regiao,ibge,gia,ddd,siafi"
    
    VLStr_Sql = "insert into CEP (" & VLStr_Cols & ") values ('" & _
            Left(VMStr_CEPCodigo & Space(9), 9) & "','" & _
            Left(txt_Logradouro.Text & Space(255), 255) & "','" & _
            Left(txt_Complemento.Text & Space(255), 255) & "','" & _
            Left(txt_Unidade.Text & Space(255), 255) & "','" & _
            Left(txt_Bairro.Text & Space(100), 100) & "','" & _
            Left(txt_Localidade.Text & Space(100), 100) & "','" & _
            Left(txt_UF.Text & Space(2), 2) & "','" & _
            Left(txt_Estado.Text & Space(100), 100) & "','" & _
            Left(txt_Regiao.Text & Space(100), 100) & "','" & _
            Left(txt_CodIBGE.Text & Space(7), 7) & "','" & _
            Left(txt_CodigoGIA.Text & Space(10), 10) & "','" & _
            Left(txt_DDDRegiao.Text & Space(3), 3) & "','" & _
            Left(txt_CodigoSIAFI.Text & Space(4), 4) & "'" & _
        ")"
        
    VGCnx_DBConect.Execute VLStr_Sql

    GoTo Fim
Erro:
    GFcn_MsgBox "Erro ao gravar informações!" & vbCrLf & vbCrLf & Err.Description
    VGCnx_DBConect.RollbackTrans
    Exit Sub
Fim:
    VGCnx_DBConect.CommitTrans
    GFcn_MsgBox "Gravação Ok!"
    Call MPrc_LimparCampos
End Sub
Private Function MFcn_ExisteRegistro() As Boolean
    Dim VLRst_Tabela    As ADODB.Recordset
    Dim VLStr_Sql       As String
    
    Dim VLLng_TotRow    As Long
    
    VLStr_Sql = "select cep from CEP where cep = '" & Left(VMStr_CEPCodigo & Space(9), 9) & "' "
    Set VLRst_Tabela = VGCnx_DBConect.Execute(VLStr_Sql, VLLng_TotRow, adCmdText)
    
    If VLLng_TotRow > 0 Then
        GFcn_MsgBox "Registro já existe na base de dados!"
        MFcn_ExisteRegistro = True
    End If
    
    VLRst_Tabela.Close
    Set VLRst_Tabela = Nothing
    
End Function
Private Sub MPrc_LimparCampos()
    msk_CepCodigo.Text = VMStr_MaskInit
    txt_Logradouro.Text = vbNullString
    txt_Localidade.Text = vbNullString
    txt_UF.Text = vbNullString
    txt_Complemento.Text = vbNullString
    txt_Bairro.Text = vbNullString
    txt_Unidade.Text = vbNullString
    txt_Estado.Text = vbNullString
    txt_CodIBGE.Text = vbNullString
    txt_CodigoGIA.Text = vbNullString
    txt_CodigoSIAFI.Text = vbNullString
    txt_DDDRegiao.Text = vbNullString
    txt_Regiao.Text = vbNullString
    
    msk_CepCodigo.SetFocus
End Sub
