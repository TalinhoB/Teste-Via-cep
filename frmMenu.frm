VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmMenu 
   BackColor       =   &H8000000A&
   Caption         =   "Consumir API"
   ClientHeight    =   6330
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   9405
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar stts_Barra 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   6015
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu CadastrarCEP 
      Caption         =   "Cadastrar"
   End
   Begin VB.Menu Exportacao 
      Caption         =   "Exportação"
   End
   Begin VB.Menu Sair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CadastrarCEP_Click()
    frmCadastrarCEP.Show
End Sub

Private Sub Exportacao_Click()
    frmExportacao.Show
End Sub

Private Sub MDIForm_Load()
    Dim VLStr_Data      As String
    Dim VLStr_NomeUsr   As String
    
    Set VMObj_Conexao = New clsConexao
    VMObj_Conexao.Conectar
    
    VLStr_Data = "Data " & Date
    VLStr_NomeUsr = "Usuário: " & GFcn_PegaNomeUsuario
    
    stts_Barra.Panels.Item(1).Text = VLStr_Data
    stts_Barra.Panels.Item(2).Text = VLStr_NomeUsr
    
End Sub

Private Sub Sair_Click()
    Unload Me
End Sub
