VERSION 5.00
Object = "{ADD24EDC-ADC1-11D2-95D1-F7A835DD4948}#3.0#0"; "NSLOCK15VB5.OCX"
Begin VB.Form FrmPrincipal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MCVL 2005 - Controle de Video Locadora"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10530
   Icon            =   "FrmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmPrincipal.frx":0CCA
   ScaleHeight     =   7080
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3480
      Top             =   2880
   End
   Begin VB.PictureBox StatusBar 
      Align           =   2  'Align Bottom
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10530
      TabIndex        =   0
      Top             =   6705
      Width           =   10530
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5040
         TabIndex        =   4
         Top             =   45
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   9000
         TabIndex        =   3
         Top             =   45
         Width           =   585
      End
      Begin VB.Label LblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   45
         Width           =   585
      End
      Begin VB.Label LblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   13680
         TabIndex        =   1
         Top             =   50
         Width           =   1065
      End
   End
   Begin nslock15vb5.ActiveLock ALock 
      Left            =   2640
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   820
      Password        =   "1ah2d3"
      SoftwareName    =   "MCVL 2005"
      LiberationKeyLength=   16
      SoftwareCodeLength=   16
   End
   Begin VB.Menu mnuMovimentos 
      Caption         =   "&Movimentos"
      Begin VB.Menu mnuLocFilmes 
         Caption         =   "Locação de Filmes"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuRetornoFilmes 
         Caption         =   "Retorno de Filmes"
         Shortcut        =   {F11}
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCadClientes 
         Caption         =   "Cadastro de Clientes"
         Begin VB.Menu mnuIncluir 
            Caption         =   "Incluir"
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnuConsultar 
            Caption         =   "Consultar"
         End
      End
      Begin VB.Menu mnuAcoes1 
         Caption         =   "Venda de Video Ações"
         Begin VB.Menu mnuLAcoes 
            Caption         =   "Lançar VideoAções"
            Shortcut        =   {F3}
         End
      End
      Begin VB.Menu mnuMiniShopping 
         Caption         =   "Venda Mini Shopping"
         Begin VB.Menu mnuLancarVenda 
            Caption         =   "Lançar Venda"
            Shortcut        =   {F4}
         End
      End
      Begin VB.Menu mnuPgDebito 
         Caption         =   "Pagamento de Débitos"
         Shortcut        =   {F5}
      End
      Begin VB.Menu bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReserva 
         Caption         =   "Reserva de Filmes"
         Begin VB.Menu mnuIncluirReservas 
            Caption         =   "Incluir Reserva de Filmes"
         End
      End
      Begin VB.Menu mnuRecCliente 
         Caption         =   "Recardo ao Cliente"
      End
   End
   Begin VB.Menu mnuGeral 
      Caption         =   "&Geral"
      Begin VB.Menu mnuCadTitulos 
         Caption         =   "Cadastro de Títulos"
         Begin VB.Menu mnuIncluir1 
            Caption         =   "Incluir"
            Shortcut        =   ^I
         End
         Begin VB.Menu mnumanutencao1 
            Caption         =   "Manutenção dos Dados"
         End
         Begin VB.Menu mnuDados1 
            Caption         =   "Consultar Dados"
         End
      End
      Begin VB.Menu mnuCadGeneros 
         Caption         =   "Cadastro de Gêneros"
         Begin VB.Menu mnuIncluir2 
            Caption         =   "Incluir"
            Shortcut        =   ^G
         End
         Begin VB.Menu mnumanutencao2 
            Caption         =   "Manutenção dos Dados"
         End
         Begin VB.Menu mnuDados2 
            Caption         =   "Consultar Dados"
         End
      End
      Begin VB.Menu mnupreco 
         Caption         =   "Cadastro de Preço"
         Begin VB.Menu mnuIncluir3 
            Caption         =   "Incluir"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnumanutencao3 
            Caption         =   "Manutenção dos Dados"
         End
         Begin VB.Menu mnuDados3 
            Caption         =   "Consultar Dados"
         End
      End
      Begin VB.Menu mnuCadFornecedores 
         Caption         =   "Cadastro de Fornecedores"
         Begin VB.Menu mnuIncluir4 
            Caption         =   "Incluir"
            Shortcut        =   ^F
         End
         Begin VB.Menu mnumanutencao4 
            Caption         =   "Manutenção dos Dados"
         End
         Begin VB.Menu mnudados4 
            Caption         =   "Consultar Dados"
         End
      End
      Begin VB.Menu mnuCadAtendentes 
         Caption         =   "Cadastro de Atendentes"
         Begin VB.Menu mnuIncluir5 
            Caption         =   "Incluir"
            Shortcut        =   ^A
         End
         Begin VB.Menu mnumanutencao5 
            Caption         =   "Manutenção dos Dados"
         End
         Begin VB.Menu mnudados5 
            Caption         =   "Consultar Dados"
         End
      End
      Begin VB.Menu mnuCadFeriados 
         Caption         =   "Cadastro de Feriados"
         Begin VB.Menu mnuIncluir6 
            Caption         =   "Incluir"
            Shortcut        =   ^E
         End
         Begin VB.Menu mnudados6 
            Caption         =   "Consultar Dados"
         End
      End
      Begin VB.Menu mnuVAcoes 
         Caption         =   "Cadastro de Videos Ações"
         Begin VB.Menu mnuIncluir7 
            Caption         =   "Incluir"
            Shortcut        =   ^V
         End
         Begin VB.Menu mnumanutencao7 
            Caption         =   "Manutenção dos Dados"
         End
         Begin VB.Menu mnudados7 
            Caption         =   "Consultar Dados"
         End
      End
      Begin VB.Menu mnuCadPShopping 
         Caption         =   "Cadastro de Prod. Shopping"
         Begin VB.Menu mnuInlcuir8 
            Caption         =   "Incluir"
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuEstoque 
            Caption         =   "Controle de Estoque"
         End
         Begin VB.Menu bar20 
            Caption         =   "-"
         End
         Begin VB.Menu mnudados8 
            Caption         =   "Consultar Dados"
         End
      End
      Begin VB.Menu mnuCadConvenios 
         Caption         =   "Cadastro de Convênios"
         Begin VB.Menu mnuIncluir9 
            Caption         =   "Incluir"
            Shortcut        =   ^N
         End
         Begin VB.Menu mnumanutencao9 
            Caption         =   "Manutenção dos Dados"
         End
         Begin VB.Menu mnudados9 
            Caption         =   "Consultar Dados"
         End
      End
      Begin VB.Menu bar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPainel 
         Caption         =   "Painel de Promoções"
         Begin VB.Menu mnu24h 
            Caption         =   "Especial 24h"
         End
         Begin VB.Menu mnu48h 
            Caption         =   "Especial 48h"
         End
         Begin VB.Menu bar21 
            Caption         =   "-"
         End
         Begin VB.Menu mnupromocao 
            Caption         =   "Pacote Promoção"
         End
         Begin VB.Menu mnualugue 
            Caption         =   "Alugue e Concorra"
         End
         Begin VB.Menu bar22 
            Caption         =   "-"
         End
         Begin VB.Menu mnudesconto 
            Caption         =   "Desconto para pgto avista"
         End
      End
      Begin VB.Menu mnuConfiguracao 
         Caption         =   "Configuração Geral"
      End
      Begin VB.Menu bar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEtiqueta 
         Caption         =   "Etiqueta em Geral"
      End
   End
   Begin VB.Menu mnuLocadora 
      Caption         =   "&Locadora"
      Begin VB.Menu mnuFila 
         Caption         =   "Fila de que o Alugou"
      End
      Begin VB.Menu mnuBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCtCaixa 
         Caption         =   "Controle de Caixa"
         Begin VB.Menu mnuretirada 
            Caption         =   "Lançar Retirada"
         End
         Begin VB.Menu mndados10 
            Caption         =   "Consultar Dados"
         End
      End
      Begin VB.Menu mnuCtDespesas 
         Caption         =   "Controle de Despesas"
         Begin VB.Menu mnudespesas1 
            Caption         =   "Lançar Despesas"
         End
         Begin VB.Menu mnuDados11 
            Caption         =   "Consultar Dados"
         End
      End
      Begin VB.Menu bar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAgenda 
         Caption         =   "Agenda Telefônica"
      End
      Begin VB.Menu mnudebitocli 
         Caption         =   "Débito de Cliente"
         Begin VB.Menu mnudebitocliente 
            Caption         =   "Lançar Débito de Cliente"
         End
         Begin VB.Menu mnuconsultar1 
            Caption         =   "Consultar Débito"
         End
      End
      Begin VB.Menu mnuAlugouAnt 
         Caption         =   "Alugou no Anterior"
      End
      Begin VB.Menu bar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFecDia 
         Caption         =   "Fechamento do Dia"
         Begin VB.Menu mnuftotal 
            Caption         =   "Fechamento total do dia"
         End
         Begin VB.Menu bar23 
            Caption         =   "-"
         End
         Begin VB.Menu mnuExtra 
            Caption         =   "Entrada Extra"
         End
         Begin VB.Menu mnuMovimentos1 
            Caption         =   "Relatorios de Movimentos"
         End
         Begin VB.Menu bar24 
            Caption         =   "-"
         End
         Begin VB.Menu mnuconferencia 
            Caption         =   "Conferência do Caixa"
         End
      End
      Begin VB.Menu mnuFecMes 
         Caption         =   "Fechamento do Mês"
      End
      Begin VB.Menu mnuFecAno 
         Caption         =   "Fechamento do Ano"
      End
      Begin VB.Menu bar8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRltExcecoes 
         Caption         =   "Relatórios Exceções"
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Begin VB.Menu mnuAlAnterior 
         Caption         =   "Alugou no Anterior"
      End
      Begin VB.Menu mnuTelefonica 
         Caption         =   "Agenda Telefônica"
      End
      Begin VB.Menu bar9 
         Caption         =   "-"
      End
      Begin VB.Menu mnupor 
         Caption         =   "Por Cadastros"
         Begin VB.Menu mnuGeneros 
            Caption         =   "Cadastro de Gêneros"
         End
         Begin VB.Menu mnuShopping 
            Caption         =   "Cadastro de Prod. Shopping"
         End
         Begin VB.Menu mnuTitulos 
            Caption         =   "Cadastro de Títulos"
         End
         Begin VB.Menu mnuConvecios 
            Caption         =   "Cadastro de Convênios"
         End
         Begin VB.Menu mnuFeriados 
            Caption         =   "Cadastro de Feriados"
         End
         Begin VB.Menu mnuAtendentes 
            Caption         =   "Cadastro de Atendentes"
         End
         Begin VB.Menu mnuPrecos 
            Caption         =   "Cadastro de Preços"
         End
         Begin VB.Menu mnuFornecedores 
            Caption         =   "Cadastro de Fornecedores"
         End
         Begin VB.Menu mnuAcoes 
            Caption         =   "Cadastro de Videos Ações"
         End
         Begin VB.Menu mnuClientes 
            Caption         =   "Cadastro de Clientes"
         End
      End
      Begin VB.Menu bar10 
         Caption         =   "-"
      End
      Begin VB.Menu mnudespesas 
         Caption         =   "Controle de Despesas"
      End
      Begin VB.Menu mnuCaixa1 
         Caption         =   "Controle de Caixa"
      End
      Begin VB.Menu bar11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLocacao 
         Caption         =   "Dados Locação"
      End
      Begin VB.Menu mnuClientes1 
         Caption         =   "Débito de Clientes"
      End
      Begin VB.Menu mnuProd 
         Caption         =   "Dados Prod. Shopping"
      End
      Begin VB.Menu bar12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoFilmes 
         Caption         =   "Locação de Filmes"
      End
      Begin VB.Menu mnuMovAcoes 
         Caption         =   "Movimento de Videos Ações"
      End
      Begin VB.Menu bar13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPgDebitos 
         Caption         =   "Pagamento de Débitos"
      End
      Begin VB.Menu bar14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResumo 
         Caption         =   "Resumo de Locações"
      End
      Begin VB.Menu mnuREtorno 
         Caption         =   "Retorno de Filmes"
      End
      Begin VB.Menu mnuRltExc 
         Caption         =   "Relatórios Exceções"
      End
      Begin VB.Menu bar15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuresCliente 
         Caption         =   "Resumo de Clientes"
      End
      Begin VB.Menu mnuReservFilm 
         Caption         =   "Reserva Filmes"
      End
      Begin VB.Menu bar16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVideoAcoes1 
         Caption         =   "Venda de Videos Ações"
      End
      Begin VB.Menu mnuMiniShopping1 
         Caption         =   "Venda Mini Shopping"
      End
   End
   Begin VB.Menu mnuUtilitarios 
      Caption         =   "&Utilitários"
      Begin VB.Menu mnuBackup 
         Caption         =   "Criar Backup"
      End
      Begin VB.Menu mnuRestaurar 
         Caption         =   "Restaurar Backup"
      End
      Begin VB.Menu bar55 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogof 
         Caption         =   "Efetuar Logoff"
      End
      Begin VB.Menu mnuUsuários 
         Caption         =   "Cadastrar Usuários"
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "&Ajuda"
      Begin VB.Menu mnuconteudo 
         Caption         =   "Conteúdo"
      End
      Begin VB.Menu mnuSobre 
         Caption         =   "Sobre o MCVL 2005"
         Shortcut        =   {F1}
      End
      Begin VB.Menu bar51 
         Caption         =   "-"
      End
      Begin VB.Menu mnuregistrar 
         Caption         =   "Registrar o MCVL 2005"
      End
   End
   Begin VB.Menu mnuFechar 
      Caption         =   "Fecha&r"
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iAnim As Integer
Dim TamStr As Integer
Const Msg As String = "Sejam Bem Vindo"

Private Sub Form_Activate()
If ALock.LastRunDate > Now Then

    MsgBox "Ocorreu uma alteração na data do sistema operacional " _
    & vbCrLf & " O programa será encerrado.", vbOKOnly + vbCritical, "Erro de sistema"
    End
End If
If Not ALock.RegisteredUser Then
     Label6.Caption = "VERSÃO DE DEMONSTRAÇÃO"
Else
     Label6.Caption = "VERSÃO REGISTRADA"
     
End If

End Sub

Private Sub Form_Initialize()
If Len(LblDate.Caption) Then
    LblDate.Caption = CVDate(Date)
    LblDate.Caption = Format$(LblDate, "Long Date")

End If
End Sub


Private Sub Form_Load()
  iAnim = 1
  TamStr = Len(Msg)
  Call Timer1_Timer
End Sub

Private Sub mnu24h_Click()
Frm24h.Show vbModal
End Sub

Private Sub mnu48h_Click()
Frm48h.Show vbModal
End Sub

Private Sub mnuAgenda_Click()
frmagenda.Show vbModal
End Sub

Private Sub mnuconferencia_Click()
FrmConferencia.Show vbModal
End Sub

Private Sub mnuConsultar_Click()
FrmConsClientes.Show vbModal
End Sub

Private Sub mnudebitocliente_Click()
frmdebito.Show vbModal
End Sub

Private Sub mnudespesas1_Click()
frmdespesas.Show vbModal
End Sub

Private Sub mnuExtra_Click()
FrmExtra.Show vbModal
End Sub

Private Sub mnuFechar_Click()
End
End Sub

Private Sub mnuFila_Click()
InputBox ("Entre com o código do filme")
End Sub

Private Sub mnuftotal_Click()
FrmFechamento.Show vbModal
End Sub

Private Sub mnuIncluir_Click()
FrmCliente.Show vbModal
End Sub

Private Sub mnuIncluir1_Click()
frmtitulos.Show vbModal
End Sub

Private Sub mnuIncluir2_Click()
frmgenero.Show vbModal
End Sub

Private Sub mnuIncluir3_Click()
frmpreco.Show vbModal
End Sub

Private Sub mnuIncluir4_Click()
FrmFornecedor.Show vbModal
End Sub

Private Sub mnuIncluir5_Click()
frmatendentes.Show vbModal
End Sub

Private Sub mnuIncluir6_Click()
frmferiado.Show vbModal
End Sub

Private Sub mnuIncluir7_Click()
frmacoes.Show vbModal
End Sub

Private Sub mnuIncluir9_Click()
frmconvenios.Show vbModal
End Sub

Private Sub mnuIncluirReservas_Click()
frmreserva.Show vbModal
End Sub

Private Sub mnuInlcuir8_Click()
frmshopping.Show vbModal
End Sub

Private Sub mnuLAcoes_Click()
frmvendaacos.Show vbModal
End Sub

Private Sub mnuLancarVenda_Click()
frmminishopping.Show vbModal
End Sub

Private Sub mnuLocFilmes_Click()
Frmlocacoes.Show vbModal
End Sub

Private Sub mnuLogof_Click()
Unload FrmPrincipal
FrmLogin.Show

End Sub

Private Sub mnuPgDebito_Click()
Frmpgto.Show vbModal
End Sub


Private Sub mnuRecCliente_Click()
Frmrecado.Show vbModal
End Sub

Private Sub mnuregistrar_Click()
Frmregistro.Show vbModal
End Sub

Private Sub mnuretirada_Click()
frmretirada.Show vbModal
End Sub

Private Sub mnuRetornoFilmes_Click()
FrmRetorno.Show vbModal
End Sub

Private Sub mnuSobre_Click()
FrmSobre.Show vbModal
End Sub

Private Sub mnuUsuários_Click()
FrmUsuarios.Show vbModal
End Sub

Private Sub Timer1_Timer()
LblTime.Caption = Time()
  FrmPrincipal.Caption = "MCVL 2005 - Controle de Video Locadora - " & Mid(Msg, 1, iAnim)
  iAnim = iAnim + 1
  If iAnim > TamStr Then
    iAnim = 1
  End If
End Sub
