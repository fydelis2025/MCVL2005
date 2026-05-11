VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MCVL 2005 - Cadastro de Clientes"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11310
   Icon            =   "FrmCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   "C:\Arquivos de programas\MCVL 2005\LocVideo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   9240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Clientes"
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton CmdSalvar 
      Height          =   495
      Left            =   10080
      Picture         =   "FrmCliente.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton CmdNovo 
      Height          =   495
      Left            =   9600
      Picture         =   "FrmCliente.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   2400
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   5055
      Left            =   0
      ScaleHeight     =   4995
      ScaleWidth      =   11235
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.PictureBox PWeb 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "Foto"
         DataSource      =   "Data1"
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   9360
         MouseIcon       =   "FrmCliente.frx":265E
         MousePointer    =   99  'Custom
         Picture         =   "FrmCliente.frx":2968
         ScaleHeight     =   1665
         ScaleWidth      =   1665
         TabIndex        =   27
         Top             =   360
         Width           =   1695
      End
      Begin MSMask.MaskEdBox Codigo 
         DataField       =   "Codigo"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Nome 
         DataField       =   "Nome"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   840
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Nascimento 
         DataField       =   "Nascimento"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   6120
         TabIndex        =   3
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Inscricao 
         DataField       =   "Inscricao"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   7560
         TabIndex        =   4
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Endereco 
         DataField       =   "Endereco"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Bairro 
         DataField       =   "Bairro"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   4560
         TabIndex        =   6
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Cidade 
         DataField       =   "Cidade"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   6360
         TabIndex        =   7
         Top             =   1440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Estado 
         DataField       =   "Estado"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   8280
         TabIndex        =   8
         Top             =   1440
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   2
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Fone1 
         DataField       =   "Fone1"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   9
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Fone2 
         DataField       =   "Fone2"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   2040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   9
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Celular 
         DataField       =   "Celular"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   3000
         TabIndex        =   11
         Top             =   2040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   9
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox NomeContato 
         DataField       =   "Contato"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   4440
         TabIndex        =   12
         Top             =   2040
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   11
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Fone3 
         DataField       =   "Fone3"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   7560
         TabIndex        =   13
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   9
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox CPF 
         DataField       =   "CPF"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox RG 
         DataField       =   "RG"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1560
         TabIndex        =   15
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   15
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox CEP 
         DataField       =   "CEP"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   3120
         TabIndex        =   16
         Top             =   2640
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Email 
         DataField       =   "Email"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   4320
         TabIndex        =   17
         Top             =   2640
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox EndTrabalho 
         DataField       =   "End_Trabalho"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   3240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Responssavel 
         DataField       =   "Responsavel"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   3840
         TabIndex        =   19
         Top             =   3240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Fone5 
         DataField       =   "Fone4"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   6720
         TabIndex        =   20
         Top             =   3240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   9
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Conjuge 
         DataField       =   "Conjuge"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   3840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox FoneConjuge 
         DataField       =   "Fone_Conjuge"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   3840
         TabIndex        =   22
         Top             =   3840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   9
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Nacimento1 
         DataField       =   "Nascimento1"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   5280
         TabIndex        =   23
         Top             =   3840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Limite 
         DataField       =   "Limite_Debito"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   6600
         TabIndex        =   24
         Top             =   3840
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Credito 
         DataField       =   "Limite_Credito"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   4560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   50
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Convenio 
         DataField       =   "Convenio"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1680
         TabIndex        =   26
         Top             =   4560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Localizar"
         Height          =   195
         Index           =   0
         Left            =   7440
         TabIndex        =   67
         Top             =   4800
         Width           =   630
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   7440
         TabIndex        =   66
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar"
         Height          =   195
         Index           =   3
         Left            =   5880
         TabIndex        =   65
         Top             =   4800
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alterar"
         Height          =   195
         Index           =   2
         Left            =   5220
         TabIndex        =   64
         Top             =   4800
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excluir"
         Height          =   195
         Index           =   1
         Left            =   4500
         TabIndex        =   63
         Top             =   4800
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incuir"
         Height          =   195
         Index           =   5
         Left            =   3825
         TabIndex        =   62
         Top             =   4800
         Width           =   390
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   5880
         TabIndex        =   61
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   5160
         TabIndex        =   60
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   4440
         TabIndex        =   59
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   3720
         TabIndex        =   58
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   6600
         TabIndex        =   57
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gravar"
         Height          =   195
         Index           =   6
         Left            =   6720
         TabIndex        =   56
         Top             =   4800
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fechar"
         Height          =   195
         Index           =   7
         Left            =   8280
         TabIndex        =   55
         Top             =   4800
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   8
         Left            =   8280
         TabIndex        =   54
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   53
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome "
         Height          =   195
         Index           =   1
         Left            =   1320
         TabIndex        =   52
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nascimento"
         Height          =   195
         Index           =   2
         Left            =   6120
         TabIndex        =   51
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   50
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro"
         Height          =   195
         Index           =   4
         Left            =   4560
         TabIndex        =   49
         Top             =   1200
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade"
         Height          =   195
         Index           =   5
         Left            =   6360
         TabIndex        =   48
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
         Height          =   195
         Index           =   6
         Left            =   8280
         TabIndex        =   47
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone1"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   46
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone2"
         Height          =   195
         Index           =   8
         Left            =   1560
         TabIndex        =   45
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Celular"
         Height          =   195
         Index           =   9
         Left            =   3000
         TabIndex        =   44
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Contato"
         Height          =   195
         Index           =   10
         Left            =   4440
         TabIndex        =   43
         Top             =   1800
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone Contato"
         Height          =   195
         Index           =   11
         Left            =   7560
         TabIndex        =   42
         Top             =   1800
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   41
         Top             =   2400
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RG"
         Height          =   195
         Index           =   13
         Left            =   1560
         TabIndex        =   40
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CEP"
         Height          =   195
         Index           =   14
         Left            =   3120
         TabIndex        =   39
         Top             =   2400
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inscrição"
         Height          =   195
         Index           =   15
         Left            =   7560
         TabIndex        =   38
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail"
         Height          =   195
         Index           =   16
         Left            =   4320
         TabIndex        =   37
         Top             =   2400
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço Trabalho"
         Height          =   195
         Index           =   17
         Left            =   120
         TabIndex        =   36
         Top             =   3000
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pessoa Responsável"
         Height          =   195
         Index           =   18
         Left            =   3840
         TabIndex        =   35
         Top             =   3000
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone"
         Height          =   195
         Index           =   19
         Left            =   6720
         TabIndex        =   34
         Top             =   3000
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limite de Débito"
         Height          =   195
         Index           =   20
         Left            =   6600
         TabIndex        =   33
         Top             =   3600
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limite de Crédito"
         Height          =   195
         Index           =   21
         Left            =   120
         TabIndex        =   32
         Top             =   4320
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Convênio"
         Height          =   195
         Index           =   22
         Left            =   1680
         TabIndex        =   31
         Top             =   4320
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cônjuge"
         Height          =   195
         Index           =   24
         Left            =   120
         TabIndex        =   30
         Top             =   3600
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fone Cônjuge"
         Height          =   195
         Index           =   25
         Left            =   3840
         TabIndex        =   29
         Top             =   3600
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nascimento"
         Height          =   195
         Index           =   26
         Left            =   5280
         TabIndex        =   28
         Top             =   3600
         Width           =   840
      End
   End
End
Attribute VB_Name = "FrmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Bairro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub celular_Change()
 If Len(Celular) = 4 Then
      Celular = Celular & "-"
      Celular.SelStart = 5
   End If
      If Len(Celular) = 8 Then
      Celular = Format(Celular, "####-####")
   End If
End Sub


Private Sub Celular_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub


Private Sub celular_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub CEP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub CEP_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub Cidade_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Cliente_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
Select Case dataerr

         Case 3044 'caminho invalido
            MsgBox "O caminho informado não é valido, verifique !"
            End 'uma maneira muito rude de interromper sua aplicacao, mas...

         Case 3024 'nao achou banco de dados
            MsgBox "O arquivo definido não foi encontrado !"
            End

         Case Else
         'MsgBox "Erro  em  : " & Error$(dataerr) 'para mostrar o erro.
         response = vbDataErrContinue  'ou vbdataErrdisplay

    End Select

End Sub

Private Sub CmdNovo_Click()
Data1.Recordset.AddNew
Codigo.Text = Data1.Recordset.RecordCount + 1
Codigo.SetFocus
End Sub

Private Sub CmdSalvar_Click()
        If Codigo.Text = Empty Then
            MsgBox "Informe o codigo do cliente!", vbInformation
        Exit Sub
        End If
        
        If Nome.Text = Empty Then
            MsgBox "Informe o nome do cliente!", vbInformation
        Exit Sub
        End If
        
        If Nascimento.Text = Empty Then
            MsgBox "Por favor, informe a data de nascimento do cliente!", vbInformation
        Exit Sub
        End If
        
        If Inscricao.Text = Empty Then
            MsgBox "Informe a data de inscrição do cliente!", vbInformation
        Exit Sub
        End If
        
            res = MsgBox("Deseja confirmar o registro?", _
                48 + vbYesNo, "Gravar registro")
            If res = vbYes Then
              Data1.UpdateRecord
              Data1.Recordset.Bookmark = Data1.Recordset.LastModified
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
            End If
        End If
End Sub

Private Sub codigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub codigo_LostFocus()
Codigo.Text = Format(Codigo, "000000")

End Sub

Private Sub Conjuge_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Convenio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
        If Codigo.Text = Empty Then
            MsgBox "Informe o codigo do cliente!", vbInformation
        Exit Sub
        End If
        
        If Nome.Text = Empty Then
            MsgBox "Informe o nome do cliente!", vbInformation
        Exit Sub
        End If
        
        If Nascimento.Text = Empty Then
            MsgBox "Por favor, informe a data de nascimento do cliente!", vbInformation
        Exit Sub
        End If
        
        If Inscricao.Text = Empty Then
            MsgBox "Informe a data de inscrição do cliente!", vbInformation
        Exit Sub
        End If
        
            res = MsgBox("Deseja confirmar o registro?", _
                48 + vbYesNo, "Gravar registro")
            If res = vbYes Then
              Data1.UpdateRecord
              Data1.Recordset.Bookmark = Data1.Recordset.LastModified
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
            End If
        End If
End If
End Sub


Private Sub CPF_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub CPF_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub Credito_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    Credito.Text = Format(Credito, "Currency")
End If
End Sub

Private Sub Credito_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
        Case Is = 8    ' Backspace
        Case 48 To 57  ' Numeros de 0 a 9
        Case 44
        Case 46
        Case Else
        KeyAscii = 0
    End Select

End Sub

Private Sub Email_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Endereco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub EndTrabalho_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Estado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Fone1_Change()
 If Len(Fone1) = 4 Then
      Fone1 = Fone1 & "-"
      Fone1.SelStart = 5
   End If
      If Len(Fone1) = 8 Then
      Fone1 = Format(Fone1, "####-####")
   End If
End Sub

Private Sub Fone1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Fone1_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub Fone2_Change()
 If Len(Fone2) = 4 Then
      Fone2 = Fone2 & "-"
      Fone2.SelStart = 5
   End If
      If Len(Fone2) = 8 Then
      Fone2 = Format(Fone2, "####-####")
   End If
End Sub

Private Sub Fone2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Fone2_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub Fone3_Change()
 If Len(Fone3) = 4 Then
      Fone3 = Fone3 & "-"
      Fone3.SelStart = 5
   End If
      If Len(Fone3) = 8 Then
      Fone3 = Format(Fone3, "####-####")
   End If
End Sub

Private Sub Fone3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Fone3_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub Fone5_Change()
 If Len(Fone5) = 4 Then
      Fone5 = Fone5 & "-"
      Fone5.SelStart = 5
   End If
      If Len(Fone5) = 8 Then
      Fone5 = Format(Fone5, "####-####")
   End If
End Sub

Private Sub Fone5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub FoneConjuge_Change()
 If Len(FoneConjuge) = 4 Then
      FoneConjuge = FoneConjuge & "-"
      FoneConjuge.SelStart = 5
   End If
      If Len(FoneConjuge) = 8 Then
      FoneConjuge = Format(FoneConjuge, "####-####")
   End If
End Sub

Private Sub FoneConjuge_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub FoneConjuge_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub Form_Activate()
'Inicia a câmera
mCapHwnd = capCreateCaptureWindow("My Own Capture Window", 0, 0, 0, 320, 240, Me.hwnd, 0)
SendMessage mCapHwnd, WM_CAP_DRIVER_CONNECT, 0, 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim strPntr As String
    strPntr = "O código do cliente é gerado automaticamente"
    colocatip Codigo, strPntr, Me
    strPntr = "Insira a data de nascimento!"
    colocatip Nascimento, strPntr, Me
    strPntr = "Insira a data de inscrição!"
    colocatip Inscricao, strPntr, Me
    strPntr = "Informe o limite de débito para o cliente!"
    colocatip Limite, strPntr, Me
    strPntr = "Informe o limite de crédito para o cliente!"
    colocatip Credito, strPntr, Me
    strPntr = "Click aqui para exibir uma webcam!"
    colocatip PWeb, strPntr, Me
End Sub

Private Sub Form_Terminate()
'Desliga a câmera
SendMessage mCapHwnd, WM_CAP_DRIVER_DISCONNECT, 0, 0
End Sub

Private Sub Frame1_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub Inscricao_Change()
If Len(Inscricao) = 2 Then
      Inscricao = Inscricao & "/"
      Inscricao.SelStart = 4
   End If
   If Len(Inscricao) = 5 Then
      Inscricao = Inscricao & "/"
      Inscricao.SelStart = 7
   End If
   If Len(Inscricao) = 8 Then
      Inscricao = Format(Inscricao, "dd/mm/yyyy")
   End If
End Sub

Private Sub Inscricao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Inscricao_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub Label3_Click(Index As Integer)
Select Case KeyCode
        Case vbKeyF1 'incluir
          Cliente.Recordset.AddNew
          Nome.SetFocus
        Case vbKeyF2 'Excluir registro atual
         Dim res
         res = MsgBox("Deseja excluir o registro atual?", _
         48 + vbYesNo, "Excluir???")
         If res = vbYes Then
          Cliente.Recordset.Delete
        If Not Cliente.Recordset.EOF Then
          Cliente.Recordset.MoveFirst
      End If
    End If
        Case vbKeyF3 'alterar um registro
            Nome.SetFocus
            
        Case vbKeyF4 'cancelar
            
            res = MsgBox("Deseja realmente cancelar este registro?", _
                48 + vbYesNo, "Cancelar Registro???")
            If res = vbYes Then
                Cliente.Recordset.CancelUpdate
            If Not Cliente.Recordset.EOF Then
                Cliente.Recordset.MoveFirst
                End If
            End If
        Case vbKeyF7 'gravar
            res = MsgBox("Deseja confirmar o registro?", _
                48 + vbYesNo, "Gravar registro")
            If res = vbYes Then
               Cliente.Recordset.Update
            If Not Cliente.Recordset.EOF Then
                Cliente.Recordset.MoveFirst
            End If
        End If
        
        Case vbKeyF8 'localizar
            MsgBox "Localizar"
        Case vbKeyF9 'sair
            Unload FrmCliente
    End Select

End Sub

Private Sub Limite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    Limite.Text = Format(Limite, "Currency")
End If
End Sub

Private Sub Limite_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
        Case Is = 8    ' Backspace
        Case 48 To 57  ' Numeros de 0 a 9
        Case 44
        Case 46
        Case Else
        KeyAscii = 0
    End Select

End Sub

Private Sub Nacimento1_Change()
If Len(Nacimento1) = 2 Then
      Nacimento1 = Nacimento1 & "/"
      Nacimento1.SelStart = 4
   End If
   If Len(Nacimento1) = 5 Then
      Nacimento1 = Nacimento1 & "/"
      Nacimento1.SelStart = 7
   End If
   If Len(Nacimento1) = 8 Then
      Nacimento1 = Format(Nacimento1, "dd/mm/yyyy")
   End If

End Sub

Private Sub Nacimento1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Nacimento1_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub Nascimento_Change()
 If Len(Nascimento) = 2 Then
      Nascimento = Nascimento & "/"
      Nascimento.SelStart = 4
   End If
   If Len(Nascimento) = 5 Then
      Nascimento = Nascimento & "/"
      Nascimento.SelStart = 7
   End If
   If Len(Nascimento) = 8 Then
      Nascimento = Format(Nascimento, "dd/mm/yyyy")
   End If
End Sub

Private Sub Nascimento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Nascimento_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub Nome_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub NomeContato_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 'incluir
          Data1.Recordset.AddNew
          Codigo.Text = Data1.Recordset.RecordCount + 1
          Codigo.SetFocus
        
        Case vbKeyF2 'Excluir registro atual
         Dim res
         res = MsgBox("Deseja excluir o registro atual?", _
         48 + vbYesNo, "Excluir???")
         If res = vbYes Then
          MsgBox "Registro deletado com sucesso!", vbInformation
          Data1.Recordset.Delete
        If Not Data1.Recordset.EOF Then
          Data1.Recordset.MoveFirst
      End If
    End If
        Case vbKeyF3 'alterar um registro
            Data1.Recordset.Edit
            Codigo.SetFocus
            
        Case vbKeyF4 'cancelar
            
            res = MsgBox("Deseja realmente cancelar este registro?", _
                48 + vbYesNo, "Cancelar Registro???")
            If res = vbYes Then
                MsgBox "Registro cancelado!", vbInformation
                Data1.Recordset.CancelUpdate
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
                End If
            End If
            
        Case vbKeyF7 'gravar
            
        If Codigo.Text = Empty Then
            MsgBox "Informe o codigo do cliente!", vbInformation
        Exit Sub
        End If
        
        If Nome.Text = Empty Then
            MsgBox "Informe o nome do cliente!", vbInformation
        Exit Sub
        End If
        
        If Nascimento.Text = Empty Then
            MsgBox "Por favor, informe a data de nascimento do cliente!", vbInformation
        Exit Sub
        End If
        
        If Inscricao.Text = Empty Then
            MsgBox "Informe a data de inscrição do cliente!", vbInformation
        Exit Sub
        End If
        
            res = MsgBox("Deseja confirmar o registro?", _
                48 + vbYesNo, "Gravar registro")
            If res = vbYes Then
              Data1.UpdateRecord
              Data1.Recordset.Bookmark = Data1.Recordset.LastModified
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
            End If
        End If
        
        Case vbKeyF8 'localizar
            FrmConsClientes.Show vbModal
        
        Case vbKeyF9 'sair
            Unload FrmCliente
    End Select
End Sub

Private Sub PWeb_Click()
'Captura a imagem atual (PODERIA SER COLCOADO NUM TIMER PARA IMAGEM CONTINUA)
Clipboard.Clear
SendMessage mCapHwnd, WM_CAP_GRAB_FRAME, 0, 0
SendMessage mCapHwnd, WM_CAP_EDIT_COPY, 0, 0
PWeb.Picture = Clipboard.GetData
End Sub

Private Sub Responssavel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub RG_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub RG_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub data1_error(dataerr As Integer, response As Integer)
  'trata erros quando nenhum código esta sendo executado

    Select Case dataerr

         Case 3044 'caminho invalido
            MsgBox "O caminho informado não é valido, verifique !"
            End 'uma maneira muito rude de interromper sua aplicacao, mas...

         Case 3024 'nao achou banco de dados
            MsgBox "O arquivo definido não foi encontrado !"
            End

         Case Else
         'MsgBox "Erro  em  : " & Error$(dataerr) 'para mostrar o erro.
         response = vbDataErrContinue  'ou vbdataErrdisplay

    End Select

End Sub

Private Sub data1_Validate(Action As Integer, Save As Integer)
   If Save Then
      Select Case MsgBox("Deseja Salvar Alterações ?", vbQuestion + vbYesNo)
        
        Case vbYes '6
          If Nome.Text = Empty Then
             MsgBox "O campo nome esta vasio!", , "Gravar "
             Nome.SetFocus
             Save = False 'nao salva os dados
             Action = vbDataActionCancel 'Cancela acao
             Exit Sub
          If Nascimento.Text = Empty Then
             MsgBox "O campo nascimento esta vasio!", , "Gravar "
             Nascimento.SetFocus
             Save = False 'nao salva os dados
             Action = vbDataActionCancel 'Cancela acao
             Exit Sub
          If Inscricao.Text = Empty Then
             MsgBox "O campo inscrição esta vasio!", , "Gravar "
             Inscricao.SetFocus
             Save = False 'nao salva os dados
             Action = vbDataActionCancel 'Cancela acao
             Exit Sub
            End If
'        Case vbNo  '7
          Save = False

  End If

End If
End Select
End If
End Sub

