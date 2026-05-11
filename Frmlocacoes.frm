VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmlocacoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MCVL 2005 - Locações"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "Frmlocacoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Arquivos de programas\MCVL 2005\LocVideo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Caixa"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox Picture1 
      Height          =   7095
      Left            =   0
      ScaleHeight     =   7035
      ScaleWidth      =   11835
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.Frame Frame1 
         Height          =   4095
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   6135
         Begin MSMask.MaskEdBox txtcodcliente 
            DataField       =   "Codigo"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   240
            TabIndex        =   22
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtcliente 
            DataField       =   "Cliente"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   1920
            TabIndex        =   23
            Top             =   600
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtcodfilme 
            DataField       =   "CodFilme"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   840
            TabIndex        =   24
            Top             =   1080
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txttipo 
            DataField       =   "Tipo"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   840
            TabIndex        =   25
            Top             =   1560
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtfilme 
            DataField       =   "Filme"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   840
            TabIndex        =   26
            Top             =   2040
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtpreco 
            DataField       =   "Valor"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   840
            TabIndex        =   27
            Top             =   2520
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código "
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   47
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente"
            Height          =   195
            Index           =   1
            Left            =   1920
            TabIndex        =   46
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código "
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   45
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   44
            Top             =   1560
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filme"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   43
            Top             =   2040
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   42
            Top             =   2520
            Width           =   360
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Novo"
            Height          =   195
            Index           =   4
            Left            =   3600
            TabIndex        =   41
            Top             =   3840
            Width           =   390
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Consultar"
            Height          =   195
            Index           =   3
            Left            =   2760
            TabIndex        =   40
            Top             =   3840
            Width           =   660
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Forma"
            Height          =   195
            Index           =   2
            Left            =   2100
            TabIndex        =   39
            Top             =   3840
            Width           =   435
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Excluir"
            Height          =   195
            Index           =   1
            Left            =   1380
            TabIndex        =   38
            Top             =   3840
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Iniciar"
            Height          =   195
            Index           =   5
            Left            =   700
            TabIndex        =   37
            Top             =   3840
            Width           =   420
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
            Index           =   4
            Left            =   3480
            TabIndex        =   36
            Top             =   3360
            Width           =   615
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
            Left            =   2760
            TabIndex        =   35
            Top             =   3360
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
            Left            =   2040
            TabIndex        =   34
            Top             =   3360
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
            Left            =   1320
            TabIndex        =   33
            Top             =   3360
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
            Left            =   600
            TabIndex        =   32
            Top             =   3360
            Width           =   615
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
            Index           =   7
            Left            =   4200
            TabIndex        =   31
            Top             =   3360
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gravar"
            Height          =   195
            Index           =   6
            Left            =   4320
            TabIndex        =   30
            Top             =   3840
            Width           =   480
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cancelar"
            Height          =   195
            Index           =   7
            Left            =   5040
            TabIndex        =   29
            Top             =   3840
            Width           =   630
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
            Left            =   5040
            TabIndex        =   28
            Top             =   3360
            Width           =   735
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   6120
            Y1              =   3120
            Y2              =   3120
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2415
         Left            =   0
         TabIndex        =   14
         Top             =   4200
         Width           =   6135
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sejam bem vindos "
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1320
            TabIndex        =   20
            Top             =   240
            Width           =   2760
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hoje"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   3240
            TabIndex        =   19
            Top             =   1080
            Width           =   510
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hora"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   3240
            TabIndex        =   18
            Top             =   1560
            Width           =   525
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   4080
            TabIndex        =   17
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   4080
            TabIndex        =   16
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fechado"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   480
            TabIndex        =   15
            Top             =   1200
            Width           =   1395
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3975
         Left            =   6240
         TabIndex        =   12
         Top             =   0
         Width           =   5535
         Begin MSComctlLib.ListView lvprod 
            Height          =   3645
            Left            =   100
            TabIndex        =   13
            ToolTipText     =   "double click to remove  product in the list"
            Top             =   240
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   6429
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Codigo"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Cliente"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Filme"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Preço"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Valor Total"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Pagar"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Valor Antecipado"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Atendente"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Resumo da Locação"
         Height          =   2535
         Left            =   6240
         TabIndex        =   1
         Top             =   4080
         Width           =   5535
         Begin MSMask.MaskEdBox txtvallocacao 
            DataField       =   "Valor_Total"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   240
            TabIndex        =   2
            Top             =   840
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtpagar 
            DataField       =   "Pagar"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   240
            TabIndex        =   3
            Top             =   1440
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtvalantec 
            DataField       =   "Valor_Atencipado"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   2160
            TabIndex        =   4
            Top             =   1440
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtcodatend 
            DataField       =   "CodAtendente"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   240
            TabIndex        =   5
            Top             =   2040
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtatendente 
            DataField       =   "Atendente"
            DataSource      =   "Data1"
            Height          =   315
            Left            =   2160
            TabIndex        =   6
            Top             =   2040
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor da Locação"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   11
            Top             =   600
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pagar"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   10
            Top             =   1200
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor Antecipado"
            Height          =   195
            Index           =   8
            Left            =   2160
            TabIndex        =   9
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cód. Atendente"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   8
            Top             =   1800
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Atendente"
            Height          =   195
            Index           =   10
            Left            =   2160
            TabIndex        =   7
            Top             =   1800
            Width           =   735
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3240
      Top             =   1320
   End
End
Attribute VB_Name = "Frmlocacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private conecta As Database
Private rscliente As Recordset
Private rsatendentes As Recordset
Private rstitulo As Recordset
Dim lst As ListItem

Private Sub Form_Load()
Set conecta = OpenDatabase("c:\arquivos de programas\MCVL 2005\locvideo.mdb")
Set rscliente = conecta.OpenRecordset("clientes", dbOpenTable)
Set rsatendentes = conecta.OpenRecordset("Atendentes", dbOpenTable)
Set rstitulo = conecta.OpenRecordset("Titulo", dbOpenTable)

rscliente.Index = "Primarykey"
rsatendentes.Index = "codigo"
rstitulo.Index = "codigo"

End Sub


Private Sub lvprod_DblClick()
If lvprod.ListItems.Count = 0 Then
    MsgBox "Sem itens para remover!", vbOKOnly + vbInformation, "Remover"
Else
    If MsgBox("Confirma exclusão " & Chr(10) & Chr(10) & StrConv(lvprod.SelectedItem.SubItems(2), vbUpperCase), vbYesNo + vbQuestion, "Remover Item") = vbYes Then
              
              lvprod.ListItems.Remove lvprod.SelectedItem.Index
                'txtAmount.Text = CStr(Format(ComputeAmount, "########0.00"))

                            
    Else
            Exit Sub
    End If
End If

End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
        
        Case vbKeyF1 'incluir
          Data1.Recordset.AddNew
          txtcodcliente.SetFocus
          Label5(0).Caption = "Caixa Aberto"
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
            txtcodcliente.SetFocus
            
        Case vbKeyF4 'cancelar
            
            res = MsgBox("Deseja realmente cancelar este registro?", _
                48 + vbYesNo, "Cancelar Registro???")
            If res = vbYes Then
                MsgBox "Registro cancelado!", vbInformation
                Label5(0).Caption = "Caixa Cancelado"
                Data1.Recordset.CancelUpdate
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
                End If
            End If
            
        Case vbKeyF7 'gravar
                If txtcodcliente.Text = Empty Then
                MsgBox "Por favor, informe o codigo do cliente!", vbInformation
            Exit Sub
            End If
            
            If txtcodfilme.Text = Empty Then
                MsgBox "O código do filme deve ser informado!", vbInformation
            Exit Sub
            End If
            
            If txtcodatend.Text = Empty Then
                MsgBox "Informe o código do atendente!", vbInformation
            Exit Sub
            End If
            
            res = MsgBox("Deseja confirmar o registro?", _
                48 + vbYesNo, "Gravar registro")
            If res = vbYes Then
              MsgBox "Registro salvo!", vbInformation
              Data1.UpdateRecord
              Data1.Recordset.Bookmark = Data1.Recordset.LastModified
            If Not Data1.Recordset.EOF Then
              Data1.Recordset.MoveFirst
            End If
        End If

        Case vbKeyF9 'sair
            Unload Frmlocacoes
    End Select
End Sub


Private Sub Timer1_Timer()
Label4(0).Caption = CVDate(Date)
Label4(1).Caption = CStr(Time)
End Sub

Private Sub txtatendente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    If txtcodcliente.Text = Empty Then
                MsgBox "Por favor, informe o codigo do cliente!", vbInformation
            Exit Sub
            End If
            
            If txtcodfilme.Text = Empty Then
                MsgBox "O código do filme deve ser informado!", vbInformation
            Exit Sub
            End If
            
            If txtcodatend.Text = Empty Then
                MsgBox "Informe o código do atendente!", vbInformation
            Exit Sub
            End If
            
            res = MsgBox("Deseja confirmar o registro?", _
                48 + vbYesNo, "Gravar registro")
            If res = vbYes Then
               MsgBox "Registro salvo!", vbInformation
              Data1.UpdateRecord
              Data1.Recordset.Bookmark = Data1.Recordset.LastModified
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
            End If
        End If
orderItem
End If
End Sub

Private Sub txtcliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtcodatend_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
 Case vbKeyReturn
    SendKeys "{TAB}"
 Case vbKeyF11
    frmatendentes.Show vbModal
End Select
End Sub

Private Sub txtcodatend_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub txtcodatend_LostFocus()
On Error Resume Next
txtcodatend.Text = Format(txtcodatend, "000000")
    rsatendentes.Seek "=", txtcodatend
If rsatendentes.EOF = True Then
    MsgBox "Atendente não cadastrado!", vbInformation
    txtcodatend.SetFocus
Else
    txtatendente.Text = rsatendentes![Nome]
End If
End Sub

Private Sub txtcodcliente_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
 Case vbKeyReturn
    SendKeys "{TAB}"
 Case vbKeyF10
    FrmCliente.Show vbModal
End Select
End Sub

Private Sub txtcodcliente_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub txtcodcliente_LostFocus()
On Error Resume Next
txtcodcliente.Text = Format(txtcodcliente, "000000")
    rscliente.Seek "=", txtcodcliente
If rscliente.EOF = True Then
    MsgBox "Cliente não cadastrado!", vbInformation
    txtcodcliente.SetFocus
Else
    txtcliente.Text = rscliente![Nome]
End If
End Sub

Private Sub txtcodfilme_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode

 Case vbKeyReturn
    SendKeys "{TAB}"
    
 Case vbKeyF6
    frmtitulos.Show vbModal
    
End Select
End Sub

Private Sub txtcodfilme_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub txtcodfilme_LostFocus()
On Error Resume Next
txtcodfilme.Text = Format(txtcodfilme, "000000")
    rstitulo.Seek "=", txtcodfilme
If rstitulo.EOF = True Then
    MsgBox "Filme não cadastrado!", vbInformation
    txtcodfilme.SetFocus
Else
    txttipo.Text = rstitulo![tipo]
    txtfilme.Text = rstitulo![Titulo]
    txtpreco.Text = rstitulo![preco]
End If
End Sub

Private Sub txtfilme_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtpagar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtpreco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    txtpreco.Text = Format(txtpreco, "Currency")
End If
End Sub

Private Sub txtpreco_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub txttipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtvalantec_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub txtvalantec_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub txtvallocacao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    txtvallocacao.Text = Format(txtvallocacao, "Currency")
End If
End Sub

Private Sub txtvallocacao_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub


Private Sub orderItem()
                 
             'Qty = Val(txtQtyOrder.Text)
             
                 Set lst = lvprod.FindItem(txtcodcliente.Text, , , lvwPartial)
                    If lst Is Nothing Then
                        Set lst = lvprod.ListItems.Add(, , txtcodcliente.Text, , 0)
                         With lst
                                .SubItems(1) = txtcliente.Text
                                .SubItems(2) = txtfilme.Text
                                .SubItems(3) = txtpreco.Text
                                .SubItems(4) = txtvallocacao.Text
                                .SubItems(5) = txtpagar.Text
                                .SubItems(6) = txtvalantec.Text
                                .SubItems(7) = txtatendente.Text
                                                                
                          End With
                    Else
                            
                               With lst
                               
                                .SubItems(3) = txtpreco + .SubItems(4)
                                '.SubItems(4) = Format(.SubItems(1) * txtPrice.Text, "#,###.#0")
                               ' txtAmount.Text = CStr(Format(ComputeAmount, "########0.00"))
                               ' txtDiscount.Text = "0.00"
                               
                               
                               End With
                    End If

End Sub

