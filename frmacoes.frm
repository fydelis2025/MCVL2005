VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmacoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Video Ações"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   Icon            =   "frmacoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Arquivos de programas\MCVL 2005\LocVideo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Acoes"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSMask.MaskEdBox Sab 
      DataField       =   "Sab"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   7080
      TabIndex        =   30
      Top             =   3000
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Sex 
      DataField       =   "Sex"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   6000
      TabIndex        =   29
      Top             =   3000
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Qui 
      DataField       =   "Qui"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   4920
      TabIndex        =   28
      Top             =   3000
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Qua 
      DataField       =   "Qua"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   3960
      TabIndex        =   27
      Top             =   3000
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Ter 
      DataField       =   "Ter"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2880
      TabIndex        =   26
      Top             =   3000
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Seg 
      DataField       =   "Seg"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1800
      TabIndex        =   25
      Top             =   3000
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Dom 
      DataField       =   "Dom"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   720
      TabIndex        =   24
      Top             =   3000
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox ValorLote1 
      DataField       =   "Valor_Lote1"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   4320
      TabIndex        =   23
      Top             =   2040
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox valoracao1 
      DataField       =   "Valor_Acao1"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1920
      TabIndex        =   22
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox valorlote 
      DataField       =   "Valor_Lote"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   4320
      TabIndex        =   21
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox valoracao 
      DataField       =   "Valor_Acao"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1920
      TabIndex        =   20
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox dias 
      DataField       =   "Dias"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   4080
      TabIndex        =   19
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox lote 
      DataField       =   "Lote"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   840
      TabIndex        =   18
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.PictureBox Picture1 
      Height          =   4215
      Left            =   0
      ScaleHeight     =   4155
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton CmdSalvar 
         Height          =   495
         Left            =   6000
         Picture         =   "frmacoes.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton CmdNovo 
         Height          =   495
         Left            =   5520
         Picture         =   "frmacoes.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   3480
         Width           =   375
      End
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   7815
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lote"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dias Mínimo de Aplicações"
            Height          =   195
            Index           =   1
            Left            =   1920
            TabIndex        =   16
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Defina os dias em que este titulo pederá ser usado"
         Height          =   735
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   7815
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DOM."
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SEG."
            Height          =   195
            Index           =   10
            Left            =   1200
            TabIndex        =   13
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TER."
            Height          =   195
            Index           =   11
            Left            =   2280
            TabIndex        =   12
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "QUA."
            Height          =   195
            Index           =   12
            Left            =   3360
            TabIndex        =   11
            Top             =   360
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "QUI."
            Height          =   195
            Index           =   13
            Left            =   4440
            TabIndex        =   10
            Top             =   360
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SEX."
            Height          =   195
            Index           =   14
            Left            =   5400
            TabIndex        =   9
            Top             =   360
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SAB."
            Height          =   195
            Index           =   15
            Left            =   6480
            TabIndex        =   8
            Top             =   360
            Width           =   360
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Valores para resgate pela locadora"
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   7815
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor de Cada Ação"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor de Lote"
            Height          =   195
            Index           =   7
            Left            =   3120
            TabIndex        =   5
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Valores de venda ao cliente"
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   7815
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor de Cada Ação"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   1425
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor de Lote"
            Height          =   195
            Index           =   4
            Left            =   3120
            TabIndex        =   2
            Top             =   360
            Width           =   945
         End
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
         Left            =   3960
         TabIndex        =   42
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fechar"
         Height          =   195
         Index           =   7
         Left            =   4080
         TabIndex        =   41
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gravar"
         Height          =   195
         Index           =   6
         Left            =   3240
         TabIndex        =   40
         Top             =   3960
         Width           =   480
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
         Left            =   3120
         TabIndex        =   39
         Top             =   3480
         Width           =   735
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
         Left            =   240
         TabIndex        =   38
         Top             =   3480
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
         Left            =   960
         TabIndex        =   37
         Top             =   3480
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
         Left            =   1680
         TabIndex        =   36
         Top             =   3480
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
         Left            =   2400
         TabIndex        =   35
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incuir"
         Height          =   195
         Index           =   5
         Left            =   345
         TabIndex        =   34
         Top             =   3960
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excluir"
         Height          =   195
         Index           =   1
         Left            =   1020
         TabIndex        =   33
         Top             =   3960
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alterar"
         Height          =   195
         Index           =   2
         Left            =   1740
         TabIndex        =   32
         Top             =   3960
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar"
         Height          =   195
         Index           =   3
         Left            =   2400
         TabIndex        =   31
         Top             =   3960
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdNovo_Click()
Data1.Recordset.AddNew
Lote.SetFocus
End Sub

Private Sub CmdSalvar_Click()
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

Private Sub dias_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If

End Sub

Private Sub dom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If

End Sub

Private Sub Lote_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If

End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
        Case vbKeyF1 'incluir
          Data1.Recordset.AddNew
          Lote.SetFocus
        
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
            Lote.SetFocus
            
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
            res = MsgBox("Deseja confirmar o registro?", _
                48 + vbYesNo, "Gravar registro")
            If res = vbYes Then
              Data1.UpdateRecord
              Data1.Recordset.Bookmark = Data1.Recordset.LastModified
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
            End If
        End If
        
        Case vbKeyF9 'sair
            Unload frmacoes
    End Select
End Sub

Private Sub qua_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If

End Sub

Private Sub qui_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If

End Sub

Private Sub sab_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
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

Private Sub seg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If

End Sub

Private Sub sex_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If

End Sub

Private Sub ter_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If

End Sub

Private Sub valoracao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    valoracao.Text = Format(valoracao, "Currency")
End If

End Sub

Private Sub valoracao1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    valoracao1.Text = Format(valoracao1, "Currency")
End If

End Sub

Private Sub valorlote_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    valorlote.Text = Format(valorlote, "Currency")
End If

End Sub

Private Sub ValorLote1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    ValorLote1.Text = Format(ValorLote1, "Currency")
End If

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
          If Lote.Text = Empty Then
             MsgBox "O campo lote esta vasio!", , "Gravar "
             Lote.SetFocus
             Save = False 'nao salva os dados
             Action = vbDataActionCancel 'Cancela acao
             Exit Sub
          ElseIf dias.Text = Empty Then
             MsgBox "O campo dias esta vasio!", , "Gravar "
             dias.SetFocus
             Save = False
             Action = vbDataActionCancel
             Exit Sub
         End If
'        Case vbNo  '7
          Save = False

  
End Select
End If
End Sub
