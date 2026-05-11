VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmtitulos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Titulos"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "frmtitulos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Arquivos de programas\MCVL 2005\LocVideo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Titulo"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3555
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton CmdSalvar 
         Height          =   495
         Left            =   6360
         Picture         =   "frmtitulos.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton CmdNovo 
         Height          =   495
         Left            =   6360
         Picture         =   "frmtitulos.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1200
         Width           =   375
      End
      Begin MSMask.MaskEdBox genero 
         DataField       =   "Codigo"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox titulo 
         DataField       =   "Titulo"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox tipo 
         DataField       =   "Tipo"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox audio 
         DataField       =   "Audio"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Legendas 
         DataField       =   "Legendas"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   2640
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox copias 
         DataField       =   "N_Copias"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   4200
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Automatico 
         DataField       =   "Automatico"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox sequencial 
         DataField       =   "Sequencial"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1560
         TabIndex        =   14
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox barras 
         DataField       =   "Codigo_Barras"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   3120
         TabIndex        =   15
         Top             =   1800
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox classificacao 
         DataField       =   "Classificacao"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox preco 
         DataField       =   "Categoria"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1680
         TabIndex        =   17
         Top             =   2400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox compra 
         DataField       =   "Data_Compra"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   3240
         TabIndex        =   18
         Top             =   2400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox fornecedor 
         DataField       =   "Fornecedor"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Custo 
         DataField       =   "Preco"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1680
         TabIndex        =   20
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox venda 
         DataField       =   "Venda"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   3240
         TabIndex        =   21
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox embalagem 
         DataField       =   "Embalagem"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   4800
         TabIndex        =   22
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Automatico"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   32
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sequencial"
         Height          =   195
         Index           =   7
         Left            =   1560
         TabIndex        =   31
         Top             =   1560
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código de Barras"
         Height          =   195
         Index           =   8
         Left            =   3120
         TabIndex        =   30
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Classificação"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   29
         Top             =   2160
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria de preço"
         Height          =   195
         Index           =   10
         Left            =   1680
         TabIndex        =   28
         Top             =   2160
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Compra"
         Height          =   195
         Index           =   11
         Left            =   3240
         TabIndex        =   27
         Top             =   2160
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fornecedor"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   26
         Top             =   2760
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço Custo"
         Height          =   195
         Index           =   13
         Left            =   1680
         TabIndex        =   25
         Top             =   2760
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venda"
         Height          =   195
         Index           =   14
         Left            =   3240
         TabIndex        =   24
         Top             =   2760
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Embalagem"
         Height          =   195
         Index           =   15
         Left            =   4800
         TabIndex        =   23
         Top             =   2760
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Título"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Audio"
         Height          =   195
         Index           =   3
         Left            =   1320
         TabIndex        =   9
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Legendas"
         Height          =   195
         Index           =   4
         Left            =   2640
         TabIndex        =   8
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Cópias"
         Height          =   195
         Index           =   5
         Left            =   4200
         TabIndex        =   7
         Top             =   840
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmtitulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub audio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Automatico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub barras_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub classificacao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub CmdNovo_Click()
Data1.Recordset.AddNew
genero.Text = Data1.Recordset.RecordCount + 1
genero.SetFocus
End Sub

Private Sub CmdSalvar_Click()
            If genero.Text = Empty Then
                MsgBox "Informe o codigo", vbInformation
            Exit Sub
            End If
            
            If Titulo.Text = Empty Then
                MsgBox "O nome do titulo tambem deverá ser preenchido", vbInformation
            Exit Sub
            End If
            
            If fornecedor.Text = Empty Then
                MsgBox "Informe o fornecedor", vbInformation
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

Private Sub compra_Change()
If Len(compra) = 2 Then
      compra = compra & "/"
      compra.SelStart = 4
   End If
   If Len(compra) = 5 Then
      compra = compra & "/"
      compra.SelStart = 7
   End If
   If Len(compra) = 8 Then
      compra = Format(compra, "dd/mm/yyyy")
   End If


End Sub

Private Sub compra_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub compra_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub copias_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub copias_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub Custo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    custo.Text = Format(custo, "Currency")
End If
End Sub

Private Sub Custo_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub embalagem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
                
            If genero.Text = Empty Then
                MsgBox "Informe o codigo", vbInformation
            Exit Sub
            End If
            
            If Titulo.Text = Empty Then
                MsgBox "O nome do titulo tambem deverá ser preenchido", vbInformation
            Exit Sub
            End If
            
            If fornecedor.Text = Empty Then
                MsgBox "Informe o fornecedor", vbInformation
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

Private Sub fornecedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub genero_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub genero_LostFocus()
genero.Text = Format(genero, "000000")
End Sub

Private Sub Legendas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub preco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    preco.Text = Format(preco, "Currency")
End If
End Sub

Private Sub sequencial_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub tipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub titulo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub venda_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    Venda.Text = Format(Venda, "Currency")
End If
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
        
        Case vbKeyF1 'incluir
          Data1.Recordset.AddNew
          genero.Text = Data1.Recordset.RecordCount + 1
          genero.SetFocus
        Case vbKeyF2 'Excluir registro atual
         Dim res
         res = MsgBox("Deseja excluir o registro atual?", _
         48 + vbYesNo, "Excluir???")
         If res = vbYes Then
          MsgBox "Registro deletado com sucesso", vbInformation
          Data1.Recordset.Delete
        If Not Data1.Recordset.EOF Then
          Data1.Recordset.MoveFirst
      End If
    End If
        Case vbKeyF3 'alterar um registro
          Data1.Recordset.Edit
            genero.SetFocus
            
        Case vbKeyF4 'cancelar
            
            res = MsgBox("Deseja realmente cancelar este registro?", _
                48 + vbYesNo, "Cancelar Registro???")
            If res = vbYes Then
                MsgBox "Registro cancelado", vbInformation
                Data1.Recordset.CancelUpdate
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
                End If
            End If
            
        Case vbKeyF7 'gravar
            
            If genero.Text = Empty Then
                MsgBox "Informe o genero", vbInformation
            Exit Sub
            End If
            
            If Titulo.Text = Empty Then
                MsgBox "O nome do titulo tambem deverá ser preenchido", vbInformation
            Exit Sub
            End If
            
            If fornecedor.Text = Empty Then
                MsgBox "Informe o fornecedor", vbInformation
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
        Case vbKeyF9 'sair
            Unload frmtitulos
    End Select
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
          If tipo.Text = Empty Then
             MsgBox "O campo tipo esta vasio!", , "Gravar "
             tipo.SetFocus
             Save = False 'nao salva os dados
             Action = vbDataActionCancel 'Cancela acao
             Exit Sub
             End If
          If audio.Text = Empty Then
             MsgBox "O campo audio esta vasio!", , "Gravar "
             audio.SetFocus
             Save = False 'nao salva os dados
             Action = vbDataActionCancel 'Cancela acao
             Exit Sub
             End If
          If compra.Text = Empty Then
             MsgBox "O campo data da compra esta vasio!", , "Gravar "
             compra.SetFocus
             Save = False 'nao salva os dados
             Action = vbDataActionCancel 'Cancela acao
             Exit Sub
            End If
        Case vbNo  '7
          Save = False
        End Select
        End If
End Sub
