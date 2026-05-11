VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmdespesas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastrar Despesas"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "frmdespesas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Arquivos de programas\MCVL 2005\LocVideo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Despesas"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2595
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton CmdSalvar 
         Height          =   495
         Left            =   5280
         Picture         =   "frmdespesas.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton CmdNovo 
         Height          =   495
         Left            =   5280
         Picture         =   "frmdespesas.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   600
         Width           =   375
      End
      Begin MSMask.MaskEdBox Documento 
         DataField       =   "Documento"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Parcela 
         DataField       =   "Parcela"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox historico 
         DataField       =   "Historico"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Data 
         DataField       =   "Data"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   2640
         TabIndex        =   4
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox VPagar 
         DataField       =   "Valor_A_Pagar"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   3960
         TabIndex        =   5
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox vencimento 
         DataField       =   "Vencimento"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox situacao 
         DataField       =   "Situacao"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox vpago 
         DataField       =   "Valor_Pago"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   3720
         TabIndex        =   8
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox cheque_dinheiro 
         DataField       =   "Cheque"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox banco 
         DataField       =   "Banco"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox n_cheque 
         DataField       =   "Num_Cheque"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   3240
         TabIndex        =   11
         Top             =   2160
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Documento"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº da Parcela"
         Height          =   195
         Index           =   1
         Left            =   3000
         TabIndex        =   21
         Top             =   120
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Histórico"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         Height          =   195
         Index           =   3
         Left            =   2640
         TabIndex        =   19
         Top             =   720
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor a Pagar"
         Height          =   195
         Index           =   4
         Left            =   3960
         TabIndex        =   18
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Vencimento"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Situação"
         Height          =   195
         Index           =   6
         Left            =   1920
         TabIndex        =   16
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Pago"
         Height          =   195
         Index           =   7
         Left            =   3720
         TabIndex        =   15
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque/Dinheiro"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
         Height          =   195
         Index           =   9
         Left            =   1560
         TabIndex        =   13
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Cheque"
         Height          =   195
         Index           =   10
         Left            =   3240
         TabIndex        =   12
         Top             =   1920
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmdespesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub banco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub cheque_dinheiro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub CmdNovo_Click()
Data1.Recordset.AddNew
Documento.SetFocus
End Sub

Private Sub CmdSalvar_Click()
        If Documento.Text = Empty Then
            MsgBox "Informe o documento", vbInformation
        Exit Sub
        End If
        
        If vencimento.Text = Empty Then
            MsgBox "Informe a data de vencimento!", vbInformation
        Exit Sub
        End If
        
        If n_cheque.Text = Empty Then
            MsgBox "Informe o numero do cheque!", vbInformation
        Exit Sub
        End If
        
            res = MsgBox("Deseja confirmar o registro?", _
                48 + vbYesNo, "Gravar registro")
            If res = vbYes Then
              MsgBox "Registro salvo", vbInformation
              Data1.UpdateRecord
              Data1.Recordset.Bookmark = Data1.Recordset.LastModified
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
            End If
        End If
End Sub

Private Sub Data_Change()
If Len(Data) = 2 Then
      Data = Data & "/"
      Data.SelStart = 4
   End If
   If Len(Data) = 5 Then
      Data = Data & "/"
      Data.SelStart = 7
   End If
   If Len(Data) = 8 Then
      Data = Format(Data, "dd/mm/yyyy")
   End If
End Sub

Private Sub Data_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Data_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub Documento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Documento_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub historico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub n_cheque_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
        If Documento.Text = Empty Then
            MsgBox "Informe o documento", vbInformation
        Exit Sub
        End If
        
        If vencimento.Text = Empty Then
            MsgBox "Informe a data de vencimento!", vbInformation
        Exit Sub
        End If
        
        If n_cheque.Text = Empty Then
            MsgBox "Informe o numero do cheque!", vbInformation
        Exit Sub
        End If
        
            res = MsgBox("Deseja confirmar o registro?", _
                48 + vbYesNo, "Gravar registro")
            If res = vbYes Then
              MsgBox "Registro salvo", vbInformation
              Data1.UpdateRecord
              Data1.Recordset.Bookmark = Data1.Recordset.LastModified
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
            End If
        End If
End If
End Sub

Private Sub n_cheque_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub Parcela_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Parcela_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub situacao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub vencimento_Change()
If Len(vencimento) = 2 Then
      vencimento = vencimento & "/"
      vencimento.SelStart = 4
   End If
   If Len(vencimento) = 5 Then
      vencimento = vencimento & "/"
      vencimento.SelStart = 7
   End If
   If Len(vencimento) = 8 Then
      vencimento = Format(vencimento, "dd/mm/yyyy")
   End If

End Sub

Private Sub vencimento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub vencimento_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub VPagar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    VPagar.Text = Format(VPagar, "Currency")
End If
End Sub

Private Sub VPagar_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub vpago_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    vpago.Text = Format(vpago, "Currency")
End If
End Sub

Private Sub vpago_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
        
        Case vbKeyF1 'incluir
          Data1.Recordset.AddNew
          Documento.SetFocus
        
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
            Documento.SetFocus
            
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
        If Documento.Text = Empty Then
            MsgBox "Informe o documento", vbInformation
        Exit Sub
        End If
        
        If vencimento.Text = Empty Then
            MsgBox "Informe a data de vencimento!", vbInformation
        Exit Sub
        End If
        
        If n_cheque.Text = Empty Then
            MsgBox "Informe o numero do cheque!", vbInformation
        Exit Sub
        End If
        
            res = MsgBox("Deseja confirmar o registro?", _
                48 + vbYesNo, "Gravar registro")
            If res = vbYes Then
              MsgBox "Registro salvo", vbInformation
              Data1.UpdateRecord
              Data1.Recordset.Bookmark = Data1.Recordset.LastModified
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
            End If
        End If
        
        Case vbKeyF9 'sair
            Unload frmdespesas
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
          If Documento.Text = Empty Then
             MsgBox "O campo documento esta vasio!", , "Gravar "
             Documento.SetFocus
             Save = False 'nao salva os dados
             Action = vbDataActionCancel 'Cancela acao
             Exit Sub
          If vencimento.Text = Empty Then
             MsgBox "O campo vencimento esta vasio!", , "Gravar "
             vencimento.SetFocus
             Save = False 'nao salva os dados
             Action = vbDataActionCancel 'Cancela acao
             Exit Sub
          If n_cheque.Text = Empty Then
             MsgBox "O campo número do cheque esta vasio!", , "Gravar "
             n_cheque.SetFocus
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
