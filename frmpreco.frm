VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmpreco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Preço"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmpreco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Arquivos de programas\MCVL 2005\LocVideo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Preco"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin MSMask.MaskEdBox preco 
         DataField       =   "Preco"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   4560
         TabIndex        =   1
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox descricao 
         DataField       =   "Descricao"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox codigo 
         DataField       =   "Codigo"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   5
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço"
         Height          =   195
         Index           =   2
         Left            =   4560
         TabIndex        =   4
         Top             =   120
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmpreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub codigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub descricao_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub preco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    preco.Text = Format(preco, "Currency")
    
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

Private Sub preco_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

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
          MsgBox "Registro deletado com sucesso", vbInformation
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
                MsgBox "Registro cancelado", vbInformation
                Data1.Recordset.CancelUpdate
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
                End If
            End If
            
        Case vbKeyF7 'gravar
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
            Unload frmpreco
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
          If descricao.Text = Empty Then
             MsgBox "O campo descrição esta vasio!", , "Gravar "
             descricao.SetFocus
             Save = False 'nao salva os dados
             Action = vbDataActionCancel 'Cancela acao
             Exit Sub
          If preco.Text = Empty Then
             MsgBox "O campo preço esta vasio!", , "Gravar "
             preco.SetFocus
             Save = False 'nao salva os dados
             Action = vbDataActionCancel 'Cancela acao
             Exit Sub
            End If
'        Case vbNo  '7
          Save = False
 
  End If
     End Select
     End If
End Sub
