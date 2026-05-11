VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmvendaacos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Venda de Video Ações"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   Icon            =   "frmvendaacos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
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
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Video_Acoes"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2115
      ScaleWidth      =   8115
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton CmdSalvar 
         Height          =   495
         Left            =   6600
         Picture         =   "frmvendaacos.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton CmdNovo 
         Height          =   495
         Left            =   6600
         Picture         =   "frmvendaacos.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   840
         Width           =   375
      End
      Begin MSMask.MaskEdBox Cliente 
         DataField       =   "Codigo"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Nome 
         DataField       =   "Nome"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Lote 
         DataField       =   "Lote"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox VVenda 
         DataField       =   "Valor_Venda"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox VResgate 
         DataField       =   "Valor_Regate"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   3240
         TabIndex        =   5
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Data 
         DataField       =   "Data_Aplicacao"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   4800
         TabIndex        =   6
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Lote1 
         DataField       =   "Lote1"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Inicio 
         DataField       =   "Inicio_Resgate"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox VUnitario 
         DataField       =   "Valor_Unitario"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   3240
         TabIndex        =   9
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Resgate 
         DataField       =   "Resgate"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   4800
         TabIndex        =   10
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   19
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lote"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Venda"
         Height          =   195
         Index           =   3
         Left            =   1680
         TabIndex        =   17
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Resgate"
         Height          =   195
         Index           =   4
         Left            =   3240
         TabIndex        =   16
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Aplicação"
         Height          =   195
         Index           =   5
         Left            =   4800
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio Resgate"
         Height          =   195
         Index           =   6
         Left            =   1680
         TabIndex        =   14
         Top             =   1320
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lote"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Unitário"
         Height          =   195
         Index           =   8
         Left            =   3240
         TabIndex        =   12
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resgate"
         Height          =   195
         Index           =   9
         Left            =   4800
         TabIndex        =   11
         Top             =   1320
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmvendaacos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private conecta As Database
Private rscliente As Recordset

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Cliente_LostFocus()
On Error Resume Next
Cliente.Text = Format(Cliente, "000000")
rscliente.Seek "=", Cliente

If rscliente.EOF = True Then
    MsgBox "Cliente não cadastrado", vbInformation
    Cliente.SetFocus
Else
Nome.Text = rscliente![Nome]
End If
End Sub

Private Sub CmdNovo_Click()
Data1.Recordset.AddNew
Cliente.SetFocus
End Sub

Private Sub CmdSalvar_Click()
If Cliente.Text = Empty Then
        MsgBox "Informe o código do cliente", vbInformation
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

Private Sub Form_Load()
Set conecta = OpenDatabase("c:\arquivos de programas\mcvl 2005\LocVideo.mdb")
Set rscliente = conecta.OpenRecordset("Clientes", dbOpenTable)
rscliente.Index = "PrimaryKey"
End Sub

Private Sub Inicio_Change()
If Len(Inicio) = 2 Then
      Inicio = Inicio & "/"
      Inicio.SelStart = 4
   End If
   If Len(Inicio) = 5 Then
      Inicio = Inicio & "/"
      Inicio.SelStart = 7
   End If
   If Len(Inicio) = 8 Then
      Inicio = Format(Inicio, "dd/mm/yyyy")
   End If
End Sub

Private Sub Inicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Inicio_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub Lote_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Lote1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Nome_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Resgate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
If Cliente.Text = Empty Then
    MsgBox "Informe o código do cliente", vbInformation
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

Private Sub VResgate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    VResgate.Text = Format(VResgate, "Currency")
End If
End Sub

Private Sub VResgate_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub VUnitario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    VUnitario.Text = Format(VUnitario, "Currency")
End If
End Sub

Private Sub VUnitario_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub VVenda_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    VVenda.Text = Format(VVenda, "Currency")
End If
End Sub

Private Sub VVenda_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
        Case vbKeyF1 'incluir
          Data1.Recordset.AddNew
          Cliente.SetFocus
        
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
            Cliente.SetFocus
            
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
      If Cliente.Text = Empty Then
        MsgBox "Informe o código do cliente", vbInformation
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
            Unload frmvendaacos
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
          If Lote.Text = Empty Then
             MsgBox "O campo lote esta vasio!", , "Gravar "
             tipo.SetFocus
             Save = False 'nao salva os dados
             Action = vbDataActionCancel 'Cancela acao
             Exit Sub
          If Data.Text = Empty Then
             MsgBox "O campo data da aplicação esta vasio!", , "Gravar "
             Data.SetFocus
             Save = False 'nao salva os dados
             Action = vbDataActionCancel 'Cancela acao
             Exit Sub
           End If
        'Case vbNo  '7
          Save = False
        End If
    End Select
End If
End Sub
