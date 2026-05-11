VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmpgto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagamento de Débitos"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   Icon            =   "Frmpgto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Arquivos de programas\MCVL 2005\LocVideo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Pgto_Debitos"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CmdSalvar 
      Height          =   495
      Left            =   5760
      Picture         =   "Frmpgto.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton CmdNovo 
      Height          =   495
      Left            =   5760
      Picture         =   "Frmpgto.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   360
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2115
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin MSMask.MaskEdBox codigo 
         DataField       =   "Codigo"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox nome 
         DataField       =   "Cliente"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox dtpgto 
         DataField       =   "Data_Pgto"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox VTotal 
         DataField       =   "Valor_Total"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Saldos 
         DataField       =   "Uso_Saldos"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   2760
         TabIndex        =   5
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox VApagar 
         DataField       =   "Valor_Pagar"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   4080
         TabIndex        =   6
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Codigo1 
         DataField       =   "CodAtendente"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox atendente 
         DataField       =   "Atendente"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   1560
         Width           =   4215
         _ExtentX        =   7435
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
         TabIndex        =   15
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Cliente"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   14
         Top             =   120
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Pgto"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total"
         Height          =   195
         Index           =   3
         Left            =   1440
         TabIndex        =   12
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Uso Saldos"
         Height          =   195
         Index           =   4
         Left            =   2760
         TabIndex        =   11
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor a Pagar"
         Height          =   195
         Index           =   5
         Left            =   4080
         TabIndex        =   10
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Atendente"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
   End
End
Attribute VB_Name = "Frmpgto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private conecta As Database
Private rscliente As Recordset
Private rsatendentes As Recordset
Private Sub atendente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
            If codigo.Text = Empty Then
                MsgBox "Por favor, informe o codigo do cliente!", vbInformation
            Exit Sub
            End If
            
            If Codigo1.Text = Empty Then
                MsgBox "Informe o codigo do atendente!", vbInformation
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

Private Sub CmdNovo_Click()
Data1.Recordset.AddNew
codigo.SetFocus
End Sub

Private Sub CmdSalvar_Click()
            If codigo.Text = Empty Then
                MsgBox "Por favor, informe o codigo do cliente!", vbInformation
            Exit Sub
            End If
            
            If Codigo1.Text = Empty Then
                MsgBox "Informe o codigo do atendente!", vbInformation
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

Private Sub codigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub codigo_LostFocus()
On Error Resume Next
codigo.Text = Format(codigo, "000000")
    rscliente.Seek "=", codigo
If rscliente.EOF = True Then
    MsgBox "Cliente não cadastrado", vbInformation
    codigo.SetFocus
Else
    Nome.Text = rscliente![Nome]
End If
End Sub

Private Sub codigo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Codigo1_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub codigo1_LostFocus()
Codigo1.Text = Format(Codigo1, "000000")
    rsatendentes.Seek "=", Codigo1
If rsatendentes.EOF = True Then
    MsgBox "Atendente não cadastrado!", vbInformation
    Codigo1.SetFocus
Else
    Nome.Text = rsatendentes![Nome]
End If
End Sub

Private Sub dtpgto_Change()
If Len(dtpgto) = 2 Then
      dtpgto = dtpgto & "/"
      dtpgto.SelStart = 4
   End If
   If Len(dtpgto) = 5 Then
      dtpgto = dtpgto & "/"
      dtpgto.SelStart = 7
   End If
   If Len(dtpgto) = 8 Then
      dtpgto = Format(dtpgto, "dd/mm/yyyy")
   End If

End Sub

Private Sub dtpgto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub dtpgto_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub Form_Load()
Set conecta = OpenDatabase("C:\arquivos de programas\mcvl 2005\locvideo.mdb")
Set rscliente = conecta.OpenRecordset("Clientes", dbOpenTable)
Set rsatendentes = conecta.OpenRecordset("Atendentes", dbOpenTable)

rscliente.Index = "PrimaryKey"
rsatendentes.Index = "codigo"
End Sub

Private Sub Nome_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Saldos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    Saldos.Text = Format(Saldos, "Currency")
End If
End Sub

Private Sub VApagar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    VApagar.Text = Format(VApagar, "Currency")
End If
End Sub

Private Sub VApagar_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub VTotal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    VTotal.Text = Format(VTotal, "Currency")
End If
End Sub

Private Sub VTotal_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
        
        Case vbKeyF1 'incluir
         Data1.Recordset.AddNew
         codigo.SetFocus
        
        Case vbKeyF2 'Excluir registro atual
         Dim res
         res = MsgBox("Deseja excluir o registro atual?", _
         48 + vbYesNo, "Excluir???")
         If res = vbYes Then
         MsgBox "Registro deletado com sucesso!", vbInformation
          Data1.Recordset.Delete
        If Not Adodc1.Recordset.EOF Then
          Data1.Recordset.MoveFirst
      End If
    End If
        Case vbKeyF3 'alterar um registro
            Data1.Recordset.Edit
            codigo.SetFocus
            
        Case vbKeyF4 'cancelar
            
            res = MsgBox("Deseja realmente cancelar este registro?", _
                48 + vbYesNo, "Cancelar Registro???")
            If res = vbYes Then
                MsgBox "Registr cancelado", vbInformation
                Data1.Recordset.CancelUpdate
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
                End If
            End If
        Case vbKeyF7 'gravar
            
            If codigo.Text = Empty Then
                MsgBox "Por favor, informe o codigo do cliente!", vbInformation
            Exit Sub
            End If
            
            If Codigo1.Text = Empty Then
                MsgBox "Informe o codigo do atendente!", vbInformation
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
            Unload Frmpgto
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
          If Nome.Text = Empty Then
             MsgBox "O campo nome esta vasio!", , "Gravar "
             Nome.SetFocus
             Save = False 'nao salva os dados
             Action = vbDataActionCancel 'Cancela acao
             Exit Sub
          If dtpgto.Text = Empty Then
             MsgBox "O campo data do pagamento esta vasio!", , "Gravar "
             dtpgto.SetFocus
             Save = False 'nao salva os dados
             Action = vbDataActionCancel 'Cancela acao
             Exit Sub
          If atendente.Text = Empty Then
             MsgBox "O campo atendente esta vasio!", , "Gravar "
             atendente.SetFocus
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
