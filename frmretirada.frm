VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmretirada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançar Retirada"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmretirada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Arquivos de programas\MCVL 2005\LocVideo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Retirada"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2115
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.CommandButton CmdSalvar 
         Height          =   495
         Left            =   3600
         Picture         =   "frmretirada.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton CmdNovo 
         Height          =   495
         Left            =   3600
         Picture         =   "frmretirada.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   840
         Width           =   375
      End
      Begin MSMask.MaskEdBox data 
         DataField       =   "Data"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox vsaque 
         DataField       =   "Valor_Saque"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
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
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox atendente 
         DataField       =   "CodAtendente"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox nome 
         DataField       =   "Atendente"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Top             =   1560
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data "
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor do Saque"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   8
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Histórico"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Atendente"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmretirada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private conecta As Database
Private rsatendentes As Recordset

Private Sub atendente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub atendente_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub atendente_LostFocus()
atendente.Text = Format(atendente, "000000")
    rsatendentes.Seek "=", atendentes
If rsatendentes.EOF = True Then
    MsgBox "Atendete não cadastrado!", vbInformation
    atendente.SetFocus
Else
    Nome.Text = rsatendentes![Nome]
End If
End Sub

Private Sub CmdNovo_Click()
Data1.Recordset.AddNew
Data.SetFocus
End Sub

Private Sub CmdSalvar_Click()
 If Data.Text = Empty Then
                MsgBox "Informe a data", vbInformation
            Exit Sub
            End If
            
            If atendente.Text = Empty Then
                MsgBox "Informe o codigo do atendente", vbInformation
            Exit Sub
            End If
            
            If historico.Text = Empty Then
                MsgBox "O historico tambem deverá ser informado", vbInformation
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
Set conecta = OpenDatabase("C:\arquivos de programas\mcvl 2005\locvideo.mdb")
Set rsatendentes = conecta.OpenRecordset("Atendentes", dbOpenTable)

rsatendentes.Index = "codigo"
End Sub

Private Sub historico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Nome_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
 If Data.Text = Empty Then
                MsgBox "Informe a data", vbInformation
            Exit Sub
            End If
            
            If atendente.Text = Empty Then
                MsgBox "Informe o codigo do atendente", vbInformation
            Exit Sub
            End If
            
            If historico.Text = Empty Then
                MsgBox "O historico tambem deverá ser informado", vbInformation
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

Private Sub vsaque_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    vsaque.Text = Format(vsaque, "Currency")
End If
End Sub

Private Sub vsaque_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
        
        Case vbKeyF1 'incluir
          Data1.Recordset.AddNew
          Data.SetFocus
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
           Data.SetFocus
            
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
            
            If Data.Text = Empty Then
                MsgBox "Informe a data", vbInformation
            Exit Sub
            End If
            
            If atendente.Text = Empty Then
                MsgBox "Informe o codigo do atendente", vbInformation
            Exit Sub
            End If
            
            If historico.Text = Empty Then
                MsgBox "O historico tambem deverá ser informado", vbInformation
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
            Unload frmretirada
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
          If vsaque.Text = Empty Then
             MsgBox "O campo valor do saque esta vasio!", , "Gravar "
             vsaque.SetFocus
             Save = False 'nao salva os dados
             Action = vbDataActionCancel 'Cancela acao
             Exit Sub
          If historico.Text = Empty Then
             MsgBox "O campo historico esta vasio!", , "Gravar "
             historico.SetFocus
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
