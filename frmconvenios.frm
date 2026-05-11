VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmconvenios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Convênios"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "frmconvenios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Arquivos de programas\MCVL 2005\LocVideo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Convenio"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   6915
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.CommandButton CmdSalvar 
         Height          =   495
         Left            =   6360
         Picture         =   "frmconvenios.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton CmdNovo 
         Height          =   495
         Left            =   6360
         Picture         =   "frmconvenios.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   375
      End
      Begin MSMask.MaskEdBox codigo 
         DataField       =   "Codigo"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox nome 
         DataField       =   "Convenio"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   480
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox desconto 
         DataField       =   "Desconto"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox vale 
         DataField       =   "Vale_Filme"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox dom 
         DataField       =   "Dom"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   480
         TabIndex        =   5
         Top             =   1560
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox seg 
         DataField       =   "Seg"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   1560
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox ter 
         DataField       =   "Ter"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   2400
         TabIndex        =   7
         Top             =   1560
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox qua 
         DataField       =   "Qua"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   3360
         TabIndex        =   8
         Top             =   1560
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox qui 
         DataField       =   "Qui"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   4200
         TabIndex        =   9
         Top             =   1560
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox sex 
         DataField       =   "Sex"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   5040
         TabIndex        =   10
         Top             =   1560
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox sab 
         DataField       =   "Sab"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   6000
         TabIndex        =   11
         Top             =   1560
         Width           =   375
         _ExtentX        =   661
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
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Convênio"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   21
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto (%)"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vale Filme 24h"
         Height          =   195
         Index           =   3
         Left            =   1440
         TabIndex        =   19
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DOM."
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   18
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SEG."
         Height          =   195
         Index           =   5
         Left            =   960
         TabIndex        =   17
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TER."
         Height          =   195
         Index           =   6
         Left            =   1920
         TabIndex        =   16
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QUA."
         Height          =   195
         Index           =   7
         Left            =   2880
         TabIndex        =   15
         Top             =   1560
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QUI."
         Height          =   195
         Index           =   8
         Left            =   3840
         TabIndex        =   14
         Top             =   1560
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SEX."
         Height          =   195
         Index           =   9
         Left            =   4680
         TabIndex        =   13
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SAB."
         Height          =   195
         Index           =   10
         Left            =   5520
         TabIndex        =   12
         Top             =   1560
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmconvenios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdNovo_Click()
Data1.Recordset.AddNew
Codigo.Text = Data1.Recordset.RecordCount + 1
Codigo.SetFocus
End Sub

Private Sub CmdSalvar_Click()
If Codigo.Text = Empty Then
            MsgBox "Informe o codigo do convenio!", vbInformation
        Exit Sub
        End If
        
        If Nome.Text = Empty Then
            MsgBox "Por favor, informe o nome do convenio!", vbInformation
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

End Sub

Private Sub codigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub desconto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub desconto_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub dom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Nome_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
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
If Codigo.Text = Empty Then
            MsgBox "Informe o codigo do convenio!", vbInformation
        Exit Sub
        End If
        
        If Nome.Text = Empty Then
            MsgBox "Por favor, informe o nome do convenio!", vbInformation
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

Private Sub vale_KeyDown(KeyCode As Integer, Shift As Integer)
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
            MsgBox "Informe o codigo do convenio!", vbInformation
        Exit Sub
        End If
        
        If Nome.Text = Empty Then
            MsgBox "Por favor, informe o nome do convenio!", vbInformation
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
            Unload frmconvenios
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
            End If
        Case vbNo  '7
          Save = False
'      End Select
  
End Select
End If
End Sub


