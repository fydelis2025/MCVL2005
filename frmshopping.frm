VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmshopping 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produtos no Shopping"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "frmshopping.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Arquivos de programas\MCVL 2005\LocVideo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Shopping"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2835
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton CmdSalvar 
         Height          =   495
         Left            =   6120
         Picture         =   "frmshopping.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton CmdNovo 
         Height          =   495
         Left            =   6120
         Picture         =   "frmshopping.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   375
      End
      Begin MSMask.MaskEdBox Nome 
         DataField       =   "Nome"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox tipo 
         DataField       =   "Tipo"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Gravadora 
         DataField       =   "Gravadora"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   840
         TabIndex        =   3
         Top             =   960
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox cantores 
         DataField       =   "Cantores"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   4080
         TabIndex        =   4
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox editora 
         DataField       =   "Editora"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox autores 
         DataField       =   "Autores"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   4080
         TabIndex        =   6
         Top             =   1560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox custo 
         DataField       =   "Preco_Custo"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox venda 
         DataField       =   "Preco_Venda"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gravar"
         Height          =   195
         Index           =   6
         Left            =   5760
         TabIndex        =   24
         Top             =   2640
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
         Left            =   5640
         TabIndex        =   23
         Top             =   2160
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
         Left            =   3480
         TabIndex        =   22
         Top             =   2160
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
         Left            =   4200
         TabIndex        =   21
         Top             =   2160
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
         Left            =   4920
         TabIndex        =   20
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incuir"
         Height          =   195
         Index           =   5
         Left            =   3585
         TabIndex        =   19
         Top             =   2640
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excluir"
         Height          =   195
         Index           =   1
         Left            =   4260
         TabIndex        =   18
         Top             =   2640
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alterar"
         Height          =   195
         Index           =   2
         Left            =   4980
         TabIndex        =   17
         Top             =   2640
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gravadora"
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   14
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantores"
         Height          =   195
         Index           =   3
         Left            =   4080
         TabIndex        =   13
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Editora"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Autores"
         Height          =   195
         Index           =   5
         Left            =   4080
         TabIndex        =   11
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço Custo"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço Venda"
         Height          =   195
         Index           =   7
         Left            =   1800
         TabIndex        =   9
         Top             =   1920
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmshopping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub autores_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub cantores_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub codigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub codigo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub CmdNovo_Click()
Data1.Recordset.AddNew
Nome.SetFocus
End Sub

Private Sub CmdSalvar_Click()
If Nome.Text = Empty Then
                MsgBox "O campo nome esta vasio", vbInformation
            Exit Sub
            End If
            
            If Gravadora.Text = Empty Then
                MsgBox "O campo gravadora esta vasio", vbInformation
            Exit Sub
            End If
            
            If editora.Text = Empty Then
                MsgBox "O campo editora esta vasio", vbInformation
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

Private Sub Custo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    custo.Text = Format(custo, "Currency")
End If
End Sub

Private Sub Custo_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub editora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Gravadora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Nome_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub tipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub venda_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    Venda.Text = Format(Venda, "Currency")
If Nome.Text = Empty Then
                MsgBox "O campo nome esta vasio", vbInformation
            Exit Sub
            End If
            
            If Gravadora.Text = Empty Then
                MsgBox "O campo gravadora esta vasio", vbInformation
            Exit Sub
            End If
            
            If editora.Text = Empty Then
                MsgBox "O campo editora esta vasio", vbInformation
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

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
        
        Case vbKeyF1 'incluir
          Data1.Recordset.AddNew
          Nome.SetFocus
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
            Nome.SetFocus
            
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
            
            If Nome.Text = Empty Then
                MsgBox "O campo nome esta vasio", vbInformation
            Exit Sub
            End If
            
            If Gravadora.Text = Empty Then
                MsgBox "O campo gravadora esta vasio", vbInformation
            Exit Sub
            End If
            
            If editora.Text = Empty Then
                MsgBox "O campo editora esta vasio", vbInformation
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
            Unload frmshopping
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
          If Gravadora.Text = Empty Then
             MsgBox "O campo gravadora esta vasio!", , "Gravar "
             Gravadora.SetFocus
             Save = False 'nao salva os dados
             Action = vbDataActionCancel 'Cancela acao
             Exit Sub
          If autores.Text = Empty Then
             MsgBox "O campo autores esta vasio!", , "Gravar "
             autores.SetFocus
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
