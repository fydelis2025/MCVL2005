VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmrecado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MCVL 2005 - Recado "
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "Frmrecado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Arquivos de programas\MCVL 2005\LocVideo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Recado"
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox recado 
      DataField       =   "Recado"
      DataSource      =   "Data1"
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   960
      Width           =   5055
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1875
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   0
      Width           =   5415
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
      Begin MSMask.MaskEdBox cliente 
         DataField       =   "Cliente"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recado"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   3
         Top             =   120
         Width           =   480
      End
   End
End
Attribute VB_Name = "Frmrecado"
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
    MsgBox "Cliente não cadastrado!", vbInformation
Else
    cliente.Text = rscliente![Nome]
End If
End Sub

Private Sub Form_Load()
Set conecta = OpenDatabase("c:\arquivos de programas\mcvl 2005\locvideo.mdb")
Set rscliente = conecta.OpenRecordset("Clientes", dbOpenTable)

rscliente.Index = "PrimaryKey"
End Sub

Private Sub recado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
If codigo.Text = Empty Then
            MsgBox "Informe o codigo do cliente", vbInformation
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
          codigo.SetFocus
        
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
          codigo.SetFocus
            
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
           If codigo.Text = Empty Then
            MsgBox "Informe o codigo do cliente", vbInformation
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
            Unload Frmrecado
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
          If cliente.Text = Empty Then
             MsgBox "O campo cliente esta vasio!", , "Gravar "
             cliente.SetFocus
             Save = False 'nao salva os dados
             Action = vbDataActionCancel 'Cancela acao
             Exit Sub
          If recado.Text = Empty Then
             MsgBox "O campo recado esta vasio!", , "Gravar "
             recado.SetFocus
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
