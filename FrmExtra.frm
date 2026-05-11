VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmExtra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MCVL 2005 - Controle do Caixa"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Arquivos de programas\MCVL 2005\LocVideo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Extra"
      Top             =   840
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSMask.MaskEdBox Atendente 
      DataField       =   "Atendente"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      Top             =   1680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox codigo 
      DataField       =   "Codigo"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Historico 
      DataField       =   "Historico"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox extra 
      DataField       =   "Entrada"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   1440
      TabIndex        =   7
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Data 
      DataField       =   "Data"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      PromptChar      =   "_"
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Atendente"
         Height          =   195
         Index           =   4
         Left            =   1320
         TabIndex        =   5
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Histórico"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entrada Extra"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   345
      End
   End
End
Attribute VB_Name = "FrmExtra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private conecta As Database
Private rsatendente As Recordset
Private Sub Atendente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    If Data.Text = Empty Then
            MsgBox "Informe a data", vbInformation
        Exit Sub
        End If
        
        If codigo.Text = Empty Then
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

Private Sub codigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub codigo_LostFocus()
codigo.Text = Format(codigo, "000000")
    rsatendente.Seek "=", codigo
If rsatendente.EOF = True Then
    MsgBox "Atendente não cadastrado!", vbInformation
    codigo.SetFocus
Else
    Atendente.Text = rsatendente![Nome]
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

Private Sub extra_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub extra_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub Form_Load()
Set conecta = OpenDatabase("c:\arquivos de programas\mcvl 2005\locvideo.mdb")
Set rsatendente = conecta.OpenRecordset("Atendentes", dbOpenTable)

rsatendente.Index = "Codigo"
End Sub

Private Sub Historico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
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
          MsgBox "Registro deletado com sucesso!", vbInformation
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
                MsgBox "Registro cancelado!", vbInformation
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
        
        If codigo.Text = Empty Then
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
            Unload FrmExtra
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

