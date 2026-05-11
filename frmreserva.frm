VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmreserva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incluir Reserva de Filmes"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   Icon            =   "frmreserva.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Arquivos de programas\MCVL 2005\LocVideo.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Reserva"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3075
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton CmdSalvar 
         Height          =   495
         Left            =   6480
         Picture         =   "frmreserva.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton CmdNovo 
         Height          =   495
         Left            =   6480
         Picture         =   "frmreserva.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   720
         Width           =   375
      End
      Begin MSMask.MaskEdBox Codigo 
         DataField       =   "Codigo"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Cliente 
         DataField       =   "Cliente"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox CodTitulo 
         DataField       =   "CodTitulo"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Titulo 
         DataField       =   "Titulo"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   1080
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Obs 
         DataField       =   "Obs"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Prazo 
         DataField       =   "Prazo"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Hario 
         DataField       =   "Horario"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
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
         Left            =   6240
         TabIndex        =   24
         Top             =   2760
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
         Left            =   6120
         TabIndex        =   23
         Top             =   2280
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
         Left            =   3240
         TabIndex        =   22
         Top             =   2280
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
         Left            =   3960
         TabIndex        =   21
         Top             =   2280
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
         Left            =   4680
         TabIndex        =   20
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F4"
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
         Index           =   3
         Left            =   5400
         TabIndex        =   19
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incuir"
         Height          =   195
         Index           =   5
         Left            =   3345
         TabIndex        =   18
         Top             =   2760
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excluir"
         Height          =   195
         Index           =   1
         Left            =   4020
         TabIndex        =   17
         Top             =   2760
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alterar"
         Height          =   195
         Index           =   2
         Left            =   4740
         TabIndex        =   16
         Top             =   2760
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar"
         Height          =   195
         Index           =   3
         Left            =   5400
         TabIndex        =   15
         Top             =   2760
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código Titulo"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         Height          =   195
         Index           =   3
         Left            =   1560
         TabIndex        =   11
         Top             =   840
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prazo"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   2040
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Horário"
         Height          =   195
         Index           =   5
         Left            =   1680
         TabIndex        =   9
         Top             =   2040
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observações"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmreserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private conecta As Database
Private rscliente As Recordset
Private rstitulo As Recordset

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub CmdNovo_Click()
Data1.Recordset.AddNew
Codigo.SetFocus
End Sub

Private Sub CmdSalvar_Click()
 If Codigo.Text = Empty Then
                MsgBox "Informe o codigo do cliente", vbInformation
            Exit Sub
            End If
            
            If CodTitulo.Text = Empty Then
                MsgBox "Informe o codigo do titulo", vbInformation
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
Codigo.Text = Format(Codigo, "000000")
    rscliente.Seek "=", Codigo
If rscliente.EOF = True Then
    MsgBox "Cliente não cadastrado!", vbInformation
    Codigo.SetFocus
Else
    Cliente.Text = rscliente![Nome]
End If
End Sub

Private Sub CodTitulo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub CodTitulo_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub

Private Sub CodTitulo_LostFocus()
CodTitulo.Text = Format(CodTitulo, "000000")
    rstitulo.Seek "=", CodTitulo
If rstitulo.EOF = True Then
    MsgBox "Titulo não cadastrado!", vbInformation
    CodTitulo.SetFocus
Else
    Titulo.Text = rstitulo![Titulo]
End If
End Sub

Private Sub Form_Load()
Set conecta = OpenDatabase("C:\arquivos de programas\mcvl 2005\locvideo.mdb")
Set rscliente = conecta.OpenRecordset("Clientes", dbOpenTable)
Set rstitulo = conecta.OpenRecordset("Titulo", dbOpenTable)

rscliente.Index = "Primarykey"
rstitulo.Index = "Codigo"
End Sub

Private Sub Hario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
 If Codigo.Text = Empty Then
                MsgBox "Informe o codigo do cliente", vbInformation
            Exit Sub
            End If
            
            If CodTitulo.Text = Empty Then
                MsgBox "Informe o codigo do titulo", vbInformation
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

Private Sub Obs_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Prazo_Change()
If Len(Prazo) = 2 Then
      Prazo = Prazo & "/"
      Prazo.SelStart = 4
   End If
   If Len(Prazo) = 5 Then
      Prazo = Prazo & "/"
      Prazo.SelStart = 7
   End If
   If Len(Prazo) = 8 Then
      Prazo = Format(Prazo, "dd/mm/yyyy")
   End If

End Sub

Private Sub Prazo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub titulo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
        
        Case vbKeyF1 'incluir
          Data1.Recordset.AddNew
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
            MsgBox "Registro cancelado", vbInformation
                Data1.Recordset.CancelUpdate
            If Not Data1.Recordset.EOF Then
                Data1.Recordset.MoveFirst
                End If
            End If
        
        Case vbKeyF7 'gravar
        If Codigo.Text = Empty Then
            MsgBox "Informe o codigo do cliente", vbInformation
        Exit Sub
        End If
            
        If CodTitulo.Text = Empty Then
                MsgBox "Informe o codigo do titulo", vbInformation
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
            Unload frmreserva
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
          If Prazo.Text = Empty Then
             MsgBox "O campo prazo esta vasio!", , "Gravar "
             Prazo.SetFocus
             Save = False 'nao salva os dados
             Action = vbDataActionCancel 'Cancela acao
             Exit Sub
          If Hario.Text = Empty Then
             MsgBox "O campo horário esta vasio!", , "Gravar "
             Hario.SetFocus
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
