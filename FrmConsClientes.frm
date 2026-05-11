VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmConsClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MSCVL 2005 - Consultar Clientes"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "FrmConsClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   255
      Left            =   5640
      MouseIcon       =   "FrmConsClientes.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton CmdProcurar 
      Caption         =   "&Procurar"
      Height          =   255
      Left            =   5640
      MouseIcon       =   "FrmConsClientes.frx":0FD4
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3600
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4395
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   3840
         Width           =   3375
      End
      Begin VB.Frame Frame1 
         Caption         =   "Procurar Por:"
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   3360
         Width           =   1695
         Begin VB.OptionButton Option3 
            Caption         =   "CPF"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "RG"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Nome"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5530
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   3
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Digite aqui o que deseja procurar!"
         Height          =   195
         Left            =   1920
         TabIndex        =   6
         Top             =   3600
         Width           =   2385
      End
   End
End
Attribute VB_Name = "FrmConsClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cnn As New ADODB.Connection
Public rs As New ADODB.Recordset

Private Sub CmdCancelar_Click()
Unload FrmConsClientes
rs.Close
End Sub

Private Sub Form_Load()
PreencherDataGrid
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
rs.Close
Set rs = Nothing
  cnn.Open "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=c:\Arquivos de programas\mcvl 2005\locvideo.mdb;"
    rs.CursorLocation = adUseClient
    rs.Open "Select * From clientes", cnn, adOpenKeyset, adLockOptimistic, adCmdText
End Sub



Private Sub button1_Click()


End Sub

Private Sub cmdprocurar_Click()

If Option1.Value = False And Option2.Value = False And Option3.Value = False Then
    busca = MsgBox("Selecione uma opção ao lado", vbInformation, "Consultar Clientes")
End If

'Buscar Códigos
If Option1.Value = True Then
    If CmdProcurar.Caption = "&Voltar" Then
        rs.Close
        rs.Open "select * from clientes ", cnn
        CmdProcurar.Caption = "&Procurar"
    Else
        Resposta = Text1.Text
        rs.Close
        rs.Open "select * from clientes where nome like '%" & Resposta & "%'", cnn
        If rs.EOF Then
            MsgBox "Não encontrei o cliente: " & Resposta
            rs.Close
            rs.Open "select * from clientes ", cnn
        Else
            CmdProcurar.Caption = "&Voltar"
            Text1.Text = ""
        End If
End If
PreencherDataGrid
End If

'Buscar Registros
If Option2.Value = True Then
    If CmdProcurar.Caption = "&Voltar" Then
        rs.Close
        rs.Open "select * from clientes ", cnn
        CmdProcurar.Caption = "&Procurar"
    Else
        Resposta = Text1.Text
        rs.Close
        rs.Open "select * from clientes where RG like '%" & Resposta & "%'", cnn
        If rs.EOF Then
            MsgBox "Não encontrei o RG: " & Resposta
            rs.Close
            rs.Open "select * from clientes ", cnn
        Else
            CmdProcurar.Caption = "&Voltar"
            Text1.Text = ""
        End If
End If
PreencherDataGrid
End If
'Buscar Nomes
If Option3.Value = True Then
    If CmdProcurar.Caption = "&Voltar" Then
        rs.Close
        rs.Open "select * from clientes ", cnn
        cmdBuscar.Caption = "&Buscar"
    Else
        Resposta = Text1.Text
        rs.Close
        rs.Open "select * from clientes where CPF like '%" & Resposta & "%'", cnn
        If rs.EOF Then
            MsgBox "Não encontrei o CPF: " & Resposta
            rs.Close
            rs.Open "select * from clientes ", cnn
        Else
            CmdProcurar.Caption = "&Voltar"
            Text1.Text = ""
        End If
End If
PreencherDataGrid
End If

End Sub


Private Sub cmdVoltar_Click()
rs.MovePrevious
If rs.BOF Then
    rs.MoveFirst
End If
PreencherDataGrid
End Sub


Sub PreencherDataGrid()

DataGrid1.Columns.Add (0)

Set DataGrid1.DataSource = rs
DataGrid1.Refresh
DataGrid1.Columns(0).Width = 752
DataGrid1.Columns(1).Width = 752
DataGrid1.Columns(2).Width = 3156

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'Executar Busca Quando Teclar ENTER no Final da Frase que Deseja Buscar

If KeyAscii = 13 Then

If Option1.Value = False And Option2.Value = False And Option3.Value = False Then
    busca = MsgBox("Selecione uma opção ao lado", vbInformation, "Consultat Clientes")
End If

If Option1.Value = True Then
    If CmdProcurar.Caption = "&Voltar" Then
        rs.Close
        rs.Open "select * from clientes ", cnn
        CmdProcurar.Caption = "&Buscar"
    Else
        Resposta = Text1.Text
        rs.Close
        rs.Open "select * from clientes where nome like '%" & Resposta & "%'", cnn
        If rs.EOF Then
            MsgBox "Não encontrei o cliente: " & Resposta
            rs.Close
            rs.Open "select * from clientes ", cnn
        Else
            cmdBuscar.Caption = "&Voltar"
            Text1.Text = ""
        End If
End If
PreencherDataGrid
End If

If Option2.Value = True Then
    If CmdProcurar.Caption = "&Voltar" Then
        rs.Close
        rs.Open "select * from clientes ", cnn
        CmdProcurar.Caption = "&Buscar"
    Else
        Resposta = Text1.Text
        rs.Close
        rs.Open "select * from clientes where RG like '%" & Resposta & "%'", cnn
        If rs.EOF Then
            MsgBox "Não encontrei o RG: " & Resposta
            rs.Close
            rs.Open "select * from clientes ", cnn
        Else
            CmdProcurar.Caption = "&Voltar"
            Text1.Text = ""
        End If
End If
PreencherDataGrid
End If

If Option3.Value = True Then
    If CmdProcurar.Caption = "&Voltar" Then
        rs.Close
        rs.Open "select * from clientes", cnn
        CmdProcurar.Caption = "&Buscar"
    Else
        Resposta = Text1.Text
        rs.Close
        rs.Open "select * from clientes where CPF like '%" & Resposta & "%'", cnn
        If rs.EOF Then
            MsgBox "Não encontrei o CPF: " & Resposta
            rs.Close
            rs.Open "select * from clientes ", cnn
        Else
            CmdProcurar.Caption = "&Voltar"
            Text1.Text = ""
        End If
End If
PreencherDataGrid
End If
End If
End Sub






