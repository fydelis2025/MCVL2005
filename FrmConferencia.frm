VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmConferencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MCVL 2005 - Conferência Caixa"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton CmdExecutar 
      Caption         =   "&Executar"
      Height          =   255
      Left            =   600
      MouseIcon       =   "FrmConferencia.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entre com a data para a verificação"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   120
         Width           =   2535
      End
   End
End
Attribute VB_Name = "FrmConferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancelar_Click()
Unload FrmConferencia
End Sub

Private Sub CmdExecutar_Click()
If IsNumeric(Data.Text) = True Then
   'DataEnvironment1.Command1 CInt(Text1.Text)
   'DataReport1.Show
Else
    MsgBox " Informe um ano válido !!!", vbCritical, "Data Report - Usando Parâmetros"
End If

End Sub

Private Sub Data_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0

End Sub
