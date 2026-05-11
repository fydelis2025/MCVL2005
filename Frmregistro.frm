VERSION 5.00
Begin VB.Form Frmregistro 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MCVL 2005 - Registro"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdregistrardepois 
      Caption         =   "Registrar Depois"
      Height          =   375
      Left            =   2760
      MouseIcon       =   "Frmregistro.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1920
      TabIndex        =   8
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   3540
      Left            =   0
      ScaleHeight     =   3540
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdregistrar 
      Caption         =   "Registrar Agora"
      Height          =   375
      Left            =   1200
      MouseIcon       =   "Frmregistro.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Liberação"
      Height          =   195
      Index           =   3
      Left            =   1080
      TabIndex        =   7
      Top             =   2400
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      Height          =   195
      Index           =   2
      Left            =   1080
      TabIndex        =   5
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para registrar"
      Height          =   195
      Index           =   1
      Left            =   2880
      TabIndex        =   4
      Top             =   1200
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Faltam"
      Height          =   195
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frmregistro.frx":0614
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Frmregistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdregistrar_Click()
If Text3.Text = "" Then
   Text3.SetFocus
   Exit Sub
End If
FrmPrincipal.ALock.LiberationKey = Text3.Text
If Not FrmPrincipal.ALock.RegisteredUser Then
  MsgBox "Chave de LIBERAÇÃO INCORRETA", vbOKOnly + vbCritical, "Chave Liberação Incorreta"
   Text3.SetFocus
Else
  MsgBox "REGISTRO EFETUADO COM SUCESSO !", vbExclamation, "Registro OK"
  FrmPrincipal.Label6.Visible = False
  FrmPrincipal.Label6.Caption = "VERSÃO REGISTRADA"
  Unload Me
End If


End Sub

Private Sub cmdregistrardepois_Click()
Unload Frmregistro
End Sub

Private Sub Form_Load()
Dim diasQueFaltaParaRegistrar As Integer
diasQueFaltaParaRegistrar = 0
diasQueFaltaParaRegistrar = 15 - (FrmPrincipal.ALock.UsedDays)
Text1.Text = diasQueFaltaParaRegistrar

If diasQueFaltaParaRegistrar <= 0 Then
   cmdregistrardepois.Enabled = False
End If
Text2.Text = FrmPrincipal.ALock.SoftwareCode


End Sub

