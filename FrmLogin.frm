VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MCVL 2005 - Login"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1425
      ScaleWidth      =   4545
      TabIndex        =   4
      Top             =   720
      Width           =   4575
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   720
         MaxLength       =   6
         PasswordChar    =   "§"
         TabIndex        =   6
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   720
         TabIndex        =   5
         Top             =   195
         Width           =   3495
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   240
         Picture         =   "FrmLogin.frx":0CCA
         Top             =   960
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   240
         Picture         =   "FrmLogin.frx":309C
         Top             =   240
         Width           =   240
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Acessar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   4620
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Controle de Acesso"
      Top             =   420
      Visible         =   0   'False
      Width           =   1740
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2880
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   1
            TextSave        =   "05:54"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   1
            TextSave        =   "30/08/05"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Data1.Recordset.FindFirst "Nome = '" & Text1 & "'"
If Data1.Recordset.NoMatch Then
    MsgBox "Usuário não cadastrado !", 16, "Acesso Negado"
    Text1 = ""
    Text1.SetFocus
    Exit Sub
Else
    Data1.Recordset.FindFirst "Senha = '" & Text2 & "'"
    If Data1.Recordset.NoMatch Then
        MsgBox "Senha Inválida !", 16, "Acesso Negado"
        Text1 = ""
        Text2 = ""
        Text1.SetFocus
        Exit Sub
    End If
End If
FrmPrincipal.Label1.Caption = "Usuário logado: " & Data1.Recordset.Fields("Nome")
Unload Me
FrmPrincipal.Show

End Sub

Private Sub Command2_Click()
Unload Me
End Sub



Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\bd.mdb"
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Image3_Click()
Form3.Show
End Sub

Private Sub Text1_GotFocus()
StatusBar1.Panels(1) = "Digite seu login"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Data1.Recordset.FindFirst "Nome = '" & Text1 & "'"
    If Data1.Recordset.NoMatch Then
        MsgBox "Usuário não cadastrado !", 16, "Acesso Negado"
        Text1 = ""
        Text1.SetFocus
    Else
        Text2.SetFocus
    End If
    
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Data1.Recordset.FindFirst "Senha = '" & Text2 & "'"
    If Data1.Recordset.NoMatch Then
        MsgBox "Senha Inválida !", 16, "Acesso Negado"
        Text1 = ""
        Text2 = ""
        Text1.SetFocus
    Else
        Command1_Click
    End If
End If
End Sub


Private Sub Text2_GotFocus()
StatusBar1.Panels(1) = "Digite sua senha de acesso"
End Sub


