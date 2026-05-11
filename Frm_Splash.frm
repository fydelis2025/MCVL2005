VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Frm_Splash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5565
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picpgb2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   19
      ScaleMode       =   0  'User
      ScaleWidth      =   73
      TabIndex        =   7
      Top             =   4320
      Width           =   1100
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4440
      Top             =   2280
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   0
      ScaleHeight     =   4575
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   0
      Width           =   1110
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregando"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   6
         Top             =   4080
         Width           =   990
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   360
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image imgpgb1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Picture         =   "Frm_Splash.frx":0000
      Top             =   3000
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Image3 
      Height          =   1230
      Left            =   1320
      Picture         =   "Frm_Splash.frx":03A6
      Top             =   720
      Width           =   3780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   5
      Top             =   2760
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   4
      Top             =   2760
      Width           =   585
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm_Splash.frx":F610
      Height          =   1215
      Left            =   1200
      TabIndex        =   3
      Top             =   3240
      Width           =   4215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Este programa está Licenciado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Top             =   2400
      Width           =   2640
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Controle de Video Locadora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   3960
   End
End
Attribute VB_Name = "Frm_Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim distance As Integer


Private Sub Form_Initialize()
Label4.Caption = Winsock1.LocalHostName
Label5.Caption = Winsock1.LocalIP
End Sub

Private Sub Form_Load()
    distance = 4
    'Horizontal Me, RGB(131, 166, 244), RGB(33, 120, 224)
    picpgb2.PaintPicture imgpgb1, 0, 0, 4, 19, 0, 0, 4, 19
    picpgb2.PaintPicture imgpgb1, 4, 0, picpgb2.Width - 9, 19, 4, 0, 10, 19
    picpgb2.PaintPicture imgpgb1, picpgb2.Width - 5, 0, 5, 19, 14, 0, 5, 19
End Sub

Private Sub Form_Terminate()
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    For i = 1 To 2
        picpgb2.PaintPicture imgpgb1.Picture, distance, 4, 8, 12, 23, 5, 8, 12
        distance = distance + 10
    Next i
    If distance > picpgb2.Width - 5 Then
        Timer1.Enabled = False
        Unload Me
        Load FrmLogin
        FrmLogin.Show
    End If
End Sub

