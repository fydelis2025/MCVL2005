VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "FLASH.OCX"
Begin VB.Form FrmSobre 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MCVL 2005 - Sobre "
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "FrmSobre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   735
      Left            =   840
      TabIndex        =   7
      Top             =   0
      Width           =   3855
      _cx             =   6800
      _cy             =   1296
      FlashVars       =   ""
      Movie           =   "C:\Meus documentos\Video Locadora\megasoft.swf"
      Src             =   "C:\Meus documentos\Video Locadora\megasoft.swf"
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
   End
   Begin VB.CommandButton Cmdfechar 
      Caption         =   "Fechar"
      Height          =   255
      Left            =   3840
      MouseIcon       =   "FrmSobre.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3645
      Left            =   0
      ScaleHeight     =   3645
      ScaleWidth      =   735
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmSobre.frx":0FD4
      Height          =   1455
      Left            =   840
      TabIndex        =   8
      Top             =   960
      Width           =   3735
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   4680
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Via Local A, Nº 1997 - Fazenda Grande II - Cajazeiras"
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   3
      Top             =   3360
      Width           =   3795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tel.: (71) 3302-1304 / 9167-3212"
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
      Left            =   1320
      TabIndex        =   4
      Top             =   3120
      Width           =   2865
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "megasoft.corporation@gmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   4
      Left            =   1320
      MouseIcon       =   "FrmSobre.frx":1117
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2760
      Width           =   2355
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail.:"
      Height          =   195
      Index           =   3
      Left            =   840
      TabIndex        =   1
      Top             =   2760
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Autor.: Megasoft Corporation"
      Height          =   195
      Index           =   2
      Left            =   840
      TabIndex        =   0
      Top             =   2520
      Width           =   2025
   End
End
Attribute VB_Name = "FrmSobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdfechar_Click()
Unload FrmSobre
End Sub

