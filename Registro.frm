VERSION 5.00
Begin VB.Form Registro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro do Orcarro"
   ClientHeight    =   2100
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5010
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNome 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1380
      TabIndex        =   2
      Top             =   1200
      Width           =   3075
   End
   Begin VB.TextBox Senha 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox ContraSenha 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "&Não Registrar"
      Height          =   375
      Left            =   3330
      TabIndex        =   4
      Top             =   1650
      Width           =   1155
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Registrar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   510
      TabIndex        =   3
      Top             =   1650
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nome:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   540
      TabIndex        =   9
      Top             =   1320
      Width           =   690
   End
   Begin VB.Label Label2 
      Caption         =   "para registra-lo ligue para ATC Informática 3072-5968 ou 9513-9696"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "Este programa esta em modo de demonstração."
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   4875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Contra Senha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   540
      TabIndex        =   6
      Top             =   960
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Senha para o Registro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   540
      TabIndex        =   5
      Top             =   540
      Width           =   2415
   End
End
Attribute VB_Name = "Registro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2.7.2 Logar todas mensagens
'2.0.3 Ajuste no funcionamento das permissões
'2.0.0 Operação de Registro

Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub ContraSenha_KeyDown(KeyCode As Integer, Shift As Integer)
If OKButton.Enabled = False Then
    If txtNome.Text > "" Then
        OKButton.Enabled = True
    End If
End If
End Sub

Private Sub Form_Load()
InicForm Me
Protecao.Gera
Senha.Text = Protecao.Senha
'ContraSenha.Text = Protecao.ContraSenha
End Sub

Private Sub OKButton_Click()
Dim VeSeDeu As Boolean

Protecao.ContraSenha = ContraSenha.Text
If Protecao.Confere(Senha.Text, ContraSenha.Text) Then
   Protecao.Implanta txtNome.Text
   FrmMenu.Caption = "Gerenciamento de Orçamentos " & App.Major & "." & App.Minor & "." & App.Revision & " Registrado para " & txtNome.Text
   '2.0.3 Ajuste no funcionamento das permissões
   Permissao = True
   FrmMenu.Mnu_Reg.Visible = False
   Unload Me
Else
   '2.7.2 Logar todas mensagens
   msgboxL "Contra Senha não confere", vbCritical, "Registro não efetivado"
   Protecao.Gera
   Senha.Text = Protecao.Senha
   ContraSenha.Text = ""
   OKButton.Enabled = False
   ContraSenha.SetFocus
End If
End Sub

Private Sub txtNome_KeyDown(KeyCode As Integer, Shift As Integer)
If OKButton.Enabled = False Then
    If txtNome.Text > "" And ContraSenha.Text > "" Then
        OKButton.Enabled = True
    End If
End If
End Sub
