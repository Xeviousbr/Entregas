VERSION 5.00
Begin VB.Form Splash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Aguarde o Carregamento"
   ClientHeight    =   2475
   ClientLeft      =   2745
   ClientTop       =   3060
   ClientWidth     =   4290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2475
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Registrado para ..."
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1305
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mail: xeviousbr@gmail.com"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1920
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Versão: "
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contatos com Telefone: 3072-5968 ou 9513-9696"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   3540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ATC Informática"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sistema de Orçamentos para Mecânica de Automóveis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'3.3.6 Senha para troca de modo pelo modo balcão
'3.1.0 Modo Balcão
'2.5.2 Modo Restrito

Private Sub Form_Activate()
Dim Cliente As String

Cliente = PegaNmCliente
If Cliente > "" Then
    Label2(4).Caption = "Registrado para: " & Cliente
Else
    If Permissao Then
        Label2(4).Caption = "Modo Registrado"
    Else
        Label2(4).Caption = "Modo Share-Ware"
    End If
End If

'3.1.0 Modo Balcão
Select Case INI.ModoOperacao
    Case tpBalcao
        Label2(4).Caption = Label2(4).Caption & " Modo Balcão"
    Case tpMecanico
        Label2(4).Caption = Label2(4).Caption & " Modo Mecânico"
End Select
''2.5.2 Modo Restrito
'If INI.Restrito Then
'    Label2(4).Caption = Label2(4).Caption & " Somente Leitura"
'End If
End Sub

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
With App
   Versao = .Major & "." & .Minor & "." & .Revision
End With
#If V272 > 0 Then
    Label2(2).Caption = "Versão: 272-3"
#Else
    Label2(2).Caption = "Versão: " & Versao
#End If
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label2_Click(Index As Integer)
Dim Result As Boolean

If Index = 0 Then

    '3.3.6 Senha para troca de modo pelo modo balcão
    frmSenha.Show 1
    Result = frmSenha.Resultado
    Unload frmSenha
    If Result Then
        '3.1.0 Modo Balcão
        INI.ModoOperacao = tpEscritorio
        DefineMenus INI.ModoOperacao
        Unload Me
    End If
End If
End Sub


