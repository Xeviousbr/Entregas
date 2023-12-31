VERSION 5.00
Begin VB.Form frmObs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Observação"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5940
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Botao 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Index           =   0
      Left            =   4620
      TabIndex        =   3
      Top             =   4140
      Width           =   1215
   End
   Begin VB.CommandButton Botao 
      Caption         =   "&Gravar"
      Height          =   315
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   4140
      Width           =   1215
   End
   Begin VB.TextBox Texto 
      Height          =   3735
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   300
      Width           =   5775
   End
   Begin VB.Label lbNome 
      Alignment       =   2  'Center
      Caption         =   "..."
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5715
   End
End
Attribute VB_Name = "frmObs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'3.1.0 Modo Balcão
'2.5.6 Modo Restrito
'2.4.0 Campo para observação do cliente

Option Explicit

Private Sub Botao_Click(Index As Integer)
Dim SQL As String

If Index Then
    If frmClientes.Func Then
        Load frmSenha
        frmSenha.Tipo = 0
        frmSenha.Show 1
        If frmSenha.Resultado = False Then
            Unload frmSenha
            Exit Sub
        End If
        Unload frmSenha
    End If

    SQL = "Update Clientes Set Observacao = '" & Texto.Text & "'"
    SQL = SQL & "Where NrCli = " & clsCLi.NrCli
    ExecSql SQL
End If
Unload Me
End Sub

Private Sub Form_Load()
InicForm Me
lbNome.Caption = GCliente
Texto.Text = Consulta("Select Observacao From Clientes Where NrCli = " & clsCLi.NrCli)

'3.1.0 Modo Balcão
''2.5.6 Modo Restrito
'If INI.Restrito Then
'    Botao(1).Enabled = False
'End If
End Sub


