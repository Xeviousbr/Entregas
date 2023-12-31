VERSION 5.00
Begin VB.Form frmSenha 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe a senha"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2565
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   2565
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Botao 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Index           =   0
      Left            =   1380
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Botao 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txSenha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1020
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Senha: "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "frmSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'4.1.7 Logar troca de modo de operação
'3.3.5 Alteração da senha geral
'2.7.2 Logar todas mensagens
'2.4.0 Desfazer o pagamento

Option Explicit

'Senha
'Tipo 0 Pede a senha do operador
'Tipo 1 Pede a senha geral

Private gOK    As Boolean

'3.1.4 Senhas com ** para a identificação do balconista
Private lcTipo  As Integer
Private lcSenha As String

Private Sub Botao_Click(Index As Integer)
Dim SQL As String

gOK = False

If Index Then

    '3.1.4 Senhas com ** para a identificação do balconista
    If lcTipo = 1 Then
        If txSenha.Text > "" Then
            lcSenha = txSenha.Text
            gOK = True
        End If
    Else
    
        '3.3.5 Alteração da senha geral
        If LCase(txSenha.Text) = ("atch" & NumDiaSem()) Then
        'If LCase(txSenha.Text) = "atch" Then
        
            '4.1.7 Logar troca de modo de operação
            Dim ModoAtual As String
            If INI.ModoOperacao = tpBalcao Then
                ModoAtual = "Balcão"
            Else
                ModoAtual = "Mecânico"
            End If
            Loga "Alterado do modo " & ModoAtual & " para o modo Escritório "
        
            gOK = True
        Else
            '2.7.2 Logar todas mensagens
            msgboxL "Senha não confirmada", vbExclamation
        End If
    End If
End If
Me.Hide
End Sub

Public Property Get Resultado() As Boolean
Resultado = gOK
End Property

'3.1.4 Senhas com ** para a identificação do balconista
Public Property Let Tipo(ByVal vNewValue As Integer)
lcTipo = vNewValue
End Property

'3.1.4 Senhas com ** para a identificação do balconista
Public Property Get Senha() As String
Senha = lcSenha
End Property

