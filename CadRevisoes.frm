VERSION 5.00
Begin VB.Form CadRevisoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Revisões"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2835
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   2835
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txBanco 
      Height          =   315
      Left            =   780
      MaxLength       =   20
      TabIndex        =   0
      Top             =   60
      Width           =   1935
   End
   Begin VB.TextBox txNumero 
      Height          =   315
      Left            =   780
      MaxLength       =   20
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1740
      TabIndex        =   2
      Top             =   420
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Revisão: "
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Meses: "
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Width           =   615
   End
End
Attribute VB_Name = "CadRevisoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'5.0.5 Alteração do nome da Revisão
'5.0.4 Cadastro de Revisões

Option Explicit

Private Mudando     As Boolean
Private Adicionando As Boolean
Private NomeOld     As String

Public Sub Mostrar(NmBanco As String, NrBanco As String, Kms As String)
Mudando = True
txBanco.Text = NmBanco
txNumero.Text = NrBanco
txBanco.Tag = txBanco.Text
txNumero.Tag = txNumero.Text
Mudando = False
NomeOld = NmBanco
End Sub

Public Sub Adicionar()
Adicionando = True
End Sub

Private Sub Command1_Click()
Dim iNr  As Long
Dim iaNr As Long
Dim sNr  As String
Dim saNr As String

iNr = Val(txNumero.Text)
sNr = FA(Trim(STR(iNr)))
'iaNr = Val(txKms.Text)
saNr = Trim(STR(iaNr))
If Adicionando Then
    ExecSql "Insert Into Revisoes (Nome, Meses) Values (" & FA(txBanco.Text) & ", " & sNr & ")"
Else
    '5.0.5 Alteração do nome da Revisão
    ExecSql "Update Revisoes Set Nome = " & FA(txBanco.Text) & " ,Meses = " & sNr & " where Nome = " & FA(NomeOld)
End If
Unload Me
End Sub

Private Sub Form_Load()
Mudando = False
Adicionando = False
InicForm Me
End Sub

Private Sub txBanco_Change()
If Mudando = False Then
    VeSePodeGravar
End If
End Sub

Private Sub VeSePodeGravar()
Dim Hab As Boolean

If Adicionando Then
    Hab = ((txBanco.Text > "") And (txNumero.Text > ""))
Else
    If txBanco.Text <> txBanco.Tag Or txNumero.Text <> txNumero.Tag Then
        Hab = True
    Else
        Hab = False
    End If
End If
Command1.Enabled = Hab
End Sub

Private Sub txKms_Change()
If Mudando = False Then
    VeSePodeGravar
End If
End Sub

Private Sub txNumero_Change()
If Mudando = False Then
    VeSePodeGravar
End If
End Sub

