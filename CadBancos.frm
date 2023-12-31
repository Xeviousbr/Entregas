VERSION 5.00
Begin VB.Form CadBancos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bancos"
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2940
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   2940
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1740
      TabIndex        =   4
      Top             =   480
      Width           =   1035
   End
   Begin VB.TextBox txNumero 
      Height          =   315
      Left            =   840
      MaxLength       =   20
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txBanco 
      Height          =   315
      Left            =   840
      MaxLength       =   20
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Número: "
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Banco: "
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   615
   End
End
Attribute VB_Name = "CadBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Mudando     As Boolean
Private Adicionando As Boolean

Public Sub Mostrar(NmBanco As String, NrBanco As String)
Mudando = True
txBanco.Text = NmBanco
txNumero.Text = NrBanco
txBanco.Tag = txBanco.Text
txNumero.Tag = txNumero.Text
Mudando = False
End Sub

Public Sub Adicionar()
Adicionando = True
End Sub

Private Sub Command1_Click()
Dim iNr  As Integer
Dim iaNr As Integer
Dim sNr  As String
Dim saNr As String

iNr = Val(txNumero.Text)
sNr = FA(Trim(STR(iNr)))
iaNr = Val(txNumero.Tag)
saNr = FA(Trim(STR(iaNr)))
If Adicionando Then
    ExecSql "Insert Into Bancos (Nome, Nr) Values (" & FA(txBanco.Text) & ", " & sNr & ")"
Else
    If txNumero.Text = txNumero.Tag Then
        ExecSql "Update Bancos Set Nome = " & FA(txBanco.Text) & " where Nr = " & sNr
    Else
        ExecSql "Update Fornecedores Set idBanco = " & sNr & " where idBanco = " & iaNr
        ExecSql "Update Bancos Set Nr = " & sNr & " where Nome = " & FA(txBanco.Text)
    End If
End If
Unload Me
End Sub

Private Sub Form_Load()
Mudando = False
Adicionando = False
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
        If txBanco.Text <> txBanco.Tag And txNumero.Text <> txNumero.Tag Then
            Hab = False
        Else
            Hab = True
        End If
    Else
        Hab = False
    End If
End If
Command1.Enabled = Hab
End Sub

Private Sub txNumero_Change()
If Mudando = False Then
    VeSePodeGravar
End If
End Sub
