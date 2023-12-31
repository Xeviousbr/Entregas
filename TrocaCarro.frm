VERSION 5.00
Begin VB.Form TrocaCarro 
   Caption         =   "Troca Carro"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5310
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5310
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   435
      Left            =   4140
      TabIndex        =   8
      Top             =   2100
      Width           =   1095
   End
   Begin VB.CommandButton btOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   435
      Left            =   120
      TabIndex        =   7
      Top             =   2100
      Width           =   1095
   End
   Begin VB.ComboBox cbCliente 
      Height          =   315
      ItemData        =   "TrocaCarro.frx":0000
      Left            =   660
      List            =   "TrocaCarro.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Será transferido para o cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1380
      Width           =   3915
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Modelo: "
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   420
      Width           =   3855
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Placa: "
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   750
      Width           =   3855
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Cor: "
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   3
      Top             =   1065
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "O carro do cliente: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   5115
   End
   Begin VB.Label Label1 
      Caption         =   "&Nome:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1740
      Width           =   435
   End
End
Attribute VB_Name = "TrocaCarro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'3.7.8 Troca de carros entre os clientes

Option Explicit

Private lcCliente As String
Private lcPlaca As String
Private lcNrCli As Integer

Private Sub btOK_Click()
Dim CliQRecebeOCarro As Integer
Dim SQL   As String

SQL = "Update Orcamento Set Cliente = '" & Trim(cbCliente.Text)
SQL = SQL & "' Where Cliente = '" & lcCliente

'3.8.1 Conserto do sumiço do cliente 09/07/2014 na troca
'SQL = SQL & "' and Pagamento < 1 "

SQL = SQL & "' and Carro = '" & lcPlaca & "'"
ExecSql SQL

'3.8.1 Conserto do sumiço do cliente 09/07/2014 na troca
CliQRecebeOCarro = Consulta("Select NrCli from clientes Where Nome = '" & Trim(cbCliente.Text) & "'")

'3.8.1 Conserto do sumiço do cliente 09/07/2014 na troca
SQL = "Update Carros Set NrCli = " & CliQRecebeOCarro
SQL = SQL & " Where Placa = '" & lcPlaca & "'"
'SQL = "Update Carros Set NrCli = " & NrCli
'SQL = SQL & " Where Placa = '" & lcPlaca & "'"

ExecSql SQL
Unload Me
End Sub

Private Sub cbCliente_Click()
btOK.Enabled = True
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
InicForm Me
End Sub

'Public Property Get NmCliente()
'NmCliente = lcCliente
'End Property

Public Property Let NmCliente(ByVal vNewValue As String)
lcCliente = Trim(vNewValue)
End Property

Public Property Let Placa(ByVal vNewValue As String)
lcPlaca = vNewValue
End Property
