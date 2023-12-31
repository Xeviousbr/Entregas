VERSION 5.00
Begin VB.Form frmRevisao 
   Caption         =   "Revisão Pendente"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6900
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6900
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Revisões"
      Height          =   495
      Left            =   5160
      TabIndex        =   7
      Top             =   2580
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cadastro"
      Height          =   495
      Left            =   2850
      TabIndex        =   5
      Top             =   2580
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   495
      Left            =   540
      TabIndex        =   4
      Top             =   2580
      Width           =   1215
   End
   Begin VB.Label lbRevisao 
      Caption         =   "lbRevisao"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2100
      Width           =   6615
   End
   Begin VB.Label lbTelefone 
      Caption         =   "lbTelefone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1620
      Width           =   6615
   End
   Begin VB.Label lbCliente 
      Caption         =   "lbCliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1140
      Width           =   6615
   End
   Begin VB.Label lbModelo 
      Caption         =   "lbModelo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "É necessário fazer a revisão do carro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   6555
   End
End
Attribute VB_Name = "frmRevisao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'5.0.4 Cadastro de Revisões

Option Explicit

Private lcNrCli As Integer

Private Sub Command1_Click()
Unload Me
End Sub

Public Property Let Revisao(ByVal vNewValue As Integer)
Dim SQL As String
Dim rsRevisao As Recordset

SQL = "SELECT Revisoes.Nome as RevNome, Revisoes.Meses, "
SQL = SQL & "Clientes.Nome as CliNome, Clientes.Telefone, "
SQL = SQL & "RevisoesCarros.Data,  DateAdd('m',-Revisoes.Meses,RevisoesCarros.Data) AS Tempo, "
SQL = SQL & "Carros.Modelo, Carros.Cor, Carros.Placa, "
SQL = SQL & "Clientes.NrCli "
SQL = SQL & "FROM ((Revisoes "
SQL = SQL & "INNER JOIN RevisoesCarros ON Revisoes.ID = RevisoesCarros.idRevisao) "
SQL = SQL & "INNER JOIN Carros ON RevisoesCarros.Placa = Carros.Placa) "
SQL = SQL & "INNER JOIN Clientes ON Carros.NrCli = Clientes.NrCli "
SQL = SQL & "WHERE RevisoesCarros.IdRevCarros = " & vNewValue
AbreTB rsRevisao, SQL, dbOpenSnapshot

lbModelo.Caption = rsRevisao!Modelo & " " & rsRevisao!Cor & " placa: " & rsRevisao!Placa

lbCliente.Caption = "Cliente: " & rsRevisao!CliNome
lbTelefone.Caption = "Telefone: " & rsRevisao!Telefone
lbRevisao.Caption = "Revisão de Tipo: " & rsRevisao!RevNome
lcNrCli = rsRevisao!NrCli
End Property

Private Sub Command2_Click()
Load frmClientes
frmClientes.NrCliente = lcNrCli
frmClientes.Show
Unload Me
End Sub

Private Sub Command3_Click()
ListaRevisoes.Show
End Sub

Private Sub Form_Load()
InicForm Me, SemAdapRes:=False
End Sub
