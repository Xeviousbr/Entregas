VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ListaRevisoes 
   Caption         =   "Revisões"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6600
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   6600
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cliente"
      Height          =   375
      Left            =   3300
      TabIndex        =   2
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mostrar Todas"
      Height          =   375
      Left            =   1620
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Bindings        =   "ListaRevisoes.frx":0000
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   10186
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
   End
End
Attribute VB_Name = "ListaRevisoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'5.0.4 Cadastro de Revisões

Option Explicit

Private Sub Command1_Click()
Mostra False
End Sub

Private Sub Command1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
   Unload Me
End If
End Sub

Private Sub Command2_Click()
Grid.Col = 0
Load frmClientes
frmClientes.NrCliente = Val(Grid.Text)
frmClientes.Show
End Sub

Private Sub Form_Click()
Grid.Clear
End Sub

Private Sub Form_Load()
InicForm Me

Grid.Row = 0
Grid.ColWidth(0) = 1

Grid.Col = 1
Grid.ColWidth(1) = 2000
Grid.Text = "Cliente"

Grid.Col = 2
Grid.ColWidth(2) = 2000
Grid.Text = "Carro"

Grid.Col = 3
Grid.ColWidth(3) = 1000
Grid.Text = "Placa"

Grid.Col = 4
Grid.ColWidth(4) = 1200
Grid.Text = "Revisão"

Grid.Col = 5
Grid.ColAlignment(5) = vbRightJustify
Grid.ColWidth(5) = 1900
Grid.Text = "Data para Revisão"
Mostra True
End Sub

Private Sub Mostra(Pendentes As Boolean)
Dim a      As Integer
Dim SQL    As String
Dim TbPesq As Recordset

SQL = "Select * From ("
SQL = SQL & "SELECT Clientes.NrCli, Clientes.Nome, Carros.Modelo, Carros.Cor, Carros.Placa, Revisoes.Nome, RevisoesCarros.Data as RevisaoNome, "
SQL = SQL & "DateAdd('m', Revisoes.Meses,RevisoesCarros.Data) as Tempo "
SQL = SQL & "FROM ((RevisoesCarros "
SQL = SQL & "INNER JOIN Carros ON RevisoesCarros.Placa = Carros.Placa) "
SQL = SQL & "INNER JOIN Clientes ON Carros.NrCli = Clientes.NrCli) "
SQL = SQL & "INNER JOIN Revisoes ON RevisoesCarros.idRevisao = Revisoes.ID ) X "
If Pendentes Then
    SQL = SQL & "Where Tempo < Now "
End If
SQL = SQL & " Order By Tempo desc"
Screen.MousePointer = vbArrow
AbreTB TbPesq, SQL, dbOpenSnapshot
On Local Error GoTo Vazio
TbPesq.MoveFirst
On Local Error GoTo 0
Grid.Rows = 1
Do While TbPesq.EOF = False
    a = a + 1
    Grid.AddItem TbPesq(0)
    Grid.Row = a
    Grid.Col = 1
    Grid.Text = TbPesq(1) & "  " & TbPesq(2)
    Grid.Col = 2
    Grid.Text = TbPesq(3)
    Grid.Col = 3
    Grid.Text = TbPesq(4)
    Grid.Col = 4
    Grid.Text = TbPesq(5)
    Grid.Col = 5
    Grid.Text = TbPesq(6)
    TbPesq.MoveNext
Loop
Screen.MousePointer = vbDefault

Vazio:
Screen.MousePointer = vbDefault
End Sub

Private Sub Grid_DblClick()
Command2_Click
End Sub

Private Sub Grid_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
   Unload Me
ElseIf KeyCode = vbKeyReturn Then
   Command2_Click
End If
End Sub
