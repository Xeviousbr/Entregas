VERSION 5.00
Object = "{00028C4A-0000-0000-0000-000000000046}#5.0#0"; "TDBG5.OCX"
Begin VB.Form Ferramentas 
   Caption         =   "Ferramentas"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6435
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data dataFuncionarios 
      Caption         =   "Consertos"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Z:\Share\Orcarro\OrCarro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   2  'Add New
      Exclusive       =   0   'False
      Height          =   435
      Left            =   -60
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT Nome FROM Mecanicos WHERE Nome > ''  AND Ativo = True and Oper = 0 ORDER BY Nome"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data Dados 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Z:\Share\Orcarro\OrCarro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Ferramentas"
      Top             =   540
      Visible         =   0   'False
      Width           =   2175
   End
   Begin TrueDBGrid50.TDBGrid Grid 
      Bindings        =   "Ferramentas.frx":0000
      Height          =   7455
      Left            =   120
      OleObjectBlob   =   "Ferramentas.frx":0014
      TabIndex        =   0
      Top             =   60
      Width           =   6135
   End
   Begin TrueDBGrid50.TDBDropDown tdbFunc 
      Bindings        =   "Ferramentas.frx":25F3
      Height          =   1275
      Left            =   0
      OleObjectBlob   =   "Ferramentas.frx":2612
      TabIndex        =   1
      Top             =   960
      Width           =   2940
   End
End
Attribute VB_Name = "Ferramentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'3.7.0 Ferramentas

Option Explicit

Private Sub Dados_Validate(Action As Integer, Save As Integer)
Dim nrFunc As Integer
Dim SQL    As String
Dim Nome   As String
Dim Codigo As String

Debug.Print Action
Select Case Action
    Case 6
        Grid.Col = 0
        Codigo = Grid.Text
        Grid.Col = 3
        Nome = Grid.Text
        If Nome > "" Then
            nrFunc = Consulta("Select codi From Mecanicos Where Nome = " & FA(Nome))
            InsertFerr Codigo, nrFunc, Nome
        End If
        
        '3.7.0 Ferramentas
        'Deu algum problema que não consegui pegar o "CÓDIGO" correto
        'Talvez precise trocar de componente
'    Case 9
'        Grid.Col = 0
'        Codigo = Grid.Text
'        Grid.Col = 3
'        Nome = Grid.Text
'        If Nome > "" Then
'            nrFunc = Consulta("Select codi From Mecanicos Where Nome = " & FA(Nome))
'            'Verificar se já tinha
'            SQL = "Select Count(*) From FerrMec Where idMec is not null "
'            SQL = SQL & " and codigo = " & FA(Codigo)
'
''            = " & nrFunc
''            SQL = SQL & " and codigo = " & FA(Codigo)
'            If Consulta(SQL) > 0 Then
'                'Se já tinha, ver se é o mesmo
'                'Se não é apagar o anterior
'            End If
'            'Criar o novo
'            InsertFerr Codigo, nrFunc, Nome
'        End If
End Select
End Sub

Private Sub InsertFerr(Codigo As String, nrFunc As Integer, Nome As String)
Dim SQL    As String

SQL = "Insert Into FerrMec (idMec, codigo, Data) Values (" & nrFunc
SQL = SQL & ", " & FA(Codigo)
SQL = SQL & ", " & DTSqld(Int(Now)) & ")"
ExecSql SQL
End Sub

Private Sub Form_Load()
Dados.DatabaseName = App.Path & "\OrCarro.mdb"
dataFuncionarios.DatabaseName = App.Path & "\OrCarro.mdb"
Dados.Enabled = True
dataFuncionarios.Enabled = True
End Sub

Private Sub Grid_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub
