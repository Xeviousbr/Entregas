VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBusca 
   Caption         =   "Pesquisa de Mecânicos"
   ClientHeight    =   4275
   ClientLeft      =   3570
   ClientTop       =   2670
   ClientWidth     =   4020
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmOpcoes"
   MaxButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Bindings        =   "frmBusca.frx":0000
      Height          =   4155
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   7329
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Menu mnu 
      Caption         =   ""
      Enabled         =   0   'False
      Index           =   2
      Visible         =   0   'False
   End
   Begin VB.Menu mnu 
      Caption         =   ""
      Enabled         =   0   'False
      Index           =   4
      Visible         =   0   'False
   End
   Begin VB.Menu mnu 
      Caption         =   ""
      Enabled         =   0   'False
      Index           =   5
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmBusca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'5.0.4 Cadastro de Revisões
'4.4.2 Cadastro de fornecedores
'3.5.1 Não excluir fisicamente Mecânico
'2.7.5 Taréfas Dinâmicas
'2.7.2 Logar todas mensagens
'2.6.7 Ajuste a primeira operação com cadastro de mecânicos
'2.6.3 Cadastro de Mecânicos

Option Explicit

Const H = 4680
Const W = 3900

Private l_Tipo As Integer

Private Sub Form_Load()
'4.4.2 Cadastro de fornecedores
'Dim a      As Integer
'Dim TbPesq As Recordset

'4.4.2 Cadastro de fornecedores
'Set TbPesq = db.OpenRecordset("Select * From Mecanicos Where codi > 0 and Ativo = True Order By Nome ")
'3.5.1 Não excluir fisicamente Mecânico
'Set TbPesq = db.OpenRecordset("Select * From Mecanicos Where codi > 0 Order By Nome")

Grid.ColWidth(0) = 2100
Grid.ColWidth(1) = 1700

End Sub

Private Sub Grid_DblClick()
Escolheu
End Sub

Private Sub Escolheu()
Dim Col1 As String

Grid.Col = 0
Select Case Tipo
    Case Is = 0
        Load CadMecanicos
        CadMecanicos.Mostrar Grid.Text
        CadMecanicos.Show
        
        '4.4.2 Cadastro de fornecedores
    Case Is = 1
        Load Forneced
        Forneced.Mostrar Grid.Text
        Forneced.MostraLugar
        Forneced.Show
        
    Case Is = 2
        Col1 = Grid.Text
        Grid.Col = 1
        Load CadBancos
        CadBancos.Mostrar Col1, Grid.Text
        CadBancos.Show
        
    '5.0.4 Cadastro de Revisões
    Case Is = 3
        Dim Kms As Double
        Col1 = Grid.Text
        Grid.Col = 1
        Kms = Consulta("Select Kms from Revisoes Where Nome = '" & Col1 & "'")
        Load CadRevisoes
        CadRevisoes.Mostrar Col1, Grid.Text, Trim(STR(Kms))
        CadRevisoes.Show
        
        
'    '4.7.6 Pesquisar por cliente ou funcionário no recibo avulso
'    Case Is = 3
'        Col1 = Grid.Text
'        Grid.Col = 1
'        If Tipo = "Cliente" Then
'            SQL = "Select Ender From Clientes Where Nome = " & FA(Col1)
'        Else
'
'        End If
'        Load frmReciboAvulso
'        'frmReciboAvulso.Fornec=
'        frmReciboAvulso.txRecebe = EdtFornec.Text
'        frmReciboAvulso.txIdent = txForn(3).Text
'        frmReciboAvulso.txEndereco = Ender.Text
'        frmReciboAvulso.Show
'
'        CadBancos.Mostrar Col1, Grid.Text
'        CadBancos.Show

End Select
Unload Me
End Sub

Private Sub Grid_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
   Unload Me
ElseIf KeyCode = vbKeyReturn Then
   Escolheu
End If
End Sub

'4.4.2 Cadastro de fornecedores
Public Property Get Tipo() As Integer
Tipo = l_Tipo
End Property

'4.4.2 Cadastro de fornecedores
Public Property Let Tipo(ByVal vNewValue As Integer)
Dim a      As Integer
Dim b      As Integer
Dim TbPesq As Recordset
Dim SQL    As String

l_Tipo = vNewValue
Select Case l_Tipo
    Case 0
        Grid.Row = 0: Grid.Col = 0: Grid.Text = "Mecânico"
        Grid.Row = 0: Grid.Col = 1: Grid.Text = "Telefone"
        Set TbPesq = db.OpenRecordset("Select Nome, Telefone From Mecanicos Where codi > 0 and Ativo = True Order By Nome ")
    Case 1
        Grid.Row = 0: Grid.Col = 0: Grid.Text = "Nome"
        Grid.Row = 0: Grid.Col = 1: Grid.Text = "Telefone"
        Set TbPesq = db.OpenRecordset("Select Nome, Telefone From Fornecedores Order By Nome ")
        Caption = "Pesquisa de Fornecedores"
    Case 2
        Grid.Row = 0: Grid.Col = 0: Grid.Text = "Banco"
        Grid.Row = 0: Grid.Col = 1: Grid.Text = "Nro"
        Set TbPesq = db.OpenRecordset("Select Nome, Nr From Bancos Order By Nome ")
        
    '5.0.4 Cadastro de Revisões
    Case 3
        Caption = "Pesquisa de Revisões"
        Grid.Row = 0: Grid.Col = 0: Grid.Text = "Revisão"
        Grid.Row = 0: Grid.Col = 1: Grid.Text = "Meses"
        Set TbPesq = db.OpenRecordset("Select Nome, Meses From Revisoes Order By Nome ")
        
        
'    Case 3
'        Grid.Row = 0: Grid.Col = 0: Grid.Text = "Nome"
'        Grid.Row = 0: Grid.Col = 1: Grid.Text = "Telefone"
'        '4.7.6 Pesquisar por cliente ou funcionário no recibo avulso
''        SQL = "Select Distinct Nome, Origem "
''        SQL = SQL & "From ( "
''        SQL = SQL & "SELECT Nome, 'Funcionário' as Origem "
''        SQL = SQL & "From Mecanicos "
''        SQL = SQL & "Union "
''        SQL = SQL & "SELECT Distinct Nome, 'Cliente' as Origem "
''        SQL = SQL & "From Clientes "
''        SQL = SQL & ") X "
''        SQL = SQL & "Where Nome > '' "
''        SQL = SQL & "ORDER BY Nome "
'        Screen.MousePointer = vbHourglass
'        Set TbPesq = db.OpenRecordset("Select Nome, Telefone From Bancos Order By Nome ")
'
'        Set TbPesq = db.OpenRecordset(SQL)
End Select
TbPesq.MoveFirst
'On Error GoTo 0

Do While TbPesq.EOF = False
    a = a + 1
    b = a + 1
    
    Grid.AddItem TbPesq(0)
    'Grid.AddItem TbPesq(0), b
    
    Grid.Row = a
    Grid.Col = 1
    Grid.Text = SN(TbPesq(1))
    TbPesq.MoveNext
Loop
Screen.MousePointer = vbDefault

'For a = 0 To TbPesq.RecordCount
''For a = 0 To TbPesq.RecordCount - 1
'
'   Grid.AddItem TbPesq(0), a + 1
'   Grid.Row = a + 1
'   Grid.Col = 1
'   Grid.Text = SN(TbPesq(1))
'   TbPesq.MoveNext
'Next
End Property
