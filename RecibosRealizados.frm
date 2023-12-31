VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form RecibosRealizados 
   Caption         =   "Pagamentos Realizados"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   6030
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   6030
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btPesquisar 
      Caption         =   "Pesquisar"
      Height          =   315
      Left            =   3660
      TabIndex        =   0
      Top             =   60
      Width           =   1455
   End
   Begin VB.CheckBox ckData 
      Caption         =   "Data Inicial"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      ToolTipText     =   "O critério da data corresponde a data do orçamento"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txDtIni 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "31/12/2013"
      ToolTipText     =   "O critério da data da criação do Vale"
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txDtFim 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Text            =   "31/12/2013"
      ToolTipText     =   "O critério da data da criação do Vale"
      Top             =   60
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5985
      Left            =   105
      TabIndex        =   4
      Top             =   780
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   10557
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Final"
      Height          =   195
      Index           =   1
      Left            =   2220
      TabIndex        =   6
      ToolTipText     =   "O critério da data corresponde a data do orçamento"
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Pesquisa de pagamentos realizados a funcionários"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   360
      Width           =   5415
   End
   Begin VB.Menu mnuImpr 
      Caption         =   "Impressao Geral do Funcionário"
      Enabled         =   0   'False
   End
   Begin VB.Menu Mnu_Pop 
      Caption         =   "Mnu_Pop"
      Visible         =   0   'False
      Begin VB.Menu MnuTexto 
         Caption         =   "Texto"
      End
      Begin VB.Menu Mnu_Reimprime 
         Caption         =   "Re-Impressão"
      End
      Begin VB.Menu Mnu_Apagar 
         Caption         =   "Apagar"
      End
      Begin VB.Menu Mnu_Editar 
         Caption         =   "Editar"
      End
   End
End
Attribute VB_Name = "RecibosRealizados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'4.8.5 ReImpressão do recibo com todos os campos
'4.8.4 ReImpressão dos Vales
'4.8.4 Retorno da impressão geral dos vales
'4.8.3 Conserto da situação em que um registro dos vales não estava aparecendo
'4.7.9 Previsão para o registro duplicado [deve ser feita uma previsão na tela que grava também]
'4.6.6 Permitir excluir vales
'4.3.6 Total na impressão do vales
'4.3.6 Pesquisa de vale por data
'4.2.1 Impressão da observação
'3.6.8 Mostrar observação do recibo usando o clique direito
'3.5.1 Listagem dos vales dos funcionario
'3.4.9 Gravar as observações do vale
'3.4.7 Gravação dos recibos
'4.8.3 Conserto da situação em que um registro dos vales não estava aparecendo

Option Explicit

Private lcFuncionario%, lcNome$, NrRegistros%
Private tbGrid    As Recordset

'4.8.6 Edição de recibos
'4.8.3 Conserto da situação em que um registro dos vales não estava aparecendo
'Private Tp(4)     As String
'Private Tp(3)     As String

Private Pesquisou As Boolean

'4.6.6 Permitir excluir vales
Private IDSelec   As Long
Private JaDigSen  As Boolean

'4.8.4 ReImpressão dos Vales
Private lcEndereco  As String

Public Property Get Funcionario%()
Funcionario% = lcFuncionario%
End Property

'4.8.4 Retorno da impressão geral dos vales
'4.3.6 Pesquisa de vale por data
Private Sub Busca()
Dim Linha As Integer
Dim SQL   As String

Screen.MousePointer = vbHourglass

'4.3.6 Pesquisa de vale por data
If Pesquisou Then
    tbGrid.Close
    MSFlexGrid1.Rows = 2
End If

MSFlexGrid1.Enabled = True

'4.8.3 Conserto da situação em que um registro dos vales não estava aparecendo
SQL = "SELECT Valor, Data AS PrimeiroDeData, Tipo, Pago, obs, IdOperador, ID AS PrimeiroDeID "
SQL = SQL & "from Vales "
SQL = SQL & "Where IdOperador = " & lcFuncionario%
If ckData.Value = 1 Then
    SQL = SQL & " and Data Between " & DTSqls(txDtIni.Text) & " And " & DTSqls(txDtFim.Text, True)
End If
SQL = SQL & " ORDER BY Data"

''4.7.9 Previsão para o registro duplicado [deve ser feita uma previsão na tela que grava também]
'SQL = "SELECT Vales.Valor, First(Vales.Data) AS PrimeiroDeData, Vales.Tipo, Vales.Pago, Vales.obs, Vales.IdOperador,  First(Vales.ID) AS PrimeiroDeID "
'SQL = SQL & " from Vales "
'SQL = SQL & " GROUP BY Vales.Valor, Vales.Tipo, Vales.Pago, Vales.obs, Vales.IdOperador "
'SQL = SQL & " HAVING (First(Vales.Data)) "
'If ckData.Value = 1 Then
'    SQL = SQL & "Between " & DTSqls(txDtIni.Text) & " And " & DTSqls(txDtFim.Text, True)
'End If
'SQL = SQL & " AND ((Vales.IdOperador)=" & lcFuncionario% & ") "
'SQL = SQL & " ORDER BY First(Vales.Data) "

'SQL = "SELECT * From Vales Where IdOperador = " & lcFuncionario%
''4.3.6 Pesquisa de vale por data
'If ckData.Value = 1 Then
'    SQL = SQL & " and Vales.Data Between " & DTSqls(txDtIni.Text) & " And " & DTSqls(txDtFim.Text, True)
'End If
'SQL = SQL & " Order By Data Desc "

AbreTB tbGrid, SQL, dbOpenDynaset
If tbGrid.EOF = False Then
    tbGrid.MoveLast
    MSFlexGrid1.Rows = tbGrid.RecordCount + 1
    NrRegistros% = tbGrid.RecordCount
    Caption = "Recibos Realizadas " & NrRegistros% & " registros"
    tbGrid.MoveFirst
    Do While tbGrid.EOF = False
        Linha = Linha + 1
        MSFlexGrid1.TextMatrix(Linha, 0) = Format(tbGrid!Valor, "##,##0.00")
        MSFlexGrid1.TextMatrix(Linha, 1) = Format(tbGrid!PrimeiroDeData, "DD/MM/YYYY")
        MSFlexGrid1.TextMatrix(Linha, 2) = TpRecs(tbGrid!Tipo)
        If tbGrid!Tipo = 0 Then
            If tbGrid!PAGO > 0 Then
                MSFlexGrid1.TextMatrix(Linha, 3) = Format(tbGrid!PAGO, "DD/MM/YYYY")
            End If
        End If
        
        '3.4.9 Gravar as observações do vale
        MSFlexGrid1.TextMatrix(Linha, 4) = SN(tbGrid!Obs, vbString)
        
        '4.7.9 Previsão para o registro duplicado [deve ser feita uma previsão na tela que grava também]
        MSFlexGrid1.TextMatrix(Linha, 5) = Trim(STR((tbGrid!PrimeiroDeID)))
        '4.6.6 Permitir excluir vales
        'MSFlexGrid1.TextMatrix(Linha, 5) = Trim(STR((tbGrid!ID)))

        tbGrid.MoveNext
    Loop
    mnuImpr.Enabled = True
End If
Screen.MousePointer = vbDefault
Pesquisou = True
End Sub

Public Property Let Funcionario(ByVal vNewValue%)

'4.3.6 Pesquisa de vale por data
lcFuncionario% = vNewValue%
'Screen.MousePointer = vbHourglass
'SQL = "SELECT * From Vales Where IdOperador = " & lcFuncionario%
'SQL = SQL & " Order By Data Desc "
'
'AbreTB tbGrid, SQL, dbOpenDynaset
'If tbGrid.EOF = False Then
'    tbGrid.MoveLast
'    MSFlexGrid1.Rows = tbGrid.RecordCount + 1
'    NrRegistros% = tbGrid.RecordCount
'    Caption = "Recibos Realizadas " & NrRegistros% & " registros"
'    tbGrid.MoveFirst
'    Do While tbGrid.EOF = False
'        Linha = Linha + 1
'        MSFlexGrid1.TextMatrix(Linha, 0) = Format(tbGrid!Valor, "##,##0.00")
'        MSFlexGrid1.TextMatrix(Linha, 1) = Format(tbGrid!Data, "DD/MM/YYYY")
'        MSFlexGrid1.TextMatrix(Linha, 2) = TpRecs(tbGrid!Tipo)
'        If tbGrid!Tipo = 0 Then
'            If tbGrid!PAGO > 0 Then
'                MSFlexGrid1.TextMatrix(Linha, 3) = Format(tbGrid!PAGO, "DD/MM/YYYY")
'            End If
'        End If
'
'        '3.4.9 Gravar as observações do vale
'        MSFlexGrid1.TextMatrix(Linha, 4) = SN(tbGrid!Obs, vbString)
'
'        'MSFlexGrid1.c
'
'        'Num(Index).ToolTipText = RegLocal$
'
'        tbGrid.MoveNext
'    Loop
'End If
'Screen.MousePointer = vbDefault
End Property

'4.3.6 Pesquisa de vale por data
Private Sub btPesquisar_Click()
Busca
End Sub

'4.3.6 Pesquisa de vale por data
Private Sub btPesquisar_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

'4.3.6 Pesquisa de vale por data
Private Sub ckData_Click()
txDtIni.Enabled = (ckData.Value = 1)
txDtFim.Enabled = txDtIni.Enabled
End Sub

Private Sub Form_Load()
TpRecs(0) = "Adiantamento"
TpRecs(1) = "Comissão"
TpRecs(2) = "Vale Transporte"
TpRecs(3) = "Pagamento"

'4.8.3 Conserto da situação em que um registro dos vales não estava aparecendo
TpRecs(4) = "?"

MSFlexGrid1.ColWidth(0) = 1000
MSFlexGrid1.ColWidth(1) = 1000
MSFlexGrid1.ColWidth(2) = 1300
MSFlexGrid1.ColWidth(3) = 1000

'3.4.9 Gravar as observações do vale
MSFlexGrid1.ColWidth(4) = 1300

'4.6.6 Permitir excluir vales
MSFlexGrid1.ColWidth(5) = 1

MSFlexGrid1.TextMatrix(0, 0) = "Valor"
MSFlexGrid1.TextMatrix(0, 1) = "Data"
MSFlexGrid1.TextMatrix(0, 2) = "Tipo"
MSFlexGrid1.TextMatrix(0, 3) = "Pagamento"

'3.4.9 Gravar as observações do vale
MSFlexGrid1.TextMatrix(0, 4) = "Observações"

'4.6.6 Permitir excluir vales
MSFlexGrid1.TextMatrix(0, 5) = "ID"

'4.3.6 Pesquisa de vale por data
txDtFim.Text = Format(Now, "DD/MM/YYYY")
txDtIni.Text = Format(Now - 30, "DD/MM/YYYY")
Pesquisou = False

'4.6.6 Permitir excluir vales
JaDigSen = False
End Sub

Private Sub Mnu_Apagar_Click()
'4.6.6 Permitir excluir vales
If IDSelec > 0 Then
    If msgboxL("Tem certeza que quer apagar esse registro", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminação de registro") = vbYes Then
        If JaDigSen = False Then
            Load frmSenha
            frmSenha.Tipo = 1
            frmSenha.Show 1
            If frmSenha.Resultado = False Then
                Unload frmSenha
                Exit Sub
            End If
            JaDigSen = True
            Unload frmSenha
        End If
        ExecSql "Delete From Vales Where ID = " & IDSelec
        IDSelec = 0
        Busca
    End If
End If
End Sub

Private Sub Mnu_Editar_Click()
'4.8.6 Edição de recibos
Dim ID As Long

ID = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
Load ReciboEdicao
ReciboEdicao.NrVale = ID
ReciboEdicao.Show 1
MSFlexGrid1.Clear
Busca
End Sub

'4.8.4 ReImpressão dos Vales
Private Sub Mnu_Reimprime_Click()
Dim Tipo   As Integer
Dim Data   As Date
Dim Valor  As Double
Dim Vale   As Double
Dim Obs    As String
Dim Semana As String
Dim Folga  As String
Dim Tb As Recordset

Select Case MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
    Case "Adiantamento"
        Tipo = 0
    Case "Comissão"
        Tipo = 1
    Case "Vale Transporte"
        Tipo = 2
    Case "Pagamento"
        Tipo = 3
    Case Else
        Tipo = 4
End Select
Valor = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
Data = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
Obs = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)

'4.8.5 ReImpressão do recibo com todos os campos
AbreTB Tb, "Select * From Vales Where ID = " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5), dbOpenSnapshot
Semana = SN(Tb!Periodo, vbString)
Folga = SN(Tb!txValor, vbString)

Recibo.RecebeDados Nome$, Endereco, Valor, Semana, Vale, Folga, Obs, Tipo, Data
Recibo.Show
End Sub

'3.5.1 Listagem dos vales dos funcionario
Private Sub mnuImpr_Click()
Dim a%
Dim Aux As String
Dim Cap$, PAGO$
Dim Total As Currency

Const Linha$ = "-----------------------------------------------------"

'4.2.1 Impressão da observação
'Const TamFita = 55

Cap$ = Me.Caption
Caption = Me.Caption & " realizando impressão"
ImprBuferizada_Inicializa
LPRINT CENTRAL("RELACAO DE VALES", TamFita / 2)
LPRINT "Funcionario: " & Nome$
LPRINT Linha$
tbGrid.MoveFirst
Do While tbGrid.EOF = False
    Aux = ComplStr(Format(tbGrid!Valor, "##,##0.00"), 10, " ", 2) & " "
    Aux = Aux & Format(tbGrid!PrimeiroDeData, "DD/MM/YYYY") & " "
    Aux = Aux & ComplStr(TpRecs(tbGrid!Tipo), 16, " ", 0)
    PAGO$ = Space(10)
    If tbGrid!Tipo = 0 Then
        If tbGrid!PAGO > 0 Then
            PAGO$ = Format(tbGrid!PAGO, "DD/MM/YYYY")
        End If
    End If
    Aux = Aux & PAGO$ & " "
    Aux = Aux & SN(tbGrid!Obs, vbString)
    LPRINT Aux
    
    '4.3.6 Total na impressão do vales
    Total = Total + tbGrid!Valor
    
    tbGrid.MoveNext
Loop

'4.3.6 Total na impressão do vales
LPRINT Linha$
LPRINT ComplStr(Format(Total, "##,##0.00"), 10, " ", 2)

ImprBuferizada_Finaliza
Me.Caption = Cap$
End Sub

Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Public Property Get Nome$()
'3.5.1 Listagem dos vales dos funcionario
Nome$ = lcNome$
End Property

Public Property Let Nome(ByVal vNewValue$)
'3.5.1 Listagem dos vales dos funcionario
lcNome$ = vNewValue$
End Property

Private Sub MSFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
'3.6.8 Mostrar observação do recibo usando o clique direito
If Button = 2 Then
    IDSelec = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5))
    MnuTexto.Caption = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
    PopupMenu Mnu_Pop
End If
End Sub

'4.8.4 ReImpressão dos Vales
Public Property Get Endereco() As String
Endereco = lcEndereco
End Property

'4.8.4 ReImpressão dos Vales
Public Property Let Endereco(ByVal vNewValue As String)
lcEndereco = vNewValue
End Property
