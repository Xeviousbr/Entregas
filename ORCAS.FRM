VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Orcas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orçamentos de "
   ClientHeight    =   7485
   ClientLeft      =   2055
   ClientTop       =   2745
   ClientWidth     =   6000
   Icon            =   "ORCAS.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7485
   ScaleWidth      =   6000
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cbPagos 
      Height          =   315
      ItemData        =   "ORCAS.frx":0442
      Left            =   1920
      List            =   "ORCAS.frx":044F
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   2220
      Width           =   2115
   End
   Begin VB.TextBox txPlaca 
      Height          =   285
      Left            =   3960
      TabIndex        =   2
      Top             =   60
      Width           =   915
   End
   Begin VB.OptionButton opObs 
      Caption         =   "Histórico"
      Height          =   255
      Index           =   2
      Left            =   3540
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2760
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Observações do Cliente"
      Height          =   2235
      Left            =   60
      TabIndex        =   12
      Top             =   5160
      Width           =   5835
      Begin VB.TextBox txObsCli 
         Height          =   1515
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   240
         Width           =   5655
      End
      Begin VB.CommandButton btGrvCliente 
         Caption         =   "Gravar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   60
      TabIndex        =   9
      Top             =   2580
      Width           =   5835
      Begin VB.CommandButton btImpr 
         Caption         =   "Impressão"
         Height          =   375
         Left            =   4500
         TabIndex        =   22
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txObs 
         Height          =   1515
         Index           =   2
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.TextBox txObs 
         Height          =   1515
         Index           =   1
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.OptionButton opObs 
         Caption         =   "Observação"
         Height          =   255
         Index           =   0
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   180
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton opObs 
         Caption         =   "Observação do Mecânico"
         Height          =   255
         Index           =   1
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   180
         Width           =   2115
      End
      Begin VB.CommandButton btGrvOrc 
         Caption         =   "Gravar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txObs 
         Height          =   1515
         Index           =   0
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   480
         Width           =   5655
      End
   End
   Begin VB.CommandButton btPesq 
      Caption         =   "&Pesquisar"
      Height          =   315
      Left            =   5040
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txDtFIM 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Text            =   "31/12/2888"
      Top             =   60
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txDtINI 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Text            =   "31/12/2888"
      Top             =   60
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1785
      Left            =   60
      TabIndex        =   6
      Top             =   360
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   3149
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Acionar"
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Placa: "
      Height          =   195
      Index           =   2
      Left            =   3420
      TabIndex        =   20
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "à"
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Período de"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   795
   End
End
Attribute VB_Name = "Orcas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'5.0.8 Ordenação da pesquisa de orçamentos por cliente, pela data
'5.0.6 Troca do SQL de seleção da pesquisa de orçamentos para resolver problema de duplicidade
'4.7.1 Retorno da permissão do mecânico ver o orçamento
'4.7.0 Retorno do menu para orçamentos mas sem acionar o orçamento
'4.2.4 Imprimir também o balconísta que criou o orçamento na impressão da observação
'4.2.3 Mais informação na impressão da observação
'4.2.2 Ajuste na Impressão da observação
'4.2.1 Impressão da observação
'3.9.6 Troca da mensagem de pago em minúsculo para orc.parc.
'3.9.1 Pagamento com parcelas
'3.8.7 Permitir lançamentos em orçamentos pagos
'3.7.8 Troca de carros entre os clientes
'3.6.9 Permitir mecânicos verem os orçamentos, mas sem os valores
'3.6.9 Filtro pagos ou nao pagos na pesquisa de orçamentos
'3.6.2 Permitir atualizar o histórico pela pesquisa de orçamentos
'3.6.1 Conserto da situação de duplicidade na pesquisa de orçamentos
'3.4.5 Impedir que funcionario edite seus orçamentos
'3.2.6 Dar mensagem de alerta em caso de comando de apagar orçamento
'3.2.6 Apagar as tarefas quando for apagado o orçamento
'3.1.7 Impedir que os mecânicos digitem nas outras observações, na pesquisa de orçamentos
'3.1.7 Melhorar a mensagem quando a pesquisa não da resultado, na pesquisa de orçamentos
'3.1.7 Impedir de dar erro caso a pesquisa de orçamentos não de resultado
'3.1.7 Inversão da ordenação na pesquisa de orçamentos, para pesquisa por placa também
'3.1.6 Inversão da ordenação na pesquisa de orçamentos
'3.1.1 Nao mostrar a pesquisa pela placa na pesquisa de orcamentos quando esta filtrado pelo nome
'3.1.1 Previsao para placa nao encontrada na pesquisa da placa na lista de orcamentos
'3.1.0 Modo Balcão
'2.8.8 Filtro para as placas na pesquisa dos orçamentos
'2.8.7 Conserto do erro da pesquisa pelo orçamento pela placa
'2.8.7 Mostrar direto a obs do mecânico na pesquisa de orçamentos, quando em modo restrito
'2.8.4 Otimizar a Obtenção do Nr do Cliente
'2.8.4 Gravar no log o tempo de carregamento da tela de orçamento
'2.7.5 Taréfas Dinâmicas
'2.7.3 Histórico do Orçamento
'2.7.2 Impedir duplicações na pesquisa de orçamentos
'2.7.2 Conserto da busca na relação de orçamentos
'2.7.2 Logar todas mensagens
'2.6.2 Impedir de mostrar o orçamento caso desse duplo-click na lista de orçamentos
'2.6.1 Permitir escritório escrever na observação do cliente
'2.6.1 Mecânicos não devem ver valores
'2.6.0 Observação do mecânico
'2.5.6 Modo Restrito
'2.4.7 Ajuste da alteração das observações quanto as pesquisas de placas
'2.4.6 Observações dos clientes na listagem de orçamentos e gravação das observações
'2.4.2 Observação do orçamento na relação mensal
'2.3.1 Previsão para o bug que não aparecia o carro certo quando se acionava um novo orçamento via cadastro de clientes

Private GNmCliente As String
Private GMensal    As String

'2.8.7 Mostrar direto a obs do mecânico na pesquisa de orçamentos, quando em modo restrito
Private PasNoActiv As Boolean

'3.1.7 Impedir de dar erro caso a pesquisa de orçamentos não de resultado
Dim Encontrou As Boolean

'3.6.2 Permitir atualizar o histórico pela pesquisa de orçamentos
Private Mudando As Boolean

'3.9.1 Pagamento com parcelas
Private Lin1 As Integer
  
Public Property Get SelCamppos() As String
SelCamppos = "SELECT DISTINCTROW Orcamento.Data, Orcamento.Total, Orcamento.Kilom, "
SelCamppos = SelCamppos & "Carros.Modelo, Orcamento.Cliente, Orcamento.Carro, Orcamento.Orcamento, Orcamento.Pagamento, "
SelCamppos = SelCamppos & "Clientes.Ender, Carros.Placa, Clientes.NrCli "
SelCamppos = SelCamppos & ",Orcamento.Obs "
SelCamppos = SelCamppos & ",Clientes.Observacao, Clientes.NrCli "
SelCamppos = SelCamppos & ",Orcamento.ObsMec "
SelCamppos = SelCamppos & ",Carros.Historico "
SelCamppos = SelCamppos & ",Clientes.Funcionario "
SelCamppos = SelCamppos & ",Orcamento.VlrPago "
End Property

'4.2.1 Impressão da observação
Private Sub Localiza()
Dim a As Integer

Lin1 = MSFlexGrid1.Top
MomChamOrc = Now
Loga "Carregar Orçameto"
OrcasRecordset.MoveFirst
a = 1
Do While a < MSFlexGrid1.Row
    OrcasRecordset.MoveNext
    a = a + 1
Loop
End Sub

Private Sub Aciona()
'Dim a As Integer

'3.1.7 Impedir de dar erro caso a pesquisa de orçamentos não de resultado
If Encontrou = False Then
    MsgBox "Refaça a pesquisa antes de acionar o orçamento"
Else

    '4.2.1 Impressão da observação
    Localiza
'    Lin1 = MSFlexGrid1.Top
'
'    '2.8.4 Gravar no log o tempo de carregamento da tela de orçamento
'    MomChamOrc = Now
'
'    Loga "Carregar Orçameto"
'    OrcasRecordset.MoveFirst
'    a = 1
'    Do While a < MSFlexGrid1.Row
'        OrcasRecordset.MoveNext
'        a = a + 1
'    Loop
    
    '2.7.5 Taréfas Dinâmicas
    CarregaTarefas OrcasRecordset("Orcamento")
    
    Load frmOrc
    frmOrc.NrOrcamento = OrcasRecordset("Orcamento")
    frmOrc.Show
End If
End Sub

Private Sub btGrvCliente_Click()
Dim SQL As String

'3.4.5 Impedir que funcionario edite seus orçamentos
MSFlexGrid1.Col = 11
If MSFlexGrid1.Text = "True" Then
    Load frmSenha
    frmSenha.Tipo = 0
    frmSenha.Show 1
    If frmSenha.Resultado = False Then
        Unload frmSenha
        Exit Sub
    End If
    Unload frmSenha
End If

'2.4.6 Observações dos clientes na listagem de orçamentos e gravação das observações
SQL = "Update Clientes Set Observacao = '" & txObsCli.Text & "' Where NrCli = " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8)
ExecSql SQL
RestauraBotoes
End Sub

Private Sub btGrvOrc_Click()
Dim SQL As String

'3.6.2 Permitir atualizar o histórico pela pesquisa de orçamentos
Mudando = True

'3.4.5 Impedir que funcionario edite seus orçamentos
MSFlexGrid1.Col = 11
If MSFlexGrid1.Text = "True" Then
    Load frmSenha
    frmSenha.Tipo = 0
    frmSenha.Show 1
    If frmSenha.Resultado = False Then
        Unload frmSenha
        
        '3.6.2 Permitir atualizar o histórico pela pesquisa de orçamentos
        Mudando = False
        
        Exit Sub
    End If
    Unload frmSenha
End If

'3.1.0 Modo Balcão
If INI.ModoOperacao = tpMecanico Then
'2.6.0 Observação do mecânico
'If INI.Restrito Then
    SQL = "Update Orcamento Set ObsMec = '" & txObs(1).Text & "' Where Orcamento = " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)
Else
    '2.6.1 Permitir escritório escrever na observação do cliente
    SQL = "Update Orcamento Set Obs = '" & txObs(0).Text & "', ObsMec = '" & txObs(1).Text & "' Where Orcamento = " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)
    'SQL = "Update Orcamento Set Obs = '" & txObs(0).Text & "' Where Orcamento = " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)
    '2.4.6 Observações dos clientes na listagem de orçamentos e gravação das observações
    'SQL = "Update Orcamento Set Obs = '" & txObs.Text & "' Where Orcamento = " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)
    ExecSql SQL
        
    If txObs(2).Text > "" Then
        SQL = "Update Carros Set Historico = '" & txObs(2).Text & "' "
    Else
        SQL = "Update Carros Set Historico = null "
    End If
    MSFlexGrid1.Col = 4
    SQL = SQL & " Where Placa = '" & MSFlexGrid1.Text & "'"
End If
ExecSql SQL
RestauraBotoes

'3.6.2 Permitir atualizar o histórico pela pesquisa de orçamentos
Mudando = False
End Sub

'4.2.1 Impressão da observação
Private Sub btImpr_Click()
Dim SQL    As String
Dim rsImpr As Recordset

btImpr.Enabled = False
Localiza

'4.2.4 Imprimir também o balconísta que criou o orçamento na impressão da observação
SQL = "SELECT Orcamento.Orcamento, Carros.Placa, Carros.Modelo, Carros.Cor, Orcamento.ObsMec, Orcamento.Cliente, Orcamento.Data, Mecanicos.Nome "
SQL = SQL & "FROM Mecanicos "
SQL = SQL & "INNER JOIN (Orcamento "
SQL = SQL & "INNER JOIN Carros ON Orcamento.Carro = Carros.Placa) "
SQL = SQL & "ON Mecanicos.codi = Orcamento.Vend "
SQL = SQL & "WHERE Orcamento.Orcamento= " & OrcasRecordset!Orcamento
'SQL = "SELECT Orcamento.Orcamento, Carros.Placa, Carros.Modelo, Carros.Cor, Orcamento.Obs "
''4.2.3 Mais informação na impressão da observação
'SQL = SQL & ", Orcamento.Cliente, Orcamento.Data "
'
'SQL = SQL & "FROM Orcamento "
'SQL = SQL & "INNER JOIN Carros ON Orcamento.Carro = Carros.Placa "
'
'SQL = SQL & "WHERE Orcamento.Orcamento = " & OrcasRecordset!Orcamento
AbreTB rsImpr, SQL

'4.2.4 Imprimir também o balconísta que criou o orçamento na impressão da observação
ImpressaoObs rsImpr!Modelo, rsImpr!Cor, rsImpr!Placa, rsImpr!ObsMec, rsImpr!Cliente, rsImpr!Data, rsImpr!Nome
'4.2.3 Mais informação na impressão da observação
'ImpressaoObs rsImpr!Modelo, rsImpr!Cor, rsImpr!Placa, rsImpr!ObsMec, rsImpr!Cliente, rsImpr!Data
'4.2.2 Ajuste na Impressão da observação
'ImpressaoObs rsImpr!Modelo, rsImpr!Cor, rsImpr!Placa, rsImpr!ObsMec

btImpr.Enabled = True
End Sub

Private Sub btPesq_Click()
Dim SQL As String

If IsDate(txDtINI.Text) And IsDate(txDtFIM.Text) Then

    SQL = SelCamppos()
    
    SQL = SQL & "FROM Orcamento, Carros, Clientes "
    SQL = SQL & "WHERE Orcamento.Carro = Carros.Placa and Orcamento.Data Between " & DTSqls(txDtINI.Text) & " And " & _
        DTSqls(txDtFIM.Text, True)
    SQL = SQL & " and Clientes.Nome = Orcamento.Cliente "
    
    '2.8.8 Filtro para as placas na pesquisa dos orçamentos
    If txPlaca.Text > "" Then
        SQL = SQL & " and Orcamento.Carro = '" & UCase(txPlaca.Text) & "'"
    End If
    
    '3.6.9 Filtro pagos ou nao pagos na pesquisa de orçamentos
    Select Case cbPagos.ListIndex
        Case 1
            SQL = SQL & " and Orcamento.Pagamento > 0 "
        Case 2
            SQL = SQL & " and Orcamento.Pagamento = 0 "
    End Select
    
    '3.1.7 Inversão da ordenação na pesquisa de orçamentos, para pesquisa por placa também
    SQL = SQL & " Order By Orcamento.Data desc "
    
    AbreTB OrcasRecordset, SQL, dbOpenSnapshot
    MostraOrcs
Else
    If IsDate(txDtFIM.Text) Then
        '2.7.2 Logar todas mensagens
        msgboxL "A data INICIAL não esta informada corretamente"
        txDtINI.SetFocus
    Else
        txDtFIM.SetFocus
        '2.7.2 Logar todas mensagens
        msgboxL "A data FINAL não esta informada corretamente"
    End If
End If
End Sub

Private Sub cbPagos_Click()
'3.6.9 Filtro pagos ou nao pagos na pesquisa de orçamentos
If Mudando = False Then
    btPesq_Click
End If
End Sub

Private Sub Command1_Click(Index As Integer)
If Index Then
    Unload Orcas
Else
    Aciona
End If
End Sub

Private Sub DBGrid1_DblClick()
Aciona
End Sub

Private Sub Form_Activate()

'3.9.1 Pagamento com parcelas
If Lin1 > 0 Then
    MostraOrcs
    MSFlexGrid1.Top = Lin1
    Lin1 = 0
    
Else

    '2.8.7 Mostrar direto a obs do mecânico na pesquisa de orçamentos, quando em modo restrito
    If PasNoActiv = False Then
        PasNoActiv = True
        
        '3.1.0 Modo Balcão
        If INI.ModoOperacao = tpMecanico Then
        'If INI.Restrito Then
        
            '4.7.1 Retorno da permissão do mecânico ver o orçamento
            '4.7.0 Retorno do menu para orçamentos mas sem acionar o orçamento
            'Command1(0).Enabled = False
        
            opObs(1).Value = True
        End If
    End If
End If
End Sub

Private Sub Form_Load()
'2.8.7 Mostrar direto a obs do mecânico na pesquisa de orçamentos, quando em modo restrito
PasNoActiv = False

InicForm Me
MSFlexGrid1.ColWidth(0) = 1000

'3.6.9 Filtro pagos ou nao pagos na pesquisa de orçamentos
Mudando = True

'3.1.0 Modo Balcão
If INI.ModoOperacao = tpMecanico Then
'2.6.1 Mecânicos não devem ver valores
'If INI.Restrito Then

    '3.6.9 Permitir mecânicos verem os orçamentos, mas sem os valores
    'Command1(0).Enabled = False
    
    MSFlexGrid1.ColWidth(1) = 0
    MSFlexGrid1.ColWidth(3) = 2700
    
    '3.1.7 Impedir que os mecânicos digitem nas outras observações, na pesquisa de orçamentos
    txObs(0).Locked = True
    txObs(2).Locked = True
    txObsCli.Locked = True
    
    '3.6.9 Filtro pagos ou nao pagos na pesquisa de orçamentos
    cbPagos.Visible = False
    
Else
    MSFlexGrid1.ColWidth(1) = 1000
    MSFlexGrid1.ColWidth(3) = 2000
    
    '3.6.9 Filtro pagos ou nao pagos na pesquisa de orçamentos
    cbPagos.ListIndex = 0
    
End If
MSFlexGrid1.ColWidth(2) = 800
MSFlexGrid1.ColWidth(4) = 1000
MSFlexGrid1.TextMatrix(0, 0) = "Data"
MSFlexGrid1.TextMatrix(0, 1) = "Valor"
MSFlexGrid1.TextMatrix(0, 2) = "Pago"
MSFlexGrid1.TextMatrix(0, 3) = "Carro"
MSFlexGrid1.TextMatrix(0, 4) = "Placa"
MSFlexGrid1.ColAlignment(2) = vbCenter

'3.6.9 Filtro pagos ou nao pagos na pesquisa de orçamentos
Mudando = False
End Sub

Public Property Get nmCliente()
nmCliente = GNmCliente
End Property

Public Property Let nmCliente(vNewValue)
Dim SQL$
'Dim NmCliente As String
Dim NrCli%

GNmCliente = vNewValue

NrCli% = Consulta("Select NrCli From Clientes Where Nome = " & GNmCliente)

Orcas.Caption = Orcas.Caption + GNmCliente

'2.7.3 Histórico do Orçamento
'Versçao 2.7.3
SQL$ = SelCamppos()

'2.8.7 Conserto do erro da pesquisa pelo orçamento pela placa
GNmCliente = Replace(GNmCliente, ",", "")

'5.0.6 Troca do SQL de seleção da pesquisa de orçamentos para resolver problema de duplicidade
SQL$ = SQL$ & "FROM Orcamento "
SQL$ = SQL$ & "INNER JOIN (Carros INNER JOIN Clientes ON Carros.NrCli = Clientes.NrCli) "
SQL$ = SQL$ & "ON (Carros.Placa = Orcamento.Carro) "
SQL$ = SQL$ & "AND (Orcamento.Cliente = Clientes.Nome) "
SQL$ = SQL$ & "WHERE Clientes.NrCli = " & NrCli%

'5.0.8 Ordenação da pesquisa de orçamentos por cliente, pela data
SQL$ = SQL$ & " Order by Orcamento.Data desc "

'SQL$ = SQL$ & "from Orcamento, Clientes, Carros "
'SQL$ = SQL$ & "Where Clientes.NrCli = " & NrCli%
'SQL$ = SQL$ & " and Orcamento.Cliente = Clientes.Nome"
'SQL$ = SQL$ & " and Carros.Placa =  Orcamento.Carro"

'2.0.5 Implantação do Log
AbreTB OrcasRecordset, SQL$, dbOpenSnapshot
'Set OrcasRecordset = db.OpenRecordset(Selec)

'3.1.1 Nao mostrar a pesquisa pela placa na pesquisa de orcamentos quando esta filtrado pelo nome
txPlaca.Visible = False
Label1(2).Visible = False

MostraOrcs
End Property

Private Sub MostraOrcs()
Dim a As Integer
Dim b As Integer

'2.7.2 Impedir duplicações na pesquisa de orçamentos
Dim ConfATU As Variant
Dim ConfANT As Variant

If OrcasRecordset.EOF Then
    
    '3.1.7 Impedir de dar erro caso a pesquisa de orçamentos não de resultado
    Encontrou = False

    '3.1.7 Melhorar a mensagem quando a pesquisa não da resultado, na pesquisa de orçamentos
    If txPlaca.Text > "" Then
        '3.1.1 Previsao para placa nao encontrada na pesquisa da placa na lista de orcamentos
        MsgBox "Placa nao localizada"
    Else
        MsgBox "Orçamento nao localizada"
    End If
Else

    '3.1.7 Impedir de dar erro caso a pesquisa de orçamentos não de resultado
    Encontrou = True
    
    OrcasRecordset.MoveLast
    OrcasRecordset.MoveFirst
    
    MSFlexGrid1.Rows = OrcasRecordset.RecordCount + 1
    a = 1
    Do While OrcasRecordset.EOF = False
    
        '2.7.2 Impedir duplicações na pesquisa de orçamentos
        ConfATU = Empty
        For b = 0 To 9
            ConfATU = ConfATU & OrcasRecordset(b)
        Next
        If ConfATU <> ConfANT Then
            ConfANT = ConfATU
                       
            MSFlexGrid1.TextMatrix(a, 0) = Format(OrcasRecordset!Data, "dd/mm/yyyy")
            MSFlexGrid1.TextMatrix(a, 1) = Format(OrcasRecordset!Total, "###,###.00")
            
            If OrcasRecordset!Pagamento > 0 Then
                            
                '3.8.7 Permitir lançamentos em orçamentos pagos
                If OrcasRecordset!VlrPago < OrcasRecordset!Total Then
                    MSFlexGrid1.TextMatrix(a, 2) = "orc.parc."
                Else
                    MSFlexGrid1.TextMatrix(a, 2) = "PAGO"
                End If
                
            Else
                MSFlexGrid1.TextMatrix(a, 2) = " "
            End If
            'MSFlexGrid1.TextMatrix(a, 2) = OrcasRecordset!Kilom
            
            MSFlexGrid1.TextMatrix(a, 3) = OrcasRecordset!Modelo
            MSFlexGrid1.TextMatrix(a, 4) = OrcasRecordset!Placa
                            
            '2.8.4 Otimizar a Obtenção do Nr do Cliente
            MSFlexGrid1.TextMatrix(a, 8) = clsCLi.NrDoNome(OrcasRecordset!Cliente)
            'MSFlexGrid1.TextMatrix(a, 8) = Consulta("Select NrCli From Clientes Where Nome = '" & OrcasRecordset!Cliente & "'")
                
            '2.4.2 Observação do orçamento na relação mensal
            MSFlexGrid1.TextMatrix(a, 5) = OrcasRecordset!Obs
        
            '2.4.6 Observações dos clientes na listagem de orçamentos e gravação das observações
            MSFlexGrid1.TextMatrix(a, 6) = SN(OrcasRecordset!Observacao, vbString)
            
            '2.6.0 Observação do mecânico
            MSFlexGrid1.TextMatrix(a, 9) = SN(OrcasRecordset!ObsMec, vbString)
            
            '2.7.3 Histórico do Orçamento
            MSFlexGrid1.TextMatrix(a, 10) = SN(OrcasRecordset!Historico, vbString)
                
            MSFlexGrid1.TextMatrix(a, 7) = OrcasRecordset!Orcamento
            
            '3.4.5 Impedir que funcionario edite seus orçamentos
            MSFlexGrid1.TextMatrix(a, 11) = Trim(STR(SN(OrcasRecordset!Funcionario, vbInteger)))
                            
            a = a + 1
        End If
        OrcasRecordset.MoveNext
    Loop
    
    '2.6.0 Observação do mecânico
    txObs(0).Text = MSFlexGrid1.TextMatrix(1, 5)
    txObs(1).Text = MSFlexGrid1.TextMatrix(1, 9)
    
    '2.7.3 Histórico do Orçamento
    txObs(2).Text = MSFlexGrid1.TextMatrix(1, 10)
End If

'2.4.2 Observação do orçamento na relação mensal
'txObs.Text = MSFlexGrid1.TextMatrix(1, 5)
End Sub

Public Property Let Mensal(ByVal vNewValue As String)
'1.8.0 Relação Mensal
Dim SQL As String

GMensal = vNewValue
DT = vNewValue

'2.7.3 Histórico do Orçamento
SQL = SelCamppos()

SQL = SQL & "FROM Orcamento, Carros, Clientes "
SQL = SQL & "WHERE Orcamento.Carro = Carros.Placa and Orcamento.Data > " & vNewValue
SQL = SQL & " and Clientes.Nome = Orcamento.Cliente"

'3.1.6 Inversão da ordenação na pesquisa de orçamentos
SQL = SQL & " Order By Orcamento.Data desc "

'2.0.5 Implantação do Log
AbreTB OrcasRecordset, SQL, dbOpenDynaset
'Set OrcasRecordset = db.OpenRecordset(Selec)

MostraOrcs
End Property

Public Property Get Mensal() As String
'1.8.0 Relação Mensal
Mensal = GMensal
End Property

Private Sub MSFlexGrid1_Click()
'2.4.2 Observação do orçamento na relação mensal
MostrarObs
End Sub

Private Sub MSFlexGrid1_DblClick()
'4.7.1 Retorno da permissão do mecânico ver o orçamento
'4.7.0 Retorno do menu para orçamentos mas sem acionar o orçamento
'If INI.ModoOperacao <> tpMecanico Then
    Aciona
'End If
End Sub

Private Sub MSFlexGrid1_EnterCell()
'3.6.2 Permitir atualizar o histórico pela pesquisa de orçamentos
If Mudando = False Then

    '2.4.2 Observação do orçamento na relação mensal
    MostrarObs
End If
End Sub

Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
'Shift Delete

'4.7.0 Retorno do menu para orçamentos mas sem acionar o orçamento
If INI.ModoOperacao <> tpMecanico Then

    If KeyCode = 46 And Shift = 2 Then
        '2.7.2 Logar todas mensagens
        If msgboxL("Deseja realmente excluir este orçamento?", vbQuestion + vbDefaultButton2 + vbYesNo, "Exclusão de orçamentos") = vbYes Then
            ExclusaoOrcamento
        End If
    End If
End If
End Sub

Private Sub RestauraBotoes()
btGrvOrc.Enabled = False
btGrvCliente.Enabled = False
Command1(0).Default = True
End Sub

Private Sub MostrarObs()
'2.6.0 Observação do mecânico
txObs(0).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
txObs(1).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 9)

'2.7.3 Histórico do Orçamento
txObs(2).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10)

'2.4.2 Observação do orçamento na relação mensal
'txObs.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
txObsCli.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)
RestauraBotoes
End Sub

Private Sub ExclusaoOrcamento()
Dim a As Integer

OrcasRecordset.MoveFirst
a = 1
Do While a < MSFlexGrid1.Row
    OrcasRecordset.MoveNext
    a = a + 1
Loop
a = OrcasRecordset("Orcamento")

If Consulta("Select Count(*) From Tarefas Where Orc = " & a & " and Situacao = 3 ") > 0 Then

    '3.2.6 Dar mensagem de alerta em caso de comando de apagar orçamento
    If msgboxL("Há taréfas concluídas para este orçamento ", vbCritical + vbYesNo + vbDefaultButton2, "Tem certeza que quer apagar?") = vbNo Then
        Exit Sub
    End If
End If

'2.0.5 Implantação do Log
ExecSql "Delete From Itens_Orc Where Orçamento = " & a
ExecSql "Delete From Orcamento Where Orcamento = " & a
'db.Execute "Delete From Itens_Orc Where Orçamento = " & a
'db.Execute "Delete From Orcamento Where Orcamento = " & a

'3.2.6 Apagar as tarefas quando for apagado o orçamento
ExecSql "Delete From Tarefas Where Orc = " & a
End Sub

Private Sub opObs_Click(Index As Integer)
'2.6.0 Observação do mecânico
If Mudando = False Then
    txObs(0).Visible = (Index = 0)
    txObs(1).Visible = (Index = 1)
    
    '2.7.3 Histórico do Orçamento
    txObs(2).Visible = (Index = 2)
    
    opObs(0).Value = IIf(Index = 0, 1, 0)
    opObs(1).Value = IIf(Index = 1, 1, 0)
    opObs(2).Value = IIf(Index = 2, 1, 0)
    
End If
End Sub

Private Sub txObsCli_KeyPress(KeyAscii As Integer)
'3.1.0 Modo Balcão
If INI.ModoOperacao <> tpMecanico Then
'2.5.6 Modo Restrito
'If INI.Restrito = False Then

    '2.4.6 Observações dos clientes na listagem de orçamentos e gravação das observações
    btGrvCliente.Enabled = True
    Command1(0).Default = False
End If
End Sub

Private Sub txObs_KeyPress(Index As Integer, KeyAscii As Integer)
'2.6.0 Observação do mecânico
If Index Then
    '2.6.1 Permitir escritório escrever na observação do cliente
    'If INI.Restrito Then
        btGrvOrc.Enabled = True
        Command1(0).Default = False
    'End If
Else

    '3.1.0 Modo Balcão
    If INI.ModoOperacao <> tpMecanico Then
    '2.5.6 Modo Restrito
    'If INI.Restrito = False Then
    
        '2.4.6 Observações dos clientes na listagem de orçamentos e gravação das observações
        btGrvOrc.Enabled = True
        Command1(0).Default = False
    End If
End If
End Sub

Public Property Let Placa(ByVal vNewValue As Variant)
'3.7.8 Troca de carros entre os clientes
Dim SQL$

Orcas.Caption = Orcas.Caption & "Placa: " & vNewValue

SQL$ = SelCamppos()

SQL$ = SQL$ & "from Carros, Orcamento, Clientes "
SQL$ = SQL$ & "Where Carros.Placa = " & vNewValue
SQL$ = SQL$ & " and Orcamento.Carro = Carros.Placa"
SQL$ = SQL$ & " and Clientes.Nome = Orcamento.Cliente"

AbreTB OrcasRecordset, SQL$, dbOpenSnapshot

txPlaca.Visible = False
Label1(2).Visible = False

MostraOrcs
End Property
