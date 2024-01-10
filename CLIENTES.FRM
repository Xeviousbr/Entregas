VERSION 5.00
Object = "{87E4A023-7D46-4600-80F0-1B2287A98461}#3.0#0"; "datacontrol.ocx"
Begin VB.Form frmClientes 
   Caption         =   "Clientes"
   ClientHeight    =   6330
   ClientLeft      =   2490
   ClientTop       =   1680
   ClientWidth     =   8010
   Icon            =   "CLIENTES.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6330
   ScaleWidth      =   8010
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Nada"
   Begin VB.Frame Frame3 
      Caption         =   "Revisões"
      Height          =   2175
      Left            =   4380
      TabIndex        =   46
      Top             =   3960
      Width           =   3495
      Begin VB.TextBox txRevNome 
         Height          =   315
         Index           =   0
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox ckRev 
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   60
         Top             =   300
         Width           =   255
      End
      Begin VB.TextBox txRevDt 
         Height          =   315
         Index           =   0
         Left            =   2340
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txRevNome 
         Height          =   315
         Index           =   1
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox ckRev 
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   57
         Top             =   660
         Width           =   255
      End
      Begin VB.TextBox txRevDt 
         Height          =   315
         Index           =   1
         Left            =   2340
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txRevNome 
         Height          =   315
         Index           =   2
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox ckRev 
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   54
         Top             =   1020
         Width           =   255
      End
      Begin VB.TextBox txRevDt 
         Height          =   315
         Index           =   2
         Left            =   2340
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txRevNome 
         Height          =   315
         Index           =   3
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox ckRev 
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   51
         Top             =   1380
         Width           =   255
      End
      Begin VB.TextBox txRevDt 
         Height          =   315
         Index           =   3
         Left            =   2340
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txRevNome 
         Height          =   315
         Index           =   4
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CheckBox ckRev 
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   48
         Top             =   1740
         Width           =   255
      End
      Begin VB.TextBox txRevDt 
         Height          =   315
         Index           =   4
         Left            =   2340
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1680
         Width           =   975
      End
   End
   Begin VB.CommandButton btRecibo 
      Caption         =   "&Recibo"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6120
      TabIndex        =   45
      Top             =   3540
      Width           =   1455
   End
   Begin VB.CommandButton btValores 
      Caption         =   "&Valores"
      Height          =   315
      Left            =   6120
      TabIndex        =   39
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton btObs 
      Caption         =   "&Observação"
      Height          =   315
      Left            =   6120
      TabIndex        =   37
      Top             =   2220
      Width           =   1455
   End
   Begin VB.CommandButton btNvOrcam 
      Caption         =   "&Novo Orçamento"
      Height          =   315
      Left            =   6120
      TabIndex        =   36
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton btOrcamento 
      Caption         =   "&Orçamento"
      Height          =   315
      Left            =   6120
      TabIndex        =   35
      Top             =   900
      Width           =   1455
   End
   Begin VB.CommandButton BtImprCabec 
      Caption         =   "&Imprimir Ficha"
      Height          =   315
      Left            =   6120
      TabIndex        =   33
      Top             =   240
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   3735
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   5655
      Begin VB.CheckBox ckAtraso 
         Alignment       =   1  'Right Justify
         Caption         =   "Em Atraso"
         Height          =   375
         Left            =   3900
         TabIndex        =   44
         ToolTipText     =   "Indica que o cliente esta em atraso"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Ender"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   9
         Left            =   1620
         MaxLength       =   40
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox ckOper 
         Alignment       =   1  'Right Justify
         Caption         =   "Tim"
         Height          =   315
         Index           =   3
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Indica que o cliente é funcionário"
         Top             =   1500
         Width           =   495
      End
      Begin VB.CheckBox ckOper 
         Alignment       =   1  'Right Justify
         Caption         =   "Claro"
         Height          =   315
         Index           =   2
         Left            =   2820
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Indica que o cliente é funcionário"
         Top             =   1500
         Width           =   495
      End
      Begin VB.CheckBox ckOper 
         Alignment       =   1  'Right Justify
         Caption         =   "Vivo"
         Height          =   315
         Index           =   1
         Left            =   2220
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Indica que o cliente é funcionário"
         Top             =   1500
         Width           =   495
      End
      Begin VB.CheckBox ckOper 
         Alignment       =   1  'Right Justify
         Caption         =   "Oi"
         Height          =   315
         Index           =   0
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Indica que o cliente é funcionário"
         Top             =   1500
         Width           =   495
      End
      Begin VB.TextBox txtFields 
         DataField       =   "CgcCpf"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   8
         Left            =   1620
         MaxLength       =   19
         TabIndex        =   13
         Top             =   2460
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "CgcCpf"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   7
         Left            =   1620
         MaxLength       =   19
         TabIndex        =   12
         Top             =   2145
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Telefone"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   0
         Left            =   1620
         TabIndex        =   11
         Top             =   1830
         Width           =   3495
      End
      Begin VB.CheckBox ckFunc 
         Alignment       =   1  'Right Justify
         Caption         =   "Funcionário"
         Height          =   375
         Left            =   3900
         TabIndex        =   5
         ToolTipText     =   "Indica que o cliente é funcionário"
         Top             =   765
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   495
         Left            =   3840
         TabIndex        =   18
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Gravar"
         Height          =   300
         Left            =   2280
         TabIndex        =   16
         Top             =   2820
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Apaga&r"
         Height          =   300
         Left            =   1200
         TabIndex        =   15
         Top             =   2820
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Adicionar"
         Height          =   300
         Left            =   120
         TabIndex        =   14
         Top             =   2820
         Width           =   975
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Telefone"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   3
         Left            =   1620
         TabIndex        =   6
         Top             =   1185
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Ender"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   2
         Left            =   1620
         MaxLength       =   40
         TabIndex        =   3
         Top             =   512
         Width           =   3495
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Nome"
         DataSource      =   "Data1"
         Height          =   285
         HideSelection   =   0   'False
         Index           =   1
         Left            =   1620
         MaxLength       =   40
         TabIndex        =   1
         Top             =   195
         Width           =   3495
      End
      Begin PrjDataControl.DataControl DataControl1 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   556
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Cep:"
         Height          =   255
         Index           =   10
         Left            =   480
         TabIndex        =   43
         Top             =   923
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Inscr.Estadual:"
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   42
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Operadora:"
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   41
         Top             =   1561
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "CPF ou CNPj:"
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   40
         Top             =   2199
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Email:"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   20
         Top             =   1880
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Telefone:"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   19
         Top             =   1242
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Endereço:"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   2
         Top             =   604
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Nome:"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   0
         Top             =   285
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Carros"
      Height          =   2175
      Left            =   120
      TabIndex        =   31
      Top             =   3960
      Width           =   4155
      Begin VB.CommandButton btTroca 
         Caption         =   "&Troca"
         Height          =   300
         Left            =   3180
         TabIndex        =   38
         Top             =   1300
         Width           =   915
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Gravar"
         Height          =   300
         Left            =   2160
         TabIndex        =   29
         Top             =   1300
         Width           =   915
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Apaga&r"
         Height          =   300
         Left            =   1080
         TabIndex        =   28
         Top             =   1300
         Width           =   915
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Adicionar"
         Height          =   300
         Left            =   60
         TabIndex        =   27
         Top             =   1300
         Width           =   915
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Cor"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   840
         TabIndex        =   26
         Top             =   870
         Width           =   2415
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Placa"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   840
         TabIndex        =   24
         Top             =   540
         Width           =   2415
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Modelo"
         DataSource      =   "Data2"
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   840
         MaxLength       =   36
         TabIndex        =   22
         Top             =   225
         Width           =   3195
      End
      Begin PrjDataControl.DataControl DataControl2 
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   1740
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   556
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Cor:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   25
         Top             =   885
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Placa:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   570
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "&Modelo:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Close"
      Height          =   300
      Left            =   9840
      TabIndex        =   30
      Top             =   975
      Width           =   975
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'5.0.9 Impedir a adição de um carro com uma placa de outro carro
'5.0.4 Cadastro de Revisões
'4.7.6 Pesquisar por cliente ou funcionário no recibo avulso
'4.5.4 Cliente em atraso
'4.2.9 Ajuste na gravação da operadora do celular
'4.2.8 CEP
'4.2.7 Operadora do celular
'4.2.7 Inscrição Estatual
'4.2.5 Chamar o gerente caso não tenha CPNj ou CPF
'4.1.4 Campo para CGC e CPF no cadastro de clientes
'4.0.2 Identificar quem editou o item do orçamento
'4.0.1 Conserto da adição do cliente, apos a implantação do email
'3.8.4 Adaptação da resolução
'3.8.1 Mostra quando o carro ta com um orçamento pendente
'3.8.1 Conserto do sumiço do cliente 09/07/2014 na troca
'3.8.1 Fecha a tela de clientes depois de trocar o carro
'3.7.9 Permitir a troca pelo balcao
'3.7.8 Troca de carros entre os clientes
'3.6.4 Não permitir gravar um cadastro com o nome de outro cliente
'3.5.1 Não excluir fisicamente Mecânico
'3.4.4 Cliente Funcionario
'3.3.1 Conserto da edição do carro
'3.2.8 Conserto da edição do segundo carro
'3.1.6 Impedir que os mecânicos vejam orçamentos atravez da tela de clientes
'3.1.2 Gravar o operador que fizer um cadastro novo de cliente
'3.1.0 Modo Balcão
'3.0.8 Retirada do botão de transferência do carro
'2.9.0 Show in TaskBar true para tela de clientes
'2.8.7 Não mostrar a obs do cliente no modo restrito
'2.7.2 Logar todas mensagens
'2.7.2 Conserto da exclusão do Carro
'2.6.1 Mecânicos não devem ver valores
'2.5.6 Modo Restrito
'2.4.6 Conserto da impressão da cor na impressão da ficha em modo fita
'2.4.6 Impedir que saia a ultima impressão feita, quando se faz a impressão da ficha do cliente em modo fita
'2.4.1 Impressão da ficha do cliente no formato fita
'2.4.0 Campo para observação do cliente
'2.3.3 Possibilitar retornar o carro ao cliente
'2.3.1 Previsão para o bug que não aparecia o carro certo quando se acionava um novo orçamento via cadastro de clientes
'2.2.7 Permitir editar o modelo do carro ao pesquisar pelo cliente
'2.2.2 Impedir o erro que impedia de alterar um cadastro de um cliente recem cadastrado
'2.1.1 Mostrar o carro escolhido já na tela do cliente
'2.1.1 Conserto da alteração do nome do cliente
'2.0.5 Implantação do Log
'2.0.5 Impedir a duplicidade de clientes
'2.0.3 Ajuste no funcionamento das permissões

Option Explicit

Private Adicionando  As Boolean
Private HouveAdicao  As Boolean
Private AdicCarro    As Boolean
Private AdicionouCar As Boolean
Private Data2_Aberto As Boolean
Private VarNrCliente As Integer
Private NrCarros     As Integer
Private NrDesteCarr  As Integer
Private NomeOrig     As String
Private Data1        As Recordset
Private Data2        As Recordset

'3.4.4 Cliente Funcionario
Private l_Func As Boolean

'4.2.8 CEP
Private Const ConstProp = 0.64
'3.8.4 Adaptação da resolução
'Private Const ConstProp = 0.4972222

Private tamEmail As Integer

'4.2.7 Operadora do celular
Private Mudando As Boolean

'5.0.4 Cadastro de Revisões
Private cRevis As clsRevisoes
Private Carregando As Boolean

Private Function LocalizaCarro() As Integer
Dim a        As Integer
Dim TbCarros As Recordset

'2.0.5 Implantação do Log
AbreTB TbCarros, "Select Placa From Carros Where NrCli = " & VarNrCliente
'Set TbCarros = db.OpenRecordset("Select Placa From Carros Where NrCli = " & VarNrCliente)

Do While TbCarros.EOF = False
    a = a + 1
    If TbCarros(0) = txtFields(5).Text Then
        LocalizaCarro = a
        Exit Function
    End If
    TbCarros.MoveNext
Loop
LocalizaCarro = 0
End Function

Private Sub GravaCarro()
Dim Posic    As Integer
Dim SQL      As String
Dim PlacaAnt As String

If AdicCarro = True Then
    SQL = "Insert Into Carros (NrCli, Modelo, Placa, Cor) Values (" & VarNrCliente & ","
    SQL = SQL & "'" & txtFields(4).Text & "',"
    SQL = SQL & "'" & txtFields(5).Text & "',"
    SQL = SQL & "'" & txtFields(6).Text & "')"
    
    '2.0.5 Implantação do Log
    ExecSql SQL
    'db.Execute SQL
    
    Posic = LocalizaCarro()
    NrCarros = Consulta("Select Count(*) From Carros Where NrCli = " & VarNrCliente)
    If Posic > 0 Then
        'DataControl2.Label = "Carro: " & NrDesteCarr & " de " & str$(NrCarros)
        DataControl2.Label = "Carro: " & Posic & " de " & STR$(NrCarros)
    End If
    'Data2_Reposition
    AdicCarro = False
    AdicionouCar = True
    NrDesteCarr = NrDesteCarr + 1
    If NrDesteCarr > NrCarros Then
        NrDesteCarr = NrCarros
    End If
Else
    
    '3.3.1 Conserto da edição do carro
'    SQL = "Select Placa From Carros Where Modelo = '" & txtFields(4).Text
'    SQL = SQL & "' and Cor = '" & txtFields(6).Text
'    SQL = SQL & "' and Placa <> '" & txtFields(5).Text
'    SQL = SQL & "' and NrCli = " & Data1(0)
    'PlacaAnt = Consulta(SQL)
    PlacaAnt = txtFields(5).Tag
    
    If PlacaAnt > "" Then
        SQL = "Select Count(*) From Orcamento Where Cliente = '" & txtFields(1) & "' and Carro = '" & PlacaAnt & "'"
        If Consulta(SQL) > 0 Then
            '2.7.2 Logar todas mensagens
            Select Case msgboxL("Deseja adaptar os Orcamentos", vbInformation + vbYesNoCancel + vbDefaultButton3, "Mudanca de placa em carro com Orcamento")
                Case vbCancel
                    Exit Sub
                Case vbYes
                    SQL = "Update Orcamento Set Carro = '" & txtFields(5).Text
                    SQL = SQL & "' Where Cliente = '" & txtFields(1) & "' and Carro = '" & PlacaAnt & "'"
                    ExecSql SQL
            End Select
        End If
    End If

    AdicionouCar = False
    SQL = "Update Carros Set Modelo = '" & txtFields(4).Text & "',"
    SQL = SQL & "Placa = '" & txtFields(5).Text & "',"
    SQL = SQL & "Cor = '" & txtFields(6).Text & "' "
    SQL = SQL & "Where NrCli = " & Data1(0)
    
    '3.3.1 Conserto da edição do carro
    SQL = SQL & " and Placa = '" & txtFields(5).Tag & "'"
    '3.2.8 Conserto da edição do segundo carro
    'SQL = SQL & " and Placa = '" & PlacaAnt & "'"
        
    '2.0.5 Implantação do Log
    ExecSql SQL
    'db.Execute SQL
    
End If
Command1.Enabled = True

'3.1.0 Modo Balcão
Command2.Enabled = (INI.ModoOperacao = tpEscritorio)
'Command2.Enabled = (INI.Restrito = False)

End Sub

Private Sub ImprimeFita()
'2.4.1 Impressão da ficha do cliente no formato fita
Dim a        As Integer
Dim sObs     As String
Dim SQL      As String
Dim TbCarros As Recordset

'2.4.6 Impedir que saia a ultima impressão feita, quando se faz a impressão da ficha do cliente em modo fita
ImprBuferizada_Inicializa

'4.2.1 Impressão da observação
'Const TamFita = 55

LPRINT CENTRAL("Cliente OrCarro", TamFita / 2)
LPRINT "Nome    : " & txtFields(1).Text
LPRINT "Endereço: " & txtFields(2).Text
LPRINT "Telefone: " & txtFields(3).Text

'4.1.4 Campo para CGC e CPF no cadastro de clientes
LPRINT "CNPj/CPF: " & txtFields(7).Text

If NrCarros > 0 Then
    LPRINT " "
    If NrCarros = 1 Then
        LPRINT "Proprietário de 1 carro"
        LPRINT "Marca: " & SemAcento(txtFields(4).Text)
        
        '2.4.6 Conserto da impressão da cor na impressão da ficha em modo fita
        LPRINT "Cor  : " & SemAcento(txtFields(5).Text)
        'LPRINT "Cor  : " & SemAcento(txtFields(5).Text)
        
        LPRINT "Placa: " & txtFields(5).Text
    Else
        SQL = "Select Modelo, Placa, Cor From Carros "
        SQL = SQL & " Where NrCli = " & clsCLi.NrCli
        SQL = SQL & " Order By Modelo, Cor, Placa "
        AbreTB TbCarros, SQL, dbOpenSnapshot
        LPRINT "Proprietário de" & NrCarros & " carros"
        Do While TbCarros.EOF = False
            LPRINT "Carro" & a
            LPRINT "Marca: " & SemAcento(TbCarros("Modelo"))
            LPRINT "Cor  : " & SemAcento(TbCarros("Cor"))
            LPRINT "Placa: " & TbCarros("Placa")
            TbCarros.MoveNext
        Loop
    End If
End If

sObs = Consulta("Select Observacao From Clientes Where NrCli = " & clsCLi.NrCli)
If sObs > "" Then
    LPRINT "Observação: "
    ImprPedaco sObs, TamFita
End If
'If INI.LinhasApos > 0 Then
'    For a = 0 To INI.LinhasApos
'        LPRINT " "
'    Next
'End If
ImprBuferizada_Finaliza
End Sub

Private Sub ImprPedaco(Texto As String, Tam As Integer)
'2.4.1 Impressão da ficha do cliente no formato fita
Dim Sair     As Boolean
Dim TamSep   As Integer
Dim TemEnter As Integer
Dim Inicio   As String
Dim Resto    As String

Resto = Texto
TamSep = InStr(Resto, vbCrLf)
Do While Sair = False
    If TamSep > 0 Then
        TemEnter = 3
        TamSep = TamSep - 1
    Else
        TemEnter = 0
        If TamSep > Tam Then
            TamSep = Tam
        Else
            If Len(Resto) Then
                TamSep = Len(Resto)
            Else
                Exit Do
            End If
            Sair = True
        End If
    End If
    Inicio = Left(Resto, TamSep)
    Resto = Mid(Resto, TamSep + TemEnter)
    LPRINT Inicio
    TamSep = InStr(Resto, vbCrLf)
Loop
End Sub

Private Sub ImprimeFolha()
'2.4.1 Impressão da ficha do cliente no formato fita
Dim MargemH As Integer
Dim MargemV As Integer
Dim posY    As Integer
Dim a       As Integer
Dim MaObra  As Currency
Dim Cont%
Dim sObs    As String

Loga "Imprimir Cabeçario"
MargemV = 9
MargemH = 15

Printer.ScaleMode = 6   'Milimetros
Printer.Font = "Arial Narrow"
Printer.CurrentY = MargemV + 0
Printer.CurrentX = MargemH + 75
Printer.FontSize = 19
Printer.FontBold = True
Printer.Print "CLIENTE"

Printer.FontSize = 16
Printer.FontBold = False

Printer.CurrentY = MargemV + 8
Printer.CurrentX = MargemH + 25
Printer.Print "Porto Alegre, " + Format$(Now, "dd") + " de " + Format$(Now, "mmmm") + " de " + Format$(Now, "yyyy")

Printer.CurrentY = MargemV + 19
Printer.CurrentX = MargemH + 0
Printer.Print "Proprietário do Veículo: " + txtFields(1).Text

Printer.CurrentX = MargemH + 0
Printer.Print "Endereço: " + txtFields(2).Text

Printer.CurrentX = MargemH + 0
Printer.Print "Veículo Marca: " + txtFields(4).Text + " " + txtFields(6).Text

Printer.CurrentX = MargemH + 0
Printer.Print "Placa " + txtFields(5).Text

'2.4.0 Campo para observação do cliente
sObs = Consulta("Select Observacao From Clientes Where NrCli = " & clsCLi.NrCli)
If sObs > "" Then
    Printer.CurrentX = MargemH + 0
    Printer.Print "Observação: " + sObs
End If

Printer.EndDoc
End Sub

Private Sub BtImprCabec_Click()
'2.4.1 Impressão da ficha do cliente no formato fita
If INI.TpImpress = 0 Then
    ImprimeFolha
Else
    ImprimeFita
End If
End Sub

Private Sub btNvOrcam_Click()
Loga "Novo Orçamento pela tela de clientes " & txtFields(1).Text
clsCLi.NrCli = Data1("NrCli")

'2.7.5 Taréfas Dinâmicas
NovasTarefas

Load frmOrc
frmOrc.Show

'2.1.1 Mostrar o carro escolhido já na tela do cliente
'2.3.1 Previsão para o bug que não aparecia o carro certo quando se acionava um novo orçamento via cadastro de clientes
frmOrc.NovoOrcDoCli txtFields(1).Text, txtFields(5).Text, VarNrCliente, txtFields(2).Text, txtFields(4).Text
frmOrc.Timer.Enabled = True

Unload Me
End Sub

Private Sub btObs_Click()
'2.4.0 Campo para observação do cliente
frmObs.Show
End Sub

Private Sub btOrcamento_Click()
Dim Pesq As String
Dim SQL  As String

'2.8.4 Gravar no log o tempo de carregamento da tela de orçamento
MomChamOrc = Now

Loga "Chamar Orçamento pela tela de clientes"
clsCLi.NrCli = Data1("NrCli")
Pesq = Chr$(34) + txtFields(1).Text + Chr$(34)
SQL = "SELECT DISTINCTROW Cliente, Data, Orcamento "
SQL = SQL & "From Orcamento "
SQL = SQL & "WHERE ((Orcamento.Cliente=" + Pesq + ")); "
AchaOrc SQL, Pesq
End Sub

'4.7.6 Pesquisar por cliente ou funcionário no recibo avulso
Private Sub btRecibo_Click()
Load frmReciboAvulso
frmReciboAvulso.txRecebe.Text = txtFields(1).Text
frmReciboAvulso.txIdent = txtFields(7).Text
frmReciboAvulso.txEndereco = txtFields(2).Text
frmReciboAvulso.Show
End Sub

Private Sub btTroca_Click()
Dim TaPago As Integer
Dim sS     As String

'3.8.1 Mostra quando o carro ta com um orçamento pendente
TaPago = Consulta("Select Count(*) From Orcamento Where Cliente = '" & txtFields(1).Text & "' and Pagamento < 1  and Carro = '" & txtFields(5).Text & "'")
If TaPago > 1 Then
    sS = "Há orçamentos que não foram pagos"
Else
    If TaPago > 0 Then
        sS = "Existe um orçamento que não foi totalmente pago"
    End If
End If
If sS > "" Then
    If MsgBox("Deseja realizar a troca mesmo assim?", vbExclamation + vbYesNo, sS) = vbNo Then
        Exit Sub
    End If
End If

'3.7.8 Troca de carros entre os clientes
If txtFields(4).Text > "" Then
    Load TrocaCarro
    TrocaCarro.Label1(1).Caption = TrocaCarro.Label1(1).Caption & txtFields(1).Text
    TrocaCarro.lblLabels(4).Caption = TrocaCarro.lblLabels(4).Caption & txtFields(4).Text
    TrocaCarro.lblLabels(5).Caption = TrocaCarro.lblLabels(5).Caption & txtFields(5).Text
    TrocaCarro.lblLabels(6).Caption = TrocaCarro.lblLabels(6).Caption & txtFields(6).Text
    TrocaCarro.nmCliente = txtFields(1).Text
    TrocaCarro.Placa = txtFields(5).Text
    clsCLi.CarCliCB TrocaCarro.cbCliente, txtFields(1).Text
    
    '3.8.1 Fecha a tela de clientes depois de trocar o carro
    TrocaCarro.Show 1
    Unload Me
    'TrocaCarro.Show
        
End If
End Sub

Private Sub btValores_Click()
frmValores.Show
End Sub

'4.2.7 Operadora do celular
Private Sub ckOper_Click(Index As Integer)
Dim a As Integer

If Mudando = False Then
    Mudando = True
    For a = 0 To 3
        If a <> Index Then
            ckOper(a).Value = 0
        End If
    Next
    Mudando = False
End If
End Sub

Private Sub ckRev_Click(Index As Integer)
'5.0.4 Cadastro de Revisões
If Carregando = False Then
    cRevis.ckRev_Click Index
End If
End Sub

Private Sub cmdAdd_Click()
'2.0.3 Ajuste no funcionamento das permissões
If Data1.RecordCount > 100 And Permissao = False Then
'    Data1.CancelUpdate
    '2.7.2 Logar todas mensagens
    msgboxL "Na versão demonstração só é permitido 100 registros"
    cmdAdd.Enabled = False
Else
    Loga "Adicionar Cliente"
    DataControl1.Label = ""
    If txtFields(1).Enabled = False Then
       txtFields(1).Enabled = True
       txtFields(2).Enabled = True
       txtFields(3).Enabled = True
       txtFields(0).Enabled = True
       
       '4.1.4 Campo para CGC e CPF no cadastro de clientes
       txtFields(7).Enabled = True
              
       '4.2.7 Inscrição Estatual
       txtFields(8).Enabled = True
       
       '4.2.8 CEP
       txtFields(9).Enabled = True
       
    End If
    txtFields(1).Text = ""
    txtFields(2).Text = ""
    txtFields(3).Text = ""
    txtFields(0).Text = ""
    
    '4.1.4 Campo para CGC e CPF no cadastro de clientes
    txtFields(7).Text = ""
    
    '4.2.7 Inscrição Estatual
    txtFields(8).Text = ""
    
    '4.2.8 CEP
    txtFields(9).Text = ""
    
    '4.2.7 Operadora do celular
    Dim a As Integer
    For a = 0 To 3
        ckOper(a).Value = 0
    Next
    
    txtFields(1).SetFocus
    
    '3.8.4 Adaptação da resolução
    frmClientes.Height = frmClientes.Height * ConstProp
    'frmClientes.Height = 2685
    
    Adicionando = True
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    NrCarros = 0
        
    '3.1.0 Modo Balcão
'    '2.6.1 Mecânicos não devem ver valores
'    If INI.Restrito = False Then
'        btOrcamento.Enabled = False
'    End If
    
    btNvOrcam.Enabled = False
End If
End Sub

Private Sub cmdDelete_Click()
Dim NrAtu  As Integer
Dim ProxNr As Integer

'2.7.2 Logar todas mensagens
If msgboxL("Deseja realmente excluir o cliente?", vbQuestion + vbDefaultButton2 + vbYesNo, "Orcarro") = vbYes Then
    Loga "Deletar Cliente"
    
    NrAtu = Data1!NrCli
    Data1.Close
    ProxNr = Consulta("Select NrCli From Clientes Where Nome > '" & Trim(txtFields(1)) & "' Order By Nome")
    
    ExecSql "Delete From Clientes Where NrCli = " & NrAtu
    
    '2.0.5 Implantação do Log
    ExecSql "Delete From Carros Where NrCli = " & NrAtu
    'db.Execute "Delete From Carros Where NrCli = " & Data1!NrCli
    
    'Data1.Delete
    'Data1.MoveNext
    AbreTB Data1, "Select * From Clientes Where NrCli = " & ProxNr
    
    Mostra
End If
End Sub

Private Sub cmdUpdate_Click()
Dim Balconista As Integer
Dim Nome       As String
Dim SQL        As String

'4.0.2 Identificar quem editou o item do orçamento
Dim sNome As String

'4.2.5 Chamar o gerente caso não tenha CPNj ou CPF
Dim Result As Boolean

Loga "Gravar Cliente"
'2.0.3 Ajuste no funcionamento das permissões
If Data1.RecordCount > 100 And Permissao = False Then
    '2.7.2 Logar todas mensagens
    msgboxL "Na versão demonstração não é permitido gravar em registro acima do 100"
    cmdUpdate.Enabled = False
Else
    If Adicionando Then
    
        '2.0.5 Impedir a duplicidade de clientes
        If Consulta("Select Count(*) From Clientes Where Nome = '" & txtFields(1) & "'") > 0 Then
            '2.7.2 Logar todas mensagens
            msgboxL "Este cliente já esta cadastrado"
            
        Else
            'AdiCliente
            'VarNrCliente = Consulta("Select Max(NrCli) From Clientes")
            
            '4.2.5 Chamar o gerente caso não tenha CPNj ou CPF
            If txtFields(7).Text = "" Then
                frmSenha.Tipo = 1
                frmSenha.Caption = "Senha do Gerente"
                frmSenha.Show 1
                Result = frmSenha.Resultado
                Unload frmSenha
                If Result = False Then
                    Exit Sub
                End If
            Else
                If calculacpf(txtFields(7).Text) = False Then
                '4.1.4 Campo para CGC e CPF no cadastro de clientes
                'If CheckCGCCPF(txtFields(7).Text) = False Then
                
                    MsgBox "CNPj ou CPF inválido", vbExclamation, "Informe o CNPJ ou CPF corretamente"
                    Exit Sub
                End If
            End If
                                    
            '4.0.2 Identificar quem editou o item do orçamento
            If INI.UtComissoes = "1" Then
                
                '3.1.2 Gravar o operador que fizer um cadastro novo de cliente
                Balconista = OperadorQueGravou(sNome, Func)
                
            End If
            If Balconista = 0 And INI.UtComissoes = "1" Then
                Exit Sub
            Else
                '4.0.2 Identificar quem editou o item do orçamento
                txtFields(1).ToolTipText = sNome
            
                '3.1.2 Gravar o operador que fizer um cadastro novo de cliente
                VarNrCliente = AdiCliente(Balconista)
                'VarNrCliente = AdiCliente()
                
                Data1.Close
                
                '2.0.5 Implantação do Log
                SQL = "Select * From Clientes Where NrCli = " & VarNrCliente
                AbreTB Data1, SQL
                'Set Data1 = db.OpenRecordset("Select * From Clientes Where NrCli = " & VarNrCliente)
                
                DataControl1.Label = "Cliente: " & VarNrCliente
                        
                '3.8.4 Adaptação da resolução
                frmClientes.Height = frmClientes.Height * (1 / ConstProp)
                'frmClientes.Height = 4995
                
                cmdAdd.Enabled = True
                
                '3.1.0 Modo Balcão
                cmdDelete.Enabled = (INI.ModoOperacao = tpEscritorio)
                'cmdDelete.Enabled = (INI.Restrito = False)
                
                Command1_Click

            End If
        End If
    Else
    
        '3.6.4 Não permitir gravar um cadastro com o nome de outro cliente
        SQL = "Select Count(*) From Clientes Where Nome = '"
        SQL = SQL & txtFields(1).Text
        SQL = SQL & "' and NrCli <> " & Data1!NrCli
        If Consulta(SQL) > 0 Then
            MsgBox "Já existe um cadastro com este nome", vbExclamation, "OrCarro"
        Else
        
            '4.2.5 Chamar o gerente caso não tenha CPNj ou CPF
            If txtFields(7).Text = "" Then
                frmSenha.Tipo = 1
                frmSenha.Caption = "Senha do Gerente"
                frmSenha.Show 1
                Result = frmSenha.Resultado
                Unload frmSenha
                If Result = False Then
                    Exit Sub
                End If
                
            '4.2.7 Conserto da situação que quando não informava o CPF/CNPj não gravava
            Else
        
                If calculacpf(txtFields(7).Text) = False Then
                '4.1.4 Campo para CGC e CPF no cadastro de clientes
                'If CheckCGCCPF(txtFields(7).Text) = False Then
                
                    MsgBox "CNPj ou CPF inválido", vbExclamation, "Informe o CNPJ ou CPF corretamente"
                    Exit Sub
                End If
            End If
                
            '4.0.2 Identificar quem editou o item do orçamento
            Balconista = OperadorQueGravou(sNome, Func)
            '3.1.2 Gravar o operador que fizer um cadastro novo de cliente
            'Balconista = OperadorQueGravou(Func)
            
            If Balconista = 0 And INI.UtComissoes = "1" Then
                Exit Sub
            Else
                 '4.0.2 Identificar quem editou o item do orçamento
                 txtFields(1).ToolTipText = sNome
            
                 VarNrCliente = Data1!NrCli
                 VerificaTamanhoEmail txtFields(0).Text
            
                 SQL = "Update Clientes Set Ender = '" & txtFields(2).Text
                 SQL = SQL & "', Telefone = '" & txtFields(3).Text & "', "
                 SQL = SQL & " Nome = '" & txtFields(1).Text & "' "
                 
                 '3.1.2 Gravar o operador que fizer um cadastro novo de cliente
                 SQL = SQL & ", Vend = " & Balconista
                 
                 '3.4.4 Cliente Funcionario
                 SQL = SQL & ", Funcionario = " & ckFunc.Value
                 
                 SQL = SQL & ", email = " & FA(txtFields(0).Text)
                 
                 '4.1.4 Campo para CGC e CPF no cadastro de clientes
                 SQL = SQL & ", CgcCpf = " & FA(txtFields(7).Text)
                 
                 '4.2.7 Inscrição Estatual
                 SQL = SQL & ", InscrEst = " & FA(txtFields(8).Text)
                 
                 '4.2.8 CEP
                 SQL = SQL & ", CEP = " & FA(txtFields(9).Text)
                 
                 '4.2.7 Operadora do celular
                 SQL = SQL & ", Operadora = " & GetOperadora()
                 
                 '4.5.4 Cliente em atraso
                 SQL = SQL & ", EmAtraso = " & ckAtraso.Value
                                  
                 SQL = SQL & " Where NrCli = " & VarNrCliente
                 
                 '2.0.5 Implantação do Log
                 ExecSql SQL
                 'db.Execute SQL
                 
                 '3.4.4 Cliente Funcionario
                 Func = ckFunc.Value
                 
                '1.7.2 Alterar o nome do cliente no orçamento quando muda no cadastro de clientes
                If txtFields(1).Text <> NomeOrig Then
                   SQL = "UPDATE Orcamento SET Cliente = '" + txtFields(1).Text + "' WHERE Cliente='" + NomeOrig + "'"
                   
                   '2.0.5 Implantação do Log
                   ExecSql SQL
                   'db.Execute SQL
                   
                   '2.1.1 Conserto da alteração do nome do cliente
                   SQL = "Update Orcamento Set Cliente = '" + txtFields(1).Text + "' WHERE Cliente='" + NomeOrig + "'"
                   ExecSql SQL
                   
                End If
            End If
        End If
    End If
End If
End Sub

'3.1.2 Gravar o operador que fizer um cadastro novo de cliente
'Private Function OperadorQueGravou(Optional Escrit As Boolean = False) As Integer
'Dim Result As Boolean
'Dim Resp   As String
'Dim Oper$
'Dim rsMec  As Recordset
'
'If INI.UtComissoes = "1" Then
'
'    '3.1.4 Senhas com ** para a identificação do balconista
'    Load frmSenha
'    frmSenha.Tipo = 1
'    frmSenha.Show 1
'    If frmSenha.Resultado = True Then
'        Resp = frmSenha.Senha
'    End If
'    'Unload frmSenha
'    'Resp = InputBox("Informe a sua identifacacao", "Informacao do recebimento")
'    If Resp > "" Then
'        If Escrit Then
'            Oper$ = " and Oper = 2 "
'        Else
'            Oper$ = " and Oper > 0 "
'        End If
'        Set rsMec = BuscaMec(Resp, Oper$)
'        If rsMec.EOF Then
'            MsgBox "Operador não identificado"
'            Exit Function
'        Else
'            txtFields(1).ToolTipText = rsMec!Nome
'            OperadorQueGravou = rsMec!codi
'        End If
'    End If
'    Unload frmSenha
'End If
'End Function

Private Sub VerificaTamanhoEmail(email As String)
Dim nvoTamEmail As Integer

nvoTamEmail = Len(email)
If nvoTamEmail > tamEmail Then
    tamEmail = nvoTamEmail + 1
    Dim sqlData1 As String
    sqlData1 = Data1.Name
    Data1.Close
    ExecSql "Alter Table Clientes Alter Column email Text(" & tamEmail & ")"
    Set Data1 = db.OpenRecordset(sqlData1, dbOpenSnapshot)
End If
End Sub

'3.1.2 Gravar o operador que fizer um cadastro novo de cliente
Private Function AdiCliente(Balconista As Integer) As Integer
'Private Function AdiCliente() As Integer

Dim UltCli    As Integer
Dim UltCliCar As Integer
Dim ULT       As Integer
Dim SQL       As String
Dim TbCli     As Recordset

UltCli = Consulta("Select Max(NrCli) From Clientes")
UltCliCar = Consulta("Select Max(NrCli) From Carros")

If UltCli > UltCliCar Then
    ULT = UltCli
Else
    ULT = UltCliCar
End If
ULT = ULT + 1

VerificaTamanhoEmail txtFields(0).Text
                    
'4.2.8 CEP
SQL = "Insert Into Clientes (NrCli, Nome, Telefone, Vend, Ender, CgcCpf, InscrEst, Operadora, CEP, email) Values (" & ULT
'4.2.7 Operadora do celular
'4.2.7 Inscrição Estatual
'SQL = "Insert Into Clientes (NrCli, Nome, Telefone, Vend, Ender, CgcCpf, InscrEst, Operadora, email) Values (" & ULT
'4.1.4 Campo para CGC e CPF no cadastro de clientes
'SQL = "Insert Into Clientes (NrCli, Nome, Telefone, Vend, Ender, CgcCpf, email) Values (" & ULT
'3.1.2 Gravar o operador que fizer um cadastro novo de cliente
'SQL = "Insert Into Clientes (NrCli, Nome, Telefone, Vend, Ender, email) Values (" & ULT
'SQL = "Insert Into Clientes (NrCli, Nome, Telefone, Ender) Values (" & ULT

SQL = SQL & ",'" & txtFields(1).Text & "'"
SQL = SQL & ",'" & SN(txtFields(3).Text, vbString) & "'"

'3.1.2 Gravar o operador que fizer um cadastro novo de cliente
SQL = SQL & "," & Balconista & " "

'4.0.1 Conserto da adição do cliente, apos a implantação do email
SQL = SQL & ",'" & SN(txtFields(2).Text, vbString) & "'"
'SQL = SQL & ",'" & SN(txtFields(2).Text, vbString) & "')"

'4.1.4 Campo para CGC e CPF no cadastro de clientes
SQL = SQL & "," & FA(txtFields(7).Text)

'4.2.7 Inscrição Estatual
SQL = SQL & "," & FA(txtFields(8).Text)

'4.2.7 Operadora do celular
SQL = SQL & "," & GetOperadora()

'4.2.8 CEP
SQL = SQL & "," & FA(txtFields(9).Text)

SQL = SQL & ",'" & SN(txtFields(0).Text, vbString) & "')"
ExecSql SQL

AdiCliente = ULT

'(" & UltCli & ", '" & txtFields(1).Text & "', '" & SN(txtFields(3).Text, vbString) & "' , '"

'2.0.5 Implantação do Log
'AbreTB TbCli, "Clientes", dbOpenTable
'Set TbCli = db.OpenRecordset("Clientes")

'SQL

'TbCli.AddNew
'TbCli!Nome = txtFields(1).Text
'TbCli!Telefone = SN(txtFields(3).Text, vbString)
'TbCli!Ender = SN(txtFields(2).Text, vbString)
'TbCli.Update
'TbCli.Close
End Function

Private Sub cmdClose_Click()
Loga "Fechar tela Cliente"
Unload Me
End Sub

Private Sub Command1_Click()
Dim Quant As Integer
Continua:

Loga "Adicionar Carro"
'2.0.3 Ajuste no funcionamento das permissões
If Quant > 3 And Permissao = False Then
'    Data2.CancelUpdate
    '2.7.2 Logar todas mensagens
    msgboxL "Na versão demonstração só é permitido 3 registros"
    Command1.Enabled = False
    
Else
    AdicCarro = True
    
    txtFields(4).Enabled = True
    txtFields(5).Enabled = True
    txtFields(6).Enabled = True
    
    txtFields(4).Text = ""
    txtFields(5).Text = ""
    txtFields(6).Text = ""
    
    txtFields(4).Tag = ""
    txtFields(5).Tag = ""
    txtFields(6).Tag = ""
    
    txtFields(4).SetFocus
    Command1.Enabled = False
    Command2.Enabled = False
'    Command4.Enabled = True
    
    DataControl2.Label = "Carro: " + STR(NrCarros + 1) + " de " + STR(NrCarros + 1)
'    opEdtCarro = opVaiAdicionarCarro
    
End If
Exit Sub

Vazio:
Resume Continua
End Sub

Private Sub Command2_Click()
'2.7.2 Logar todas mensagens
If msgboxL("Deseja realmente excluir esse carro?", vbQuestion + vbDefaultButton2 + vbYesNo, "Orcarro") = vbYes Then
    Loga "Apagar Carro"
    
    '2.7.2 Conserto da exclusão do Carro
    ExecSql "Delete From Carros Where Placa = '" & txtFields(5).Text & "'"
    '2.0.5 Implantação do Log
    'ExecSql "Delete From Carros Where NrCli = " & Data1!NrCli
    
    Data2.Delete
    Data2.MoveNext
    If Data2.EOF = False And Data2.BOF = False Then
    
        '2.7.2 Conserto da exclusão do Carro
        NrCarros = NrCarros - 1
        
        Data2_Reposition
    End If
End If
End Sub

'Private Sub Command3_Click()
'frmTransfCarro.Show
'Unload Me
'End Sub

Private Sub Command4_Click()
Dim Quant  As Integer
Dim SQL    As String
Dim CliCom As String
Dim Resp   As VbMsgBoxResult

'5.0.9 Impedir a adição de um carro com uma placa de outro carro
SQL = "SELECT Clientes.Nome "
SQL = SQL & "FROM Carros INNER JOIN Clientes ON Carros.NrCli = Clientes.NrCli "
SQL = SQL & "WHERE Carros.Placa = " & FA(txtFields(5).Text)
CliCom = Consulta(SQL)
If CliCom > "" Then
    Resp = msgboxL("Cliente: " & CliCom & vbCrLf & "Deseja abrir o cadastro dele?", vbYesNo + vbDefaultButton2 + vbQuestion, "Já existe um cliente com essa placa")
    If Resp = vbYes Then
        AbreTB Data1, "Select * From Clientes Where Nome = " & FA(CliCom)
        Carregando = True
        Mostra
    End If
    Exit Sub
End If

If Func Then
    Load frmSenha
    frmSenha.Tipo = 0
    frmSenha.Show 1
    If frmSenha.Resultado = False Then
        Unload frmSenha
        Exit Sub
    End If
    Unload frmSenha
End If
Loga "Gravar Carro"
Quant = Consulta("Select Count(*) From Carros Where NrCli = " & Data1(0))
'2.0.3 Ajuste no funcionamento das permissões
If Quant > 3 And Permissao = False Then
    Command4.Enabled = False
    '2.7.2 Logar todas mensagens
    msgboxL "Na versão demonstração só pode editar 3 carros por cliente"
Else
    GravaCarro
    Command1.SetFocus
    btNvOrcam.Enabled = True
End If
End Sub

Private Sub CarregaCarro(Optional NrDesteCarro As Integer)
Dim Selec$
Dim a As Integer

Selec$ = "SELECT DISTINCTROW NrCli, Modelo, Placa, Cor "
Selec$ = Selec$ + "FROM Carros "
Selec$ = Selec$ + "WHERE NrCli=" & VarNrCliente

'2.0.5 Implantação do Log
'dbOpenDynamic
AbreTB Data2, Selec$, dbOpenDynaset
'Set Data2 = db.OpenRecordset(Selec$)

Data2.MoveLast
Data2.MoveFirst
If NrDesteCarro > 1 Then
    For a = 2 To NrDesteCarro
        Data2.MoveNext
    Next
End If
NrCarros = Data2.RecordCount
Data2_Aberto = True
End Sub

Private Sub Mostra()

'4.2.7 Operadora do celular
Mudando = True

On Error Resume Next
txtFields(1).Text = Data1!Nome

'2.3.3 Possibilitar retornar o carro ao cliente
GCliente = txtFields(1).Text

'2.4.0 Campo para observação do cliente
clsCLi.NrCli = Data1!NrCli

txtFields(2).Text = Data1!Ender
txtFields(3).Text = Data1!Telefone

tamEmail = Data1("email").Size
txtFields(0).Text = SN(Data1!email)

'4.1.4 Campo para CGC e CPF no cadastro de clientes
txtFields(7).Text = SN(Data1!CgcCpf)

'4.2.7 Inscrição Estatual
txtFields(8).Text = SN(Data1!InscrEst)

'4.2.8 CEP
txtFields(9).Text = SN(Data1!CEP)

'4.2.7 Operadora do celular
Dim a As Integer
Dim Operadora As Integer
Operadora = SN(Data1!Operadora)
For a = 0 To 3
    ckOper(a).Value = IIf((a + 1) = Operadora, 1, 0)
Next

NomeOrig = Data1!Nome
If txtFields(1).Text > "" Then
    DataControl1.Label = "Cliente: " & Data1(0)
    VarNrCliente = Data1(0)
    CarregaCarro
    If Data2.RecordCount > 0 Then
        Data2.MoveLast
        Data2.MoveFirst
        
        '2.2.7 Permitir editar o modelo do carro ao pesquisar pelo cliente
        txtFields(4).Enabled = True
        txtFields(5).Enabled = True
        txtFields(6).Enabled = True
        
        If txtFields(4).MaxLength = 0 Then
'            txtFields(4).Enabled = True
'            txtFields(5).Enabled = True
'            txtFields(6).Enabled = True
            
            txtFields(4).MaxLength = Data2.Fields(1).Size
            txtFields(5).MaxLength = Data2.Fields(2).Size
            txtFields(6).MaxLength = Data2.Fields(3).Size
        End If
        NrDesteCarr = 1
    End If
    
    '3.1.2 Gravar o operador que fizer um cadastro novo de cliente
    If INI.UtComissoes = "1" And Data1!Vend > 0 Then
    
        '3.5.1 Não excluir fisicamente Mecânico
        txtFields(1).ToolTipText = Consulta("Select Nome From Mecanicos Where codi = " & Data1!Vend & " and Ativo = True ")
        'txtFields(1).ToolTipText = Consulta("Select Nome From Mecanicos Where codi = " & Data1!Vend)
        
    Else
        txtFields(1).ToolTipText = ""
    End If
    
    '3.4.4 Cliente Funcionario
    ckFunc.Value = IIf(Data1!Funcionario = True, 1, 0)
    Func = ckFunc.Value
    
    '4.5.4 Cliente em atraso
    ckAtraso.Value = IIf(Data1!EmAtraso = True, 1, 0)
    
    '4.5.6 Não permitir novo orçamento para cliente em atraso em modo balcão
    If ckAtraso.Value Then
        If INI.ModoOperacao <> tpEscritorio Then
            btNvOrcam.Enabled = False
        Else
            btNvOrcam.Enabled = True
        End If
    Else
        btNvOrcam.Enabled = True
    End If
    
    Data2_Reposition
Else
   frmClientes.Tag = ""
End If

'4.2.7 Operadora do celular
Mudando = False

'3.1.0 Modo Balcão
''2.6.1 Mecânicos não devem ver valores
'If INI.Restrito = False Then
'    btOrcamento.Enabled = (Consulta("Select Count(*) From Orcamento Where Cliente = '" & txtFields(1).Text & "'") > 0)
'End If
End Sub

'Private Sub Data1_Validate(Action As Integer, Save As Integer)
'Select Case Action
'  'Case vbDataActionMoveFirst
'  'Case vbDataActionMovePrevious
'  'Case vbDataActionMoveNext
'  'Case vbDataActionMoveLast
'  'Case vbDataActionAddNew
'  'Case vbDataActionUpdate
'  'Case vbDataActionDelete
'  'Case vbDataActionFind
'  'Case vbDataActionBookmark
'  'Case vbDataActionClose
'End Select
'End Sub

Private Sub Data2_Reposition()
If Data2.EOF = False And Data2.BOF = False Then
    txtFields(4).Text = Data2!Modelo
    txtFields(5).Text = Data2!Placa
    txtFields(6).Text = Data2!Cor
    
    '3.3.1 Conserto da edição do carro
    txtFields(4).Tag = txtFields(4).Text
    txtFields(5).Tag = txtFields(5).Text
    txtFields(6).Tag = txtFields(6).Text
    
    '2.3.3 Possibilitar retornar o carro ao cliente
    GGPlaca = txtFields(5).Text
    
    If txtFields(4).Text > "" Then
        DataControl2.Label = "Carro: " + STR$(NrDesteCarr) + " de " + STR$(NrCarros)
        'DataControl2.Label = "Carro: " + str$(Data2.AbsolutePosition + 1) + " de " + str$(NrCarros)
    '    CarregaCarros
    Else
        DataControl2.Label = ""
    End If
    
    '5.0.4 Cadastro de Revisões
    cRevis.CarregaRevisoes
    
Else
    txtFields(4).Text = ""
    txtFields(5).Text = ""
    txtFields(6).Text = ""
    Command1.Enabled = False
    Command2.Enabled = False
    Command4.Enabled = True
    txtFields(4).Enabled = True
    txtFields(5).Enabled = True
    txtFields(6).Enabled = True 'Exit Sub
'
'Excluiu_Registro:
'Data2.MoveLast
'Data2.MoveFirst
'txtFields(4).Text = Data2!Modelo
'Resume Next

    DataControl2.Label = ""
    AdicCarro = True
    'NrDesteCarr = 0
End If
End Sub


'Private Sub Data2_Validate(Action As Integer, Save As Integer)
'Select Case Action
'  'Case vbDataActionMoveFirst
'  'Case vbDataActionMovePrevious
'  'Case vbDataActionMoveNext
'  'Case vbDataActionMoveLast
'  'Case vbDataActionAddNew
'  'Case vbDataActionUpdate
'  'Case vbDataActionDelete
'  'Case vbDataActionFind
'  'Case vbDataActionBookmark
'  'Case vbDataActionClose
'End Select
'End Sub

Private Sub DataControl1_Clicado(Clique As PrjDataControl.Acao)
Dim Nr    As Integer
Dim NrAtu As Integer

'5.0.4 Cadastro de Revisões
Carregando = True

AdicCarro = False
If Data1.EOF = False Then
    NrAtu = Data1(0)
End If

Data1.Close
Select Case Clique
    Case PrjDataControl.acPrimeiro
        Loga "Primeiro Cliente"
        Nr = Consulta("Select NrCli From Clientes Order By Nome")
    Case PrjDataControl.acAnterior
        Loga "Cliente Anterior"
        Nr = Consulta("Select NrCli From Clientes Where Nome < '" & Trim(txtFields(1)) & "' Order By Nome desc")
    Case PrjDataControl.acProximo
        Loga "Cliente Proximo"
        Nr = Consulta("Select NrCli From Clientes Where Nome > '" & Trim(txtFields(1)) & "' Order By Nome")
    Case PrjDataControl.acUtimo
        Loga "Cliente Ultimo"
        Nr = Consulta("Select NrCli From Clientes Order By Nome")
End Select
If Nr = 0 Then
    Nr = NrAtu
End If
VarNrCliente = Nr

'2.0.5 Implantação do Log
AbreTB Data1, "Select * From Clientes Where NrCli = " & Nr
'Set Data1 = db.OpenRecordset("Select * From Clientes Where NrCli = " & Nr)

If Data1.EOF = False And Data1.BOF = False Then
    Mostra
End If

'5.0.4 Cadastro de Revisões
Carregando = False
End Sub

Private Sub DataControl2_Clicado(Clique As PrjDataControl.Acao)
AdicCarro = False
Select Case Clique
    Case PrjDataControl.acPrimeiro
        Data2.MoveFirst
        NrDesteCarr = 1
    Case PrjDataControl.acAnterior
            NrDesteCarr = NrDesteCarr - 1
            If AdicionouCar Then
                CarregaCarro NrDesteCarr
            Else
                If Data2.BOF = False Then
                    Data2.MovePrevious
                End If
            End If
    Case PrjDataControl.acProximo
        If Data2.EOF = False Then
            Data2.MoveNext
        End If
    Case PrjDataControl.acUtimo
        Data2.MoveLast
        NrDesteCarr = NrCarros
End Select
If Data2.EOF = False And Data2.BOF = False Then
    Data2_Reposition
End If
AdicionouCar = False
NrDesteCarr = LocalizaCarro()
DataControl2.Label = "Carro: " & NrDesteCarr & " de " & NrCarros
End Sub

Private Sub Form_Activate()
If Data1.RecordCount = 0 Then
   txtFields(1).Enabled = False
   txtFields(2).Enabled = False
   txtFields(3).Enabled = False
   txtFields(0).Enabled = False
   
   '4.1.4 Campo para CGC e CPF no cadastro de clientes
   txtFields(7).Enabled = False
   
   '4.2.7 Inscrição Estatual
   txtFields(8).Enabled = False
   
   '4.2.8 CEP
   txtFields(9).Enabled = False
   
Else
   txtFields(1).MaxLength = Data1.Fields(1).Size
   txtFields(2).MaxLength = Data1.Fields(2).Size
   txtFields(3).MaxLength = Data1.Fields(3).Size
End If

'5.0.4 Cadastro de Revisões
Carregando = False
End Sub

Private Sub Form_Load()
Dim SQL As String

'5.0.4 Cadastro de Revisões
Carregando = True

InicForm Me, MudaTodos:=True

'5.0.4 Cadastro de Revisões
Set cRevis = New clsRevisoes

'2.3.3 Possibilitar retornar o carro ao cliente
GCliente = ""
GGPlaca = ""

'4.2.8 CEP
SQL = "SELECT DISTINCTROW NrCli, Nome, Ender, Telefone, Vend, Funcionario, email, CgcCpf, InscrEst, CEP, Operadora "
'4.2.7 Operadora do celular
'4.2.7 Inscrição Estatual
'SQL = "SELECT DISTINCTROW NrCli, Nome, Ender, Telefone, Vend, Funcionario, email, CgcCpf, InscrEst, Operadora "
'4.1.4 Campo para CGC e CPF no cadastro de clientes
'SQL = "SELECT DISTINCTROW NrCli, Nome, Ender, Telefone, Vend, Funcionario, email, CgcCpf "

'3.1.2 Gravar o operador que fizer um cadastro novo de cliente
'SQL = "SELECT DISTINCTROW NrCli, Nome, Ender, Telefone, Vend, Funcionario, email "

'4.5.4 Cliente em atraso
SQL = SQL & ", EmAtraso "

SQL = SQL & "from Clientes "
SQL = SQL & "ORDER BY Nome"
AbreTB Data1, SQL, dbOpenDynaset

'2.2.2 Impedir o erro que impedia de alterar um cadastro de um cliente recem cadastrado
Adicionando = False

'3.1.6 Impedir que os mecânicos vejam orçamentos atravez da tela de clientes
If INI.ModoOperacao = tpEscritorio Then

    '4.7.6 Pesquisar por cliente ou funcionário no recibo avulso
    btRecibo.Enabled = True

    '3.7.9 Permitir a troca pelo balcao
    '3.7.8 Troca de carros entre os clientes
'    btTroca.Enabled = True
    
Else
    '3.7.9 Permitir a troca pelo balcao
    If INI.ModoOperacao = tpMecanico Then
        btTroca.Enabled = False
    End If
    
    cmdDelete.Enabled = False
    Command2.Enabled = False
    
    '4.5.4 Cliente em atraso
    ckAtraso.Enabled = False
    
End If

If INI.ModoOperacao = tpMecanico Then
    btOrcamento.Enabled = False
    btNvOrcam.Enabled = False
    btObs.Enabled = False
End If

'3.1.0 Modo Balcão
'If INI.ModoOperacao = tpBalcao Then
'    cmdDelete.Enabled = False
'    Command2.Enabled = False
'
'''2.5.6 Modo Restrito
''If INI.Restrito Then
''    cmdUpdate.Enabled = False
''    Command4.Enabled = False
''    BtImprCabec.Enabled = False
''    btNvOrcam.Enabled = False
''
''    '2.6.1 Mecânicos não devem ver valores
''    btOrcamento.Enabled = False
''    cmdDelete.Enabled = False
''    Command2.Enabled = False
''
''    '2.8.7 Não mostrar a obs do cliente no modo restrito
''    btObs.Enabled = False
''
'End If
End Sub

Private Sub Frame1_DblClick()
'3.1.0 Modo Balcão
Command2.Enabled = (INI.ModoOperacao = tpEscritorio)
'Command2.Enabled = (INI.Restrito = False)
End Sub

Private Sub txtFields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
End Sub

Public Property Let NrCliente(vNewValue As String)
VarNrCliente = vNewValue
Data1.FindFirst "[NrCli] = " & vNewValue
Mostra
'CarregaCarros
If Data2_Aberto Then
    Data2_Reposition
End If
End Property

Private Function VerificaCarro() As Boolean
Dim VerifCar As Recordset
Dim Selec    As String
Dim Carro    As String

Carro = Chr(34) + txtFields(5) + Chr(34)
Selec = "SELECT NrCli From Carros WHERE Placa  = '" & Carro & "'"

Selec = "SELECT Clientes.Nome "
Selec = Selec & "From Clientes, Carros "
Selec = Selec & "WHERE Carros.Placa  = " & Carro & " and Clientes.NrCli = Carros.NrCli"

'2.0.5 Implantação do Log
AbreTB VerifCar, Selec, dbOpenDynaset
'Set VerifCar = db.OpenRecordset(Selec, dbOpenDynaset)

On Error GoTo NãoExisteOCliente
VerifCar.MoveLast
On Error GoTo 0
'2.7.2 Logar todas mensagens
msgboxL "Impossível fazer esta gravação porque este carro pertence ao cliente" + Chr(13) + VerifCar(0), vbCritical, "Duplicidade de Placa"
VerifCar.Close
VerificaCarro = False
Exit Function

NãoExisteOCliente:
VerifCar.Close
Selec = "SELECT NrCli from carros where placa = " & Carro

'2.0.5 Implantação do Log
AbreTB VerifCar, Selec, dbOpenDynaset
'Set VerifCar = db.OpenRecordset(Selec, dbOpenDynaset)

VerifCar.Edit
VerifCar(0) = Data1(0)
VerifCar.Update
VerifCar.Close
VerificaCarro = True
End Function

Public Property Let nrPlaca(ByVal vNewValue As String)
Data2.MoveLast
Data2.MoveFirst
Data2.FindFirst "[Placa] = '" & vNewValue & "'"
Data2_Reposition
End Property

'3.4.4 Cliente Funcionario
Public Property Get Func() As Boolean
Func = l_Func
End Property

'3.4.4 Cliente Funcionario
Public Property Let Func(ByVal vNewValue As Boolean)
l_Func = vNewValue
End Property

Private Sub txtFields_LostFocus(Index As Integer)
If Index = 0 Then
    If txtFields(0).Text > "" Then
        If CriticaEmail(txtFields(0).Text) = False Then
            txtFields(0).Text = ""
            txtFields(0).SetFocus
        End If
    End If
End If
End Sub

'4.2.7 Operadora do celular
Private Function GetOperadora() As Integer
Dim a As Integer

'4.2.9 Ajuste na gravação da operadora do celular
For a = 0 To 3
'For a = 0 To 4

    If ckOper(a).Value = 1 Then
        GetOperadora = a + 1
        Exit Function
    End If
Next
End Function
