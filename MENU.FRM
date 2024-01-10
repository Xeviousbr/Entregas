VERSION 5.00
Begin VB.Form FrmMenu 
   Caption         =   "Gerenciamento de Orçamentos"
   ClientHeight    =   2880
   ClientLeft      =   2040
   ClientTop       =   3030
   ClientWidth     =   6810
   Icon            =   "MENU.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2880
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Imagem 
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2835
      ScaleWidth      =   6795
      TabIndex        =   1
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton btTarefa 
         Caption         =   "Lançar Taréfa"
         Height          =   495
         Left            =   2108
         TabIndex        =   3
         Top             =   1680
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label lbOperacao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   120
         TabIndex        =   2
         Top             =   2280
         Visible         =   0   'False
         Width           =   6540
      End
   End
   Begin VB.TextBox txtZip 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Text            =   "UtilizadoPeloZip"
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Menu mnuTabela 
      Caption         =   "Tabela"
   End
   Begin VB.Menu Mnu_Orc 
      Caption         =   "&Orçamentos"
      Begin VB.Menu MnuOrc 
         Caption         =   "&Novo"
         Index           =   0
      End
      Begin VB.Menu MnuOrc 
         Caption         =   "Localizar pelo &Nome"
         Index           =   1
      End
      Begin VB.Menu MnuOrc 
         Caption         =   "Localizar pela &Placa"
         Index           =   2
      End
      Begin VB.Menu MnuOrc 
         Caption         =   "&Relação de Orçamentos"
         Index           =   3
      End
      Begin VB.Menu MnuOrc 
         Caption         =   "&Totais"
         Index           =   4
         Begin VB.Menu MnuRelTot 
            Caption         =   "&Detalhado"
            Index           =   0
         End
         Begin VB.Menu MnuRelTot 
            Caption         =   "&Resumido"
            Index           =   1
         End
      End
      Begin VB.Menu MnuPagPagar 
         Caption         =   "&Valor pagos e a pagar"
      End
   End
   Begin VB.Menu Mnu_Cli 
      Caption         =   "&Clientes"
      Begin VB.Menu MnuCli 
         Caption         =   "&Editar"
         Index           =   0
      End
      Begin VB.Menu MnuCli 
         Caption         =   "Localizar pelo &Nome"
         Index           =   1
      End
      Begin VB.Menu MnuCli 
         Caption         =   "Localizar pela &Placa"
         Index           =   2
      End
      Begin VB.Menu MnuCli 
         Caption         =   "Localizar pelo &CPF/CNPj"
         Index           =   3
      End
      Begin VB.Menu MnuCli 
         Caption         =   "Revisões"
         Index           =   4
         Begin VB.Menu MnuRev 
            Caption         =   "Pesquisa"
            Index           =   0
         End
         Begin VB.Menu MnuRev 
            Caption         =   "Cadastro"
            Index           =   1
         End
         Begin VB.Menu MnuRev 
            Caption         =   "Pendencias"
            Index           =   2
         End
      End
   End
   Begin VB.Menu Mnu_Mec 
      Caption         =   "Operadores"
      Begin VB.Menu MnuMec 
         Caption         =   "Pesquisa"
         Index           =   0
      End
      Begin VB.Menu MnuMec 
         Caption         =   "Cadastro"
         Index           =   1
      End
      Begin VB.Menu MnuMec 
         Caption         =   "Recibo"
         Index           =   2
         Begin VB.Menu mnuRec 
            Caption         =   "Adiantamento"
            Index           =   0
         End
         Begin VB.Menu mnuRec 
            Caption         =   "Comissão"
            Index           =   1
         End
         Begin VB.Menu mnuRec 
            Caption         =   "Vale Transporte"
            Index           =   2
         End
         Begin VB.Menu mnuRec 
            Caption         =   "Pagamento"
            Index           =   3
         End
         Begin VB.Menu mnuRec 
            Caption         =   "Avulso"
            Index           =   4
         End
      End
      Begin VB.Menu MnuMec 
         Caption         =   "Taréfas"
         Index           =   3
      End
      Begin VB.Menu MnuMec 
         Caption         =   "Pesquisa de Tarefas"
         Index           =   4
      End
      Begin VB.Menu MnuProdutividade 
         Caption         =   "Produtividade"
      End
      Begin VB.Menu MnuRelComiss 
         Caption         =   "Relatório de Pagamentos"
      End
   End
   Begin VB.Menu MnuF 
      Caption         =   "&Fornecedores"
      Begin VB.Menu MnuFo 
         Caption         =   "&Pesquisa"
         Index           =   0
      End
      Begin VB.Menu MnuFo 
         Caption         =   "&Cadatro"
         Index           =   1
      End
      Begin VB.Menu MnuFo 
         Caption         =   "&Recibo"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu MnuFo 
         Caption         =   "&Bancos"
         Index           =   3
         Begin VB.Menu MnuBco 
            Caption         =   "&Pesquisa"
            Index           =   0
         End
         Begin VB.Menu MnuBco 
            Caption         =   "&Cadastro"
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "Backup"
      Begin VB.Menu BkpDisk 
         Caption         =   "Backup para ?"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuRestDisk 
         Caption         =   "Restauração apartir de ?"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuBackup 
         Caption         =   "Definir caminho do Backup"
      End
   End
   Begin VB.Menu Mnu_Reg 
      Caption         =   "Registrar"
      Visible         =   0   'False
   End
   Begin VB.Menu MnuOpcoes 
      Caption         =   "Opções"
      Begin VB.Menu mnuConfig 
         Caption         =   "Configuração"
      End
      Begin VB.Menu MnuModelo 
         Caption         =   "Modelo dos Itens de Orçamento"
      End
      Begin VB.Menu MnuFerramentas 
         Caption         =   "Ferramentas"
      End
      Begin VB.Menu MnuAjustes 
         Caption         =   "Ajuste de orçamentos"
      End
      Begin VB.Menu mnuImpress 
         Caption         =   "&Impressoras"
      End
   End
   Begin VB.Menu Mnu_Help 
      Caption         =   "So&bre"
   End
End
Attribute VB_Name = "FrmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'5.1.1 Pesquisa por CPF e CNPj
'5.0.4 Cadastro de Revisões
'4.9.8 Selecionar a impressora ao imprimir
'4.9.7 Tela para gravar se a impressora é tipo fita ou não
'4.9.1 OrCarro.ini fixado em C:\OrCarro
'4.3.8 Correção da chamada do menu dos recibos
'4.3.7 Tela para recibo avulso
'4.4.2 Cadastro de fornecedores
'4.3.0 Recibo avulso
'3.8.6 Operação de adaptação de orçamentos
'3.8.4 Adaptação da resolução
'3.8.3 Relatório de pagamentos
'3.7.6 Mecânico com recebimento semanal
'3.7.0 Ferramentas
'3.7.0 Retirar preparação para tarefas dinâmicas
'3.4.8 Ajuste na restauração do Backup
'3.4.1 Conferência interna na operação de gravação das tarefas pelos mecânicos
'3.4.0 Ajuste na gravação das tarefas pelos mecanicos
'3.2.9 Nova tela de tabela de preços
'3.2.7 Gravação do valor da mão de obra
'3.2.1 Pesquisa de tarefas realizadas
'3.1.0 Modo Balcão
'3.0.2 Backup da pasta Log
'2.9.7 Relatório de Totais Resumido
'2.9.6 Retirar o orçamento novo do log de tempo do carregamento do orçamento
'2.9.2 Retornar a mostrar a tela de tarefas para os mecânicos
'2.9.1 Deixar de mostrar a tela de tarefas, pelo menu, no modo mecanico
'2.8.6 Data da conclusão da tarefa
'2.8.5 Conserto da situação em que tarefas de orçamentos antigos não podiam ser concluídas
'2.8.0 Melhorar o log quanto as tarefas
'2.7.7 Ajuste no critério de primeira taréfa dinâmica
'2.7.5 Taréfas Dinâmicas
'2.7.4 Registro facilitado das taréfas dos mecânicos
'2.7.2 Logar todas mensagens
'2.6.3 Cadastro de Mecânicos
'2.6.0 Mostrar a tela de lista de orçamentos mesmo quando não tem orçamentos no mes
'2.5.6 Modo Restrito
'2.5.0 Configuração do modelo de itens
'2.4.6 Relação de totais de valores de Orçamentos
'2.4.2 Observação do orçamento na relação mensal

Option Explicit

Private frmClientesLido As Boolean
Private TemQueAdaptar   As Boolean
Private Carregando      As Boolean
Private SituMax         As Integer
Private TamArqZip       As Integer
Private OperZip         As Byte        'OperZip=0=Zipando OperZip=1=Deszipando

'5.0.4 Cadastro de Revisões
Private JaPassou        As Boolean

Private Sub LocalizaCliente(Pesq As String)
Dim SQL     As String
Dim CadeCli As Recordset

If InStr(Pesq, "'") > 0 Then
    '2.7.2 Logar todas mensagens
    msgboxL "Digite o nome sem apóstrofes"
    Exit Sub
End If

Pesq = "*" & Trim$(Pesq) & "*"

SQL = "SELECT NrCli "
SQL = SQL + "From Clientes "
SQL = SQL + "WHERE ((Nome Like '" + Pesq + "')); "

AbreTB CadeCli, SQL, dbOpenDynaset

If CadeCli.EOF Then
   '2.7.2 Logar todas mensagens
   msgboxL "Não localizei o cliente"
Else
   Load frmClientes
   frmClientes.NrCliente = CadeCli("NrCli")
   frmClientes.Show
End If
End Sub

Private Sub BkpDisk_Click()
FazBackup CaminhoBkp & "\OrCarro.zip"
End Sub

Private Sub ChamaTarefasSimples()
'2.7.4 Registro facilitado das taréfas dos mecânicos
Dim Placa   As String
Dim Resp    As String
Dim SQL     As String
Dim Obs     As String
Dim nmMec   As String
Dim rsAux   As Recordset
Dim rsMec   As Recordset

Placa = InputBox("Informe a placa")
If Placa > "" Then
    SQL = "SELECT Tarefas.ID, Tarefas.concerto, Tarefas.Situacao, Tarefas.Mec, Orcamento.Obs, Orcamento.ObsMec "
    SQL = SQL & "from Tarefas, Orcamento, Carros "
    SQL = SQL & "Where Carros.Placa='" & Placa & "' "
    SQL = SQL & "and Orcamento.Carro = Carros.Placa "
    SQL = SQL & "and Tarefas.Orc = Orcamento.Orcamento "
    
    '2.8.5 Conserto da situação em que tarefas de orçamentos antigos não podiam ser concluídas
    '2.7.7 Ajuste no critério de primeira taréfa dinâmica
    'SQL = SQL & "and Tarefas.id > " & INI.Orc1
    
    '2.7.9 Conserto da conclusão da tarefa pelo botão
    SQL = SQL & "and Tarefas.Situacao < 3 "
    'SQL = SQL & "and Tarefas.Situacao = 1 "
    
    AbreTB rsAux, SQL, dbOpenDynaset
    If rsAux.EOF Then
        MsgBox "Não existe tarefa para este carro"
    Else
        '2.7.5 Taréfas Dinâmicas
        rsAux.MoveLast
        rsAux.MoveFirst
        If rsAux.RecordCount = 1 Then
            If SN(rsAux!Obs, vbString) > " " Then
                Obs = "Observação: " & vbCrLf & rsAux!Obs & vbCrLf & "Observação do Mecânico: " & vbCrLf & SN(rsAux!ObsMec, vbString) & vbCrLf & vbCrLf
            End If
            Select Case rsAux!Situacao
                Case 1
                    Resp = InputBox(Obs & "Informe sua identificação", "Taréfa esta livre")
                    If Resp = "" Then Exit Sub
                    Set rsMec = BuscaMec(Resp)
                    If rsMec.EOF Then
                        MsgBox "Mecânico não identificado"
                    Else
                    
                        '2.8.0 Melhorar o log quanto as tarefas
                        If INI.Log Then
                            LogaTarefa rsAux!ID, Placa, rsAux!concerto, rsMec!Nome
                        End If
                    
                        'MECÂNICO ASSUMIU A TAREFA
                        '3.4.1 Conferência interna na operação de gravação das tarefas pelos mecânicos
'                        SQL = "Update Tarefas "
'                        SQL = SQL & "Set Mec = " & rsMec!codi
'
'                        '2.7.5 Taréfas Dinâmicas
'                        SQL = SQL & ", Situacao = 2 "
'                        'SQL = SQL & ", Situacao = 1 "
'
'                        '3.4.0 Ajuste na gravação das tarefas pelos mecanicos
'                        SQL = SQL & ", DtAssumiu = " & DTSqls(Format(Now, "DD/MM/YYYY HH:MM:SS"))
'                        '3.3.8 Data que assumiu a assumiu as tarefas
'                        'SQL = SQL & ", DtAssumiu = " & DTSqls(Format(Now, "DD/MM/YYYY HH:SS"))
'
'                        SQL = SQL & " Where ID = " & rsAux!ID
'                        ExecSql SQL
                        GravaTarefaMecanicos rsAux!ID, 2, rsMec!codi
                        
                        lbOperacao.Caption = "Taréfa do carro " & Placa & " foi anotada para " & rsMec!Nome
                        lbOperacao.Visible = True
                    End If
                Case 2
                
                    '3.5.1 Não excluir fisicamente Mecânico
                    nmMec = Consulta("Select Mecanicos.Nome From Mecanicos, Tarefas Where Mecanicos.codi = Tarefas.Mec and Tarefas.ID = " & rsAux!ID & " and Mecanicos.Ativo = True ")
                    'nmMec = Consulta("Select Mecanicos.Nome From Mecanicos, Tarefas Where Mecanicos.codi = Tarefas.Mec and Tarefas.ID = " & rsAux!ID)
                    
                    Resp = InputBox(Obs & "deseja CONCLUÍ-LA ?" & vbCrLf & vbCrLf & "Informe sua identificação" & vbCrLf & nmMec, "Taréfa esta em andamento")
                    If Resp = "" Then Exit Sub
                    Set rsMec = BuscaMec(Resp)
                    If rsMec.EOF Then
                        MsgBox "Mecânico não identificado"
                    Else
                    
                        '2.8.5 Conserto da situação em que tarefas de orçamentos antigos não podiam ser concluídas
                        If nmMec = rsMec!Nome Or nmMec = "" Then
                        'If nmMec = rsMec!Nome Then
                        
                            '2.8.0 Melhorar o log quanto as tarefas
                            If INI.Log Then
                                LogaTarefa rsAux!ID, Placa, rsAux!concerto, rsMec!Nome, True
                            End If
                            
                            'GravaTarefa ID, Situacao, Mecanico
                        
                            'MECÂNICO CONCLUIU A TAREFA
                            '3.4.1 Conferência interna na operação de gravação das tarefas pelos mecânicos
'                            SQL = "Update Tarefas "
'                            SQL = SQL & "Set Mec = " & rsMec!codi
'                            SQL = SQL & ", Situacao = 3 "
'
'                            '3.4.0 Ajuste na gravação das tarefas pelos mecanicos
'                            SQL = SQL & ",DtConclusao = " & DTSqls(Format(Now, "DD/MM/YYYY HH:MM:SS"))
'                            '3.3.8 Data que assumiu a assumiu as tarefas
'                            'SQL = SQL & ",DtConclusao = " & DTSqls(Format(Now, "DD/MM/YYYY HH:SS"))
'                            '2.8.6 Data da conclusão da tarefa
'                            'SQL = SQL & ",DtConclusao = Int(Now) "
'
'                            SQL = SQL & " Where ID = " & rsAux!ID
'                            ExecSql SQL
                            GravaTarefaMecanicos rsAux!ID, 3, rsMec!codi
                            
                            lbOperacao.Caption = "Taréfa do carro " & Placa & " foi concluída"
                            lbOperacao.Visible = True
                        Else
                            MsgBox "Mecânico não identificado"
                        End If
                    End If
                Case 3
                    MsgBox "Esta taréfa já consta como concluída"
            End Select
        Else
            Load frmTarefas
            frmTarefas.Placa = Placa
            frmTarefas.Show
        End If
'        If rsAux!Mec = 0 Then
'        End If
    End If
End If
End Sub

Private Sub LogaTarefa(Orc As Long, Placa As String, Conserto As Integer, Mec As String, Optional Conclusao As Boolean = False)
'2.8.0 Melhorar o log quanto as tarefas
Dim sConserto As String

sConserto = Consulta("Select concerto from tpConcertos Where tipo = " & Conserto)
Loga " ", lDBG
Loga "Tarefa de " & sConserto & " do carro " & Placa & " " & IIf(Conclusao, "concluída", "assumida") & " por " & Mec, lDBG
Loga " ", lDBG
End Sub

Private Sub btTarefa_Click()
'2.7.4 Registro facilitado das taréfas dos mecânicos
If Carregando = False Then
    ChamaTarefasSimples
End If
End Sub

Private Sub Form_Activate()
'5.0.4 Cadastro de Revisões
If JaPassou = False Then
    JaPassou = True
'    ProcuraRevisoes
End If
End Sub

'5.0.4 Cadastro de Revisões
Private Sub ProcuraRevisoes()
Dim SQL        As String
Dim rsRevisoes As Recordset

SQL = "Select * From ( "
SQL = SQL & "SELECT RevisoesCarros.IdRevCarros "
SQL = SQL & ", DateAdd('m', Revisoes.Meses,RevisoesCarros.Data) as Tempo "
SQL = SQL & "FROM ((Revisoes INNER JOIN RevisoesCarros ON Revisoes.ID = RevisoesCarros.idRevisao) "
SQL = SQL & "INNER JOIN Carros ON RevisoesCarros.Placa = Carros.Placa) "
SQL = SQL & "INNER JOIN Clientes ON Carros.NrCli = Clientes.NrCli ) X "
SQL = SQL & "Where Tempo < Now "
AbreTB rsRevisoes, SQL, dbOpenSnapshot
If rsRevisoes.EOF = False Then
    Load frmRevisao
    frmRevisao.Revisao = rsRevisoes!IdRevCarros
    frmRevisao.Show
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then End
End Sub

Private Sub Form_Load()
Dim Inicio     As Double
Dim X          As Long
Dim PathImagem As String

'4.5.7 Aceleração na inicializalização
'Dim Fim        As Double

Const Segundos = 1

Carregando = True

'2.9.6 Retirar o orçamento novo do log de tempo do carregamento do orçamento
MomChamOrc = -1

'4.5.7 Aceleração na inicializalização
'Fim = Now + (Segundos / 24 / 60 / 60)

'4.9.1 OrCarro.ini fixado em C:\OrCarro
AplicDirNat = "C:\OrCarro"

Inicializa
If INI.Log = True Then
    PreparaOLog
End If
Loga "", lDBG
Loga "Orcarro acionado em " & Format(Now, "DD/MM/YYYY HH:MM:SS")
PoeNome
Splash.Show
X = DoEvents
LinhaDeComando = UCase(Command)

'2.5.6 Modo Restrito
PoemTituloAplicacao Me

'3.1.0 Modo Balcão
DefineMenus INI.ModoOperacao
'ModoRestrito INI.Restrito

'2.4.0 Campo para observação do cliente
'PreparaObsCliente

'4.5.7 Aceleração na inicializalização
'If InStr(LinhaDeComando, "ATC") = 0 Then
'    Do While Fim > Now
'        X = DoEvents
'    Loop
'End If

FrmMenu.Mnu_Reg.Visible = (Permissao = False)

Unload Splash

PathImagem = CarregaConfig
If PathImagem > "" Then
    FrmMenu.Visible = True
    X = DoEvents()
    lbOperacao.Visible = True
    lbOperacao.Caption = "Procurando por imagem: " & PathImagem
    lbOperacao.Refresh
    If FileExists(PathImagem) Then
    
        '3.8.4 Adaptação da resolução
        Imagem.Visible = True
    
        Imagem.Picture = LoadPicture(PathImagem)
    Else
        ExecSql "Update Config Set Imagem = ''"
    End If
End If
lbOperacao.Visible = False
lbOperacao.BackColor = INI.Cor

'3.1.0 Modo Balcão
'If INI.Restrito = True Then
    '2.9.2 Retornar a mostrar a tela de tarefas para os mecânicos
    '2.9.1 Deixar de mostrar a tela de tarefas, pelo menu, no modo mecanico
    'Mnu_Mec.Visible = False
    
'    btTarefa.Visible = True
'End If

'3.7.0 Retirar preparação para tarefas dinâmicas
'2.7.5 Taréfas Dinâmicas
'If INI.Versao < "2.7.5" Then
'    lbOperacao.Caption = "É necessário adaptar as taréfas"
'    lbOperacao.Visible = True
'    PrepTarDin.Visible = True
'End If

Loga "PC = " & INI.PC, lDBG

'3.8.4 Adaptação da resolução
InicForm Me

Carregando = False
End Sub

'Private Sub PreparaObsCliente()
''2.4.0 Campo para observação do cliente
'If Consulta("Select Versao From Config") < "2.4.0" Then
'    ExecSql "ALTER TABLE Clientes ADD Observacao Memo"
'    ExecSql "Update Config Set Versao = '" & AppVersao & "'"
'End If
'End Sub

Private Sub Form_Unload(Cancel As Integer)
Loga "Programa Fechado Normalmente"
Loga "", lDBG
End
End Sub

Private Sub Mnu_Help_Click()
''Application.Dialogs(xlDialogPrinterSetup).Show
'CommonDialog1.PrinterDefault = False
'CommonDialog1.ShowPrinter
'If CommonDialog1.PrinterDefault Then
'    CommonDialog1 = CommonDialog1
'Else
'    CommonDialog1 = CommonDialog1
'End If
'
'For Each Imp in Printers
'If Imp.hDC = CommonDialog1.hDC Then
'   'Essa foi a escolhida
'End If
'Next

Load Splash
Splash.Caption = "Orçamentos para mecânica"
'Colorizar Me
Splash.Show
End Sub

Private Sub FazBackup(Optional Caminho As String = "")
Dim CaminhoBkp  As String
Dim Linha1      As String
Dim Aqui        As String
Dim CaminhoOri  As String

Caption = "Compactando OrCarro.mdb"
db.Close
Refresh

If frmClientesLido Then
    Unload frmClientes
End If

CompactaBD Base

Caption = "Zipando"
Refresh

Open "OrCarro.mdb" For Binary As #2
TamArqZip = LOF(2) / 1024
Close #2

OperZip = 0
InicializaZip Me, txtZip
Compacta "OrCarro.zip", "OrCarro.mdb"

'3.0.2 Backup da pasta Log
Compacta "OrCarro.zip", App.Path & "\Log\*.log"

Set db = OpenDatabase(Base)
Aqui = App.Path
If Caminho = "" Then
    If CD_ShowOpen_Save(Me.hWnd, , CaminhoBkp, "OrCarro.zip", ".zip", "Informe onde devo salvar o backup", "*.zip", True) = True Then
        Caminho = CaminhoBkp
    Else
        Exit Sub
    End If
End If
On Local Error Resume Next
ChDrive Left(Aqui, 3)
ChDir Aqui
On Local Error GoTo 0

On Error GoTo Err_Backup
FileCopy "OrCarro.zip", Caminho
On Error GoTo 0

Kill "OrCarro.zip"

Caption = "Abrindo a Base"
Set db = OpenDatabase(Base)

Caption = "Backup Concluído"
'2.7.2 Logar todas mensagens
msgboxL "Backup Concluído", vbInformation, "Estoq"

Sai_backup:
Caption = "Gerenciador de Orçamentos"
Exit Sub

Err_Backup:
'2.7.2 Logar todas mensagens
If msgboxL("Não foi possível copiar os arquivos" + Chr(13) + "Quer tentar novamente?", vbYesNo + vbQuestion, "Estoq") = vbYes Then
   Resume
Else
   Resume Sai_backup
End If
End Sub

Private Sub Mnu_Reg_Click()
Registro.Show
End Sub

Private Sub MnuAjustes_Click()
'3.8.6 Operação de adaptação de orçamentos
frmAjustes.Show
End Sub

Private Sub MnuBackup_Click()
Dim sPath   As String
Dim PathAtu As String

PathAtu = App.Path
If CD_ShowFolder(Me.hWnd, , "", sPath, "Informe a unidade de Backup") Then
    CaminhoBkp = sPath
    INI.CaminhoBkp = CaminhoBkp
    SetaDirBackup
    ChDir PathAtu
End If
End Sub

Private Sub MnuBco_Click(Index As Integer)
If Index Then
    Load CadBancos
    CadBancos.Adicionar
    CadBancos.Show
Else
    Load frmBusca
    frmBusca.Tipo = 2
    frmBusca.Show
End If
End Sub

Private Sub MnuCli_Click(Index As Integer)
Dim Selec As String
Dim Pesq$

GCliente = ""
GGPlaca = ""
Select Case Index
   Case 0
      Load frmClientes
      frmClientes.NrCliente = PrimeiroCli
      frmClientes.Show
      frmClientesLido = True
   Case 1
      Pesq$ = InputBox("Digite o nome")
      If Pesq$ > "" Then
         Loga "Pesquisa de Cliente pelo Nome: " & Pesq$
         LocalizaCliente Pesq$
      End If
   Case 2
      Pesq$ = InputBox("Digite a placa")
      If Pesq$ > "" Then
         Loga "Pesquisa de Cliente pela Placa: " & Pesq$
         LocalizaPlaca Pesq$
      End If
      
   '5.1.1 Pesquisa por CPF e CNPj
   Case 3
      Pesq$ = InputBox("Digite o CPF/CNPj")
      If Pesq$ > "" Then
         Loga "Pesquisa de Cliente pelo CPF: " & Pesq$
         LocalizaCPF Pesq$
      End If
      
End Select
End Sub

Private Sub mnuConfig_Click()
frmConfig.Show
End Sub

Private Sub MnuFerramentas_Click()
'3.7.0 Ferramentas
Ferramentas.Show
End Sub

'4.4.2 Cadastro de fornecedores
Private Sub MnuFo_Click(Index As Integer)
Select Case Index
    Case 0
        Load frmBusca
        frmBusca.Tipo = 1
        frmBusca.Show
    Case 1
        Load Forneced
        Forneced.Show
        Forneced.Adicionar
End Select
End Sub

'4.9.7 Tela para gravar se a impressora é tipo fita ou não
Private Sub mnuImpress_Click()
frmConfigImpr.Show

'4.9.8 Selecionar a impressora ao imprimir
frmConfigImpr.Menu
End Sub

Private Sub MnuMec_Click(Index As Integer)
'2.6.3 Cadastro de Mecânicos
Select Case Index
    Case 0
        '2.6.7 Ajuste a primeira operação com cadastro de mecânicos
        On Local Error Resume Next
        
        '4.4.2 Cadastro de fornecedores
        Load frmBusca
        frmBusca.Tipo = 0
        
        frmBusca.Show
        On Local Error GoTo 0
    Case 1
        CadMecanicos.Show
        
    '3.4.4 Vale Transporte
'    Case 2
'        frmRecibo.Show

    Case 3
        '2.7.5 Taréfas Dinâmicas
        Load frmTarefas
        frmTarefas.Placa = ""
        '2.7.4 Campo de senha para o mecânico
        frmTarefas.Show
        
        '3.2.1 Pesquisa de tarefas realizadas
    Case 4
        TarefasRealizadas.Show
        
End Select
End Sub

Private Sub MnuModelo_Click()
'2.5.0 Configuração do modelo de itens
frmModeloItens.Show
End Sub

Private Sub MnuOrc_Click(Index As Integer)
Dim Pesq$
Dim SQL As String

GCliente = ""
GGPlaca = ""
Select Case Index
   Case 0
   
      
      'frmOrc.Show
      
      '2.7.5 Taréfas Dinâmicas
      NovasTarefas
      
      Load frmOrc
      frmOrc.NrOrcamento = 0
      frmOrc.Show
      
   Case 1
      Pesq$ = InputBox("Digite o nome:", "Localização de Orçamentos")
      If Pesq$ > "" Then
          Loga "Pesquisa de Orçamento pelo Nome: " & Pesq$
          Pesq$ = Chr$(34) + Trim$(Pesq$) + Chr$(34)
          'Fazer a busca do orcamento
          SQL = "SELECT DISTINCTROW Cliente, Data, Orcamento "
          SQL = SQL & "From Orcamento "
          SQL = SQL & "WHERE ((Orcamento.Cliente=" + Pesq$ + ")); "
          AchaOrc SQL, Pesq$
      End If
   Case 2
      Pesq$ = InputBox("Digite a placa:", "Localização de Orçamentos")
      If Pesq$ > "" Then
        Loga "Pesquisa de Orçamento pela Placa: " & Pesq$
         Pesq$ = Chr$(34) + Trim$(Pesq$) + Chr$(34)
         
        SQL = "SELECT Carros.Modelo, Orcamento.Cliente, Orcamento.Carro, Orcamento.Orcamento, Carros.Placa "
        SQL = SQL & "from Orcamento "
        SQL = SQL & "INNER JOIN Carros ON Orcamento.Carro = Carros.Placa "
        SQL = SQL & "Where Carros.Placa = " & Pesq$
        
        AchaOrc SQL, Placa:=Pesq$
      End If
    Case 3
        RelacaoMensal
End Select
End Sub

Private Sub LocalizaPlaca(Placa As String)
Dim Selec As String
Dim CadeCli As Recordset

Selec = "SELECT Clientes.NrCli, Carros.Placa "
Selec = Selec & "FROM Carros INNER JOIN Clientes ON Carros.NrCli = Clientes.NrCli "
Selec = Selec & "WHERE Carros.Placa='" & Placa & "' "

AbreTB CadeCli, Selec, dbOpenDynaset

If CadeCli.EOF Then
   '2.7.2 Logar todas mensagens
   msgboxL "Não localizei o cliente"
Else
   Load frmClientes
   frmClientes.NrCliente = CadeCli("NrCli")
   frmClientes.nrPlaca = Placa
   frmClientes.Show
End If
End Sub

'5.1.1 Pesquisa por CPF e CNPj
Private Sub LocalizaCPF(CPF As String)
Dim Selec As String
Dim CadeCli As Recordset

Selec = "SELECT NrCli From Clientes "
Selec = Selec & "WHERE CgcCpf='" & CPF & "' "
AbreTB CadeCli, Selec, dbOpenDynaset
If CadeCli.EOF Then
   msgboxL "Não localizei o cliente"
Else
   Load frmClientes
   frmClientes.NrCliente = CadeCli("NrCli")
   frmClientes.Show
End If
End Sub

Private Function PrimeiroCli() As Integer
Dim Prime As Recordset
Dim Selec As String

Selec = "select NrCli from Clientes order by Nome "

AbreTB Prime, "Config", dbOpenDynaset

On Error GoTo SemClientes
Prime.MoveFirst
PrimeiroCli = Prime(0)
Prime.Close
Exit Function

SemClientes:
End Function

Private Sub RelacaoMensal()
'1.8.0 Relação Mensal
Dim DT         As Date
Dim DtNow      As Date
Dim SDTSqls     As String
Dim Selec      As String
Dim MesPassado As String
Dim CadeOrc    As Recordset

DT = DtINIRel()

SDTSqls = Format(DT, "dd/mm/yyyy")
SDTSqls = "#" & Mid$(SDTSqls, 4, 3) + Left$(SDTSqls, 3) + Right$(SDTSqls, 4) & "#"
MesPassado = SDTSqls

Selec = "SELECT DISTINCTROW Cliente, Data, Orcamento "
Selec = Selec & "From Orcamento "
Selec = Selec & "WHERE Data > " & MesPassado

AbreTB CadeOrc, Selec, dbOpenDynaset

'2.6.0 Mostrar a tela de lista de orçamentos mesmo quando não tem orçamentos no mes
If CadeOrc.EOF Then
    CadeOrc.Close
    DtNow = Consulta("Select max(Data) From Orcamento")
    
    If IsNull(DT) Then
        '2.7.2 Logar todas mensagens
        msgboxL "Não existem orçamentos gravados"
    End If
    DT = DtNow - 30
    SDTSqls = Format(DT, "dd/mm/yyyy")
    SDTSqls = "#" & Mid$(SDTSqls, 4, 3) + Left$(SDTSqls, 3) + Right$(SDTSqls, 4) & "#"
    MesPassado = SDTSqls
    Selec = "SELECT DISTINCTROW Cliente, Data, Orcamento "
    Selec = Selec & "From Orcamento "
    Selec = Selec & "WHERE Data > " & MesPassado
    AbreTB CadeOrc, Selec, dbOpenDynaset
    
    If CadeOrc.RecordCount < 1 Then
        Load frmOrc
        frmOrc.NrOrcamento = CadeOrc("Orcamento")
        CadeOrc.Close
        frmOrc.Show
        Exit Sub
    Else
        '2.7.2 Logar todas mensagens
        msgboxL "Não há orçamentos no ultimo mes"
    End If
Else
    DtNow = Now
End If

CadeOrc.MoveLast
   Load Orcas
    With Orcas

         '2.4.2 Observação do orçamento na relação mensal
'            .Height = 3630
'            .MSFlexGrid1.Top = 420
'            .Command1(0).Top = 2580
'            .Command1(1).Top = 2580

        .Label1(0).Visible = True
        .Label1(1).Visible = True
        .txDtINI.Visible = True
        .txDtFIM.Visible = True
        .btPesq.Visible = True
        .txDtINI.Text = Format(DT, "DD/MM/YYYY")
        .txDtFIM.Text = Format(DtNow, "DD/MM/YYYY")
        .Mensal = MesPassado
    End With
    Orcas.Show
'End If
CadeOrc.Close
End Sub

Private Sub FazRestauracao(Optional Caminho As String = "")
Dim Linha1     As String
Dim CaminhoOri As String

'3.4.8 Ajuste na restauração do Backup
CaminhoOri = CaminhoBkp & "\OrCarro.zip"
'CaminhoOri = CaminhoBkp & "OrCarro.zip"

'Fechar a base
db.Close

'Fazer um backup do arquivo atual
Caption = "Fazendo backup interno"
If frmClientesLido Then
    Unload frmClientes
End If

FileCopy "OrCarro.mdb", "OrCarro.bak"

'Descompactar o novo arquivo
Open CaminhoOri For Binary As #2
TamArqZip = LOF(2) / 1024
Close #2

If TamArqZip = 0 Then
    Kill CaminhoOri
    '2.7.2 Logar todas mensagens
    msgboxL "Não achei o arquivo de backup"
Else
    Caption = "Descompactando o arquivo"
    InicializaZip Me, txtZip
    OperZip = 1
    DesCompacta CaminhoOri, "*.*", App.Path, False
            
    'Abrir a base
    Caption = "Abrindo a Base"
    Set db = OpenDatabase(Base)
    
    Caption = "Restauração completada"
    '2.7.2 Logar todas mensagens
    msgboxL "Restauração completada"
    Caption = "Gerenciador de Orçamentos"
End If
End Sub

Private Sub MnuPagPagar_Click()
'2.4.6 Relação de totais de valores de Orçamentos
frmVlrPagPagar.Show
End Sub

Private Sub MnuProdutividade_Click()
frmProdutividade.Show
End Sub

Private Sub mnuRec_Click(Index As Integer)
'4.3.8 Correção da chamada do menu dos recibos
If Index = 4 Then
    frmReciboAvulso.Show
Else
    frmRecibo.Tipo = Index
    frmRecibo.Show
End If
'4.3.7 Tela para recibo avulso
'frmReciboAvulso.Show
'3.4.4 Vale Transporte
'frmRecibo.Tipo = Index
'frmRecibo.Show
End Sub

Private Sub MnuRelComiss_Click()
'3.8.3 Relatório de pagamentos
frmRelPagamentos.Show
End Sub

Private Sub MnuRelTot_Click(Index As Integer)
'2.9.7 Relatório de Totais Resumido
Load frmRelTotais
frmRelTotais.Funcao Index
frmRelTotais.Show
End Sub

Private Sub MnuRestDisk_Click()
FazRestauracao CaminhoBkp & "OrCarro.zip"
End Sub

Private Sub mnuTabela_Click()
'3.2.9 Nova tela de tabela de preços
TabelaPreco.Show
'3.2.7 Gravação do valor da mão de obra
'frmModeloItens.Show

End Sub

'3.7.0 Retirar preparação para tarefas dinâmicas
'Private Sub PrepTarDin_Click()
''2.7.5 Taréfas Dinâmicas
'Dim clsCLi As New clsTarefas
'
'clsCLi.PrepTarDin lbOperacao
'lbOperacao.Caption = "Adaptação realizada com sucesso"
'lbOperacao.Refresh
'End Sub

Private Sub txtZip_Change()
Dim Aux                 As Integer
Dim UltSituValue        As Integer
Dim KBytesCopiadosAntes As Integer
Dim KBytesProcessados   As Integer
Dim Valor               As Single
Dim Operacao            As String

'Indicação de progresso da compactação/descompactação por arquivo
'----------------------------------------------------------------
'Tipo de ação que esta sendo feita no momento
Operacao = TipoAção(Val(GetAction(txtZip))) & " "
'Nome do arquivo que esta sendo compactado
Operacao = Operacao & GetFileName(txtZip)
'Porcentagem de compactação do arquivo
If OperZip = 1 Then
    KBytesProcessados = GetKBytesProcessados(txtZip)
    If KBytesProcessados > 0 Then
        'SituMax
        Caption = "Descompactando.." + Format(KBytesProcessados / TamArqZip * 100, "##.##") + "%"
    End If
Else
    Aux = KBytesCopiadosAntes + GetKBytesProcessados(txtZip)
    If Aux > UltSituValue Then
        UltSituValue = Aux
    End If
    If UltSituValue > 0 Then
        Valor = UltSituValue / TamArqZip * 100
        If Valor < 100 Then
            Caption = "Compactando.." + Format(UltSituValue / TamArqZip * 100, "##.##") + "%"
        Else
            Caption = "Arquivo zip criado"
        End If
    End If
End If
DoEvents
'Debug.Print Situ.Value
End Sub

'5.0.4 Cadastro de Revisões
Private Sub MnuRev_Click(Index As Integer)
Select Case Index
    Case 0
        Load frmBusca
        frmBusca.Tipo = 3
        frmBusca.Show
    Case 1
        Load CadRevisoes
        CadRevisoes.Adicionar
        CadRevisoes.Show
    Case 2
        '5.4.7 Listagem de Revisões
        ListaRevisoes.Show
End Select
End Sub

