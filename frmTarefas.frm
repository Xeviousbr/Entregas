VERSION 5.00
Object = "{00028C4A-0000-0000-0000-000000000046}#5.0#0"; "TDBG5.OCX"
Begin VB.Form frmTarefas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Taréfas dos Mecânicos"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9675
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9675
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btExcluir 
      Caption         =   "Excluir Taréfa"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2580
      TabIndex        =   23
      Top             =   6480
      Width           =   4065
   End
   Begin VB.TextBox txTotal 
      Alignment       =   1  'Right Justify
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   6780
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txPecas 
      Alignment       =   1  'Right Justify
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   5220
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton btImprime 
      Caption         =   "Impressão"
      Height          =   375
      Left            =   8280
      TabIndex        =   18
      Top             =   780
      Width           =   1275
   End
   Begin VB.TextBox txCarros 
      Alignment       =   1  'Right Justify
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   5220
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txMec 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   900
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cbMecanico 
      Height          =   315
      ItemData        =   "frmTarefas.frx":0000
      Left            =   900
      List            =   "frmTarefas.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox txVlrBruto 
      Alignment       =   1  'Right Justify
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   5220
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txDet 
      Height          =   855
      IMEMode         =   3  'DISABLE
      Left            =   2640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      PasswordChar    =   "*"
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Top             =   5220
      Width           =   4005
   End
   Begin VB.CommandButton btAssumir 
      Caption         =   "Assumir Taréfa"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2595
      TabIndex        =   11
      Top             =   6180
      Width           =   4065
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Z:\Share\Orcarro\OrCarro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"frmTarefas.frx":0004
      Top             =   2760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton okOK 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      Top             =   360
      Width           =   1275
   End
   Begin VB.TextBox txVlrRec 
      Alignment       =   1  'Right Justify
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txVlrAssumido 
      Alignment       =   1  'Right Justify
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.Data Dados 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Z:\Share\Orcarro\OrCarro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6300
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"frmTarefas.frx":0126
      Top             =   60
      Visible         =   0   'False
      Width           =   2175
   End
   Begin TrueDBGrid50.TDBGrid Grid 
      Bindings        =   "frmTarefas.frx":0363
      Height          =   1635
      Left            =   120
      OleObjectBlob   =   "frmTarefas.frx":0377
      TabIndex        =   2
      ToolTipText     =   "Para alterar o estado da tarefa clique com BOTÃO DIREITO"
      Top             =   1260
      Width           =   9435
   End
   Begin TrueDBGrid50.TDBGrid GridLivre 
      Bindings        =   "frmTarefas.frx":415E
      Height          =   1875
      Left            =   2655
      OleObjectBlob   =   "frmTarefas.frx":4172
      TabIndex        =   9
      ToolTipText     =   "Selecione para ver a observação"
      Top             =   3300
      Width           =   4005
   End
   Begin VB.Label lbTotal 
      Caption         =   "Total: "
      Height          =   195
      Left            =   6300
      TabIndex        =   22
      Top             =   960
      Width           =   435
   End
   Begin VB.Label lbPecas 
      Caption         =   "Peças: "
      Height          =   195
      Left            =   4200
      TabIndex        =   20
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Carros atendidos:"
      Height          =   195
      Left            =   3960
      TabIndex        =   17
      Top             =   240
      Width           =   1275
   End
   Begin VB.Label lbVlrBruto 
      AutoSize        =   -1  'True
      Caption         =   "Mão de Obra: "
      Height          =   195
      Left            =   4200
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Taréfas sem mecânico"
      Height          =   195
      Index           =   3
      Left            =   2640
      TabIndex        =   10
      Top             =   3000
      Width           =   4215
   End
   Begin VB.Label lbDisp 
      Caption         =   "Não Disponível"
      Height          =   195
      Left            =   2880
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor realizado:"
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   5
      Top             =   960
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor total das taréfas:"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Mecânico: "
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "frmTarefas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'5.1.1 Permtir que escritório possa assumir e excluir tarefas
'5.1.0 Acréscimo de opções em caso do programa concluir o descancelamento da tarefa
'3.7.2 Mostrar os valores das peças dos orçamentos dos mecânicos
'3.7.0 Ajuste quanto a orçamentos com pagamento desfeito na visualização das tarefas livres
'3.6.1 Permitir o mecânico concluir a tarefa caso seja uma tarefa duplicada da placa
'3.5.1 Não excluir fisicamente Mecânico
'3.4.8 Mecanicos não devem imprimir tarefas
'3.4.0 Ajuste na gravação das tarefas pelos mecanicos
'3.3.9 Nova forma de desconcluir e des-assumir as tarefas
'3.3.8 Data que assumiu a assumiu as tarefas
'3.3.0 Impedir de ir para outro mecânico quanto já ter mostrado um na tela de tarefas
'3.3.0 Atualiza valor total do orçamento caso o recalculo mostre que estava errado
'3.2.5 Loga os valores mostrados na tarefa
'3.2.4 Combo para selecionar o mecânico na tela de tarefas
'3.1.6 Ajuste da visualização das tarefas livres
'3.1.4 Retorno da visualização do valor bruto das taréfas para o modo escritório
'3.1.4 Deixar passar só mecânicos na tela de taréfas
'3.1.2 Impedir que os mecânicos alterem os estado das tarefas pela tela de tarefas
'3.1.0 Modo Balcão
'3.0.7 Adaptação da tela de taréfas pra quando não usa comissões
'3.0.0 Mostrar o valor de comissão também para as tarefas livres, para os mecânicos
'2.9.5 Carro na tela de tarefas
'2.9.3 Mostrar as tarefas livres filtradas das tarefas se estiver no modo mecânico
'2.9.2 Em modo mecânico não mostras as tarefas livres
'2.9.2 Permitir que o modo escritório faça manutenção nas tarefas
'2.9.1 Permitir mecânicos assumirem tarefas com orçamentos já pagos
'2.9.1 Mostrar as tarefas sempre filtradas por placa quando for modo mecanico
'2.9.0 Mostrar o valor total das tarefas do mecânico, em modo escritório
'2.8.8 Mudança da crítica da liberação da comissão
'2.8.6 Data da conclusão da tarefa
'2.8.6 Scrolls para as observações na tela de tarefas
'2.8.5 Conserto da situação em que tarefas de orçamentos antigos não podiam ser concluídas
'2.8.3 Impedir de dar mensagem se teclar ENTER no botão
'2.8.1 Conserto do calculo do valor disponivel na tela de tarefas
'2.8.1 Em modo escritorio nao deve permitir a alteracao da situacao na tela de tarefas
'2.8.1 Avisar para salvar, caso saia da tela de tarefas com uma alteracao pendente
'2.8.1 Sair da tela de tarefas com ESC
'2.8.1 Melhorar a operacao de assumir a tarefa pela tela de tarefas
'2.8.0 Melhorar o log quanto as tarefas
'2.7.9 Conserto do erro ao concluir taréfa pela tela de taréfas

Option Explicit

Private Atualizand As Boolean
Private codMec     As Integer
Private sPercComis As String
Private lcPlaca    As String

'2.8.0 Melhorar o log quanto as tarefas
Private Alterado As Boolean

'2.8.3 Impedir de dar mensagem se teclar ENTER no botão
Private Momento As String

'2.8.8 Mudança da crítica da liberação da comissão
Private PercComiss As Single

'3.3.0 Impedir de ir para outro mecânico quanto já ter mostrado um na tela de tarefas
Private Mudando As Boolean

Private Sub btAssumir_Click()
Dim SQL      As String
Dim rsTarefa As Recordset

GridLivre.Col = 0

'5.1.1 Permtir que escritório possa assumir e excluir tarefas
'If btAssumir.Caption = "Excluir" Then
'    AbreTB rsTarefa, "Select Orc, Concerto From Tarefas Where ID = " & GridLivre.Text
'    ExecSql "Delete From Tarefas Where ID = " & GridLivre.Text
'    CarregaTarefasLivres
'Else

    SQL = "Update Tarefas "
    SQL = SQL & "Set Mec = " & codMec
    SQL = SQL & ", Situacao = 2 "
    SQL = SQL & " Where ID = " & GridLivre.Text
    ExecSql SQL
    btAssumir.Enabled = False
    Atualiza
    
'End If
End Sub

'5.1.1 Permtir que escritório possa assumir e excluir tarefas
Private Sub btExcluir_Click()
Dim rsTarefa As Recordset

If MsgBox("Tem certeza que deseja apagar esta tarefa", vbQuestion + vbYesNo + vbDefaultButton2, "Exclusão de tarefa") = vbYes Then
    GridLivre.Col = 0
    AbreTB rsTarefa, "Select Orc, Concerto From Tarefas Where ID = " & GridLivre.Text
    ExecSql "Delete From Tarefas Where ID = " & GridLivre.Text
    CarregaTarefasLivres
End If
End Sub

'3.4.7 Impressão das tarefas
Private Sub btImprime_Click()
Dim a%, Aux$, Cap$, Modelo$

Const Linha$ = "-----------------------------------------------------"

Cap$ = Me.Caption
Caption = Me.Caption & " realizando impressão das tarefas"
ImprBuferizada_Inicializa
If cbMecanico.ListIndex > -1 Then
    Dados.Recordset.MoveFirst
    LPRINT "Tarefas do mecanico: " & cbMecanico.Text
    LPRINT " em " & Format(Now, "DD/MM/YYYY HH:MM")
    LPRINT Linha$
    LPRINT "Placa   Carro               Concerto Assumiu Concluiu"
    LPRINT Linha$
    Do While Dados.Recordset.EOF = False
        Aux$ = Dados.Recordset.Fields("Carro").Value & " "
        Aux$ = Aux & ComplStr(Dados.Recordset.Fields("Modelo").Value, 19, " ", 0) & " "
        Aux$ = Aux & ComplStr(Dados.Recordset.Fields("Concerto").Value, 8, " ", 0) & "  "
        Aux$ = Aux & Format(Dados.Recordset.Fields("DtAssumiu").Value, "DD/MM") & "   "
        Aux$ = Aux & Format(Dados.Recordset.Fields("DtConclusao").Value, "DD/MM")
        LPRINT Aux$
        Dados.Recordset.MoveNext
    Loop
    LPRINT " "
    LPRINT Linha$
End If
LPRINT "Tarefas Livres "
Data2.Recordset.MoveFirst
LPRINT "Placa   Carro               Concerto "
LPRINT Linha$
Do While Data2.Recordset.EOF = False
    Aux$ = Data2.Recordset.Fields("Carro").Value & " "
    Modelo$ = Consulta("Select Modelo From Carros Where Placa = '" & Data2.Recordset.Fields("Carro").Value & "'")
    Aux$ = Aux & ComplStr(Modelo$, 19, " ", 0) & " "
    Aux$ = Aux & Data2.Recordset.Fields("Concerto").Value
    LPRINT Aux$
    Data2.Recordset.MoveNext
Loop
ImprBuferizada_Finaliza
Me.Caption = Cap$
End Sub

Private Sub cbMecanico_Click()
'3.3.0 Impedir de ir para outro mecânico quanto já ter mostrado um na tela de tarefas
If Mudando = False Then
    
    '3.2.4 Combo para selecionar o mecânico na tela de tarefas
    Busca
    LogaAsTarefas
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'2.8.1 Sair da tela de tarefas com ESC
If KeyAscii = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Dim SQL   As String
Dim TbMec As Recordset

'3.8.4 Adaptação da resolução
InicForm Me, SemAdapRes:=False

Dados.DatabaseName = App.Path & "\OrCarro.mdb"
Data2.DatabaseName = App.Path & "\OrCarro.mdb"

'2.8.0 Melhorar o log quanto as tarefas
Alterado = False

'3.1.0 Modo Balcão
If INI.ModoOperacao = tpMecanico Then
'2.8.1 Em modo escritorio nao deve permitir a alteracao da situacao na tela de tarefas
'If INI.Restrito = False Then

    '2.9.2 Permitir que o modo escritório faça manutenção nas tarefas
    'Grid.Columns(4).Locked = True
    
    '2.9.0 Mostrar o valor total das tarefas do mecânico, em modo escritório
    'lbVlrBruto.Visible = True
    'txVlrBruto.Visible = True
    
    '3.1.2 Impedir que os mecânicos alterem os estado das tarefas pela tela de tarefas
    Grid.Columns(5).Locked = True
    
    '3.2.4 Combo para selecionar o mecânico na tela de tarefas
    txMec.Visible = True
    cbMecanico.Visible = False
    
    '3.4.8 Mecanicos não devem imprimir tarefas
    btImprime.Enabled = False
    
    '3.7.2 Mostrar os valores das peças dos orçamentos dos mecânicos
    lbPecas.Visible = False
    txPecas.Visible = False
    lbTotal.Visible = False
    txTotal.Visible = False
    
Else
    
    '3.1.4 Retorno da visualização do valor bruto das taréfas para o modo escritório
    lbVlrBruto.Visible = True
    txVlrBruto.Visible = True
    
    '3.2.4 Combo para selecionar o mecânico na tela de tarefas
    txMec.Visible = False
    cbMecanico.Visible = True
End If

'3.5.1 Não excluir fisicamente Mecânico
SQL = "Select Nome From Mecanicos Where Oper = 0 and Nome > '' and Ativo = True Order by Nome "
'SQL = "Select Nome From Mecanicos Where Oper = 0 and Nome > '' Order by Nome "

AbreTB TbMec, SQL, dbOpenDynaset
Do While TbMec.EOF = False
    cbMecanico.AddItem TbMec.Fields("Nome")
    TbMec.MoveNext
Loop
TbMec.Close
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'2.8.1 Avisar para salvar, caso saia da tela de tarefas com uma alteracao pendente
If okOK.Enabled Then
    If MsgBox("Tem certeza que deseja sair sem salvar?", vbQuestion + vbYesNo + vbDefaultButton2, "Ha alteracoes pendentes") = vbNo Then
        Cancel = 1
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'2.8.0 Melhorar o log quanto as tarefas
If Alterado Then
    LogaAsTarefas
End If
End Sub

Private Sub Grid_ColEdit(ByVal ColIndex As Integer)
okOK.Enabled = True

'2.8.0 Melhorar o log quanto as tarefas
Alterado = True
End Sub

Private Sub Grid_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim XX&

If Button = 2 Then
    If INI.ModoOperacao = tpEscritorio Then
        AlteraTarefa
        
    '3.6.1 Permitir o mecânico concluir a tarefa caso seja uma tarefa duplicada da placa
    Else
        AlteraTarefaMec
    End If
    XX& = DoEvents()
    Sleep 100
    Busca
    
End If
End Sub

Private Sub AlteraTarefa()
'3.3.9 Nova forma de desconcluir e des-assumir as tarefas
Dim TpTarefa$, Vlr$, Placa$, Texto$, ID$, Status$, SQL$, DtAssumiu$

Grid.Col = 0
ID$ = Grid.Text
Grid.Col = 1
Vlr$ = Grid.Text
Grid.Col = 2
Placa$ = Grid.Text
Grid.Col = 4
TpTarefa$ = Grid.Text
Grid.Col = 5
DtAssumiu$ = Grid.Text
Grid.Col = 6
Status$ = Grid.Text
Texto$ = "a tarefa de " & TpTarefa & " (Placa " & Placa$ & "=" & Vlr$ & ") do mecânico " & cbMecanico.Text
If Status$ = "2" Then
    If MsgBox("Deseja retirar " & Texto$, vbYesNo + vbDefaultButton2, "Alteração da resposabilidade da tarefa") = vbNo Then
        Exit Sub
    End If
Else
    If MsgBox("Deseja desconcluir " & Texto$, vbYesNo + vbDefaultButton2, "Alteração da resposabilidade da tarefa") = vbNo Then
        Exit Sub
    End If
End If
If Status$ = "2" Then
    SQL$ = "Update Tarefas Set Mec = 0, Situacao = 1, DtAssumiu = null"
Else
    SQL$ = "Update Tarefas Set Situacao = 2, DtConclusao = null"
    If DtAssumiu$ = "" Then
        SQL$ = SQL$ & ", DtAssumiu = " & DTSqls(Format(Now, "DD/MM/YYYY"))
    End If
End If
SQL$ = SQL$ & " Where ID = " & ID$
ExecSql SQL$

'3.6.1 Permitir o mecânico concluir a tarefa caso seja uma tarefa duplicada da placa
'Busca
End Sub

Private Sub AlteraTarefaMec()
'3.6.1 Permitir o mecânico concluir a tarefa caso seja uma tarefa duplicada da placa
Dim Oper%
Dim TpTarefa$, Vlr$, Placa$, Texto$, ID$, Status$, SQL$, DtAssumiu$, DTS$, DtConcl$

Grid.Col = 0
ID$ = Grid.Text
Grid.Col = 1
Vlr$ = Grid.Text
Grid.Col = 2
Placa$ = Grid.Text
Grid.Col = 4
TpTarefa$ = Grid.Text
Grid.Col = 5
DtAssumiu$ = Grid.Text
Grid.Col = 6
Status$ = Grid.Text
Grid.Col = 8
DtConcl$ = Trim(Grid.Text)
Texto$ = "a Tarefa de " & TpTarefa & " da Placa " & Placa$ & " ( " & Vlr$ & " )"
If DtConcl$ = "" Then
    If MsgBox("Deseja CONCLUIR " & Texto$, vbYesNo + vbDefaultButton2, "Alteração do estado da tarefa") = vbNo Then
        If MsgBox("Deseja Deixar de Assumir " & Texto$, vbYesNo + vbDefaultButton2, "Alteração do estado da tarefa") = vbNo Then
            Exit Sub
        Else
            Oper% = 1   'DESassume a tarefa
        End If
    Else
        Oper% = 2       'Conclui a tarefa
    End If
Else
    If MsgBox("Deseja Desconcluir " & Texto$, vbYesNo + vbDefaultButton2, "Alteração do estado da tarefa") = vbNo Then
    
        '5.1.0 Acréscimo de opções em caso do programa concluir o descancelamento da tarefa
        If MsgBox("Deseja CONCLUIR " & Texto$, vbYesNo + vbDefaultButton2, "Alteração do estado da tarefa") = vbNo Then
            If MsgBox("Deseja Deixar de Assumir " & Texto$, vbYesNo + vbDefaultButton2, "Alteração do estado da tarefa") = vbNo Then
                Exit Sub
            Else
                Oper% = 1   'DESassume a tarefa
            End If
        Else
            Oper% = 2       'Conclui a tarefa
        End If
        
    Else
        Oper% = 3       'DESconclui a tarefa
    End If
End If

DTS$ = DTSqls(Format(Now, "DD/MM/YYYY"))
SQL$ = "Update Tarefas Set "
Select Case Oper%
    Case 1              'DESassume a tarefa
        SQL$ = SQL$ & "Situacao = 1, DtAssumiu = null, Mec = 0 "
    Case 2              'Conclui a tarefa
        SQL$ = SQL$ & "Situacao = 3, DtConclusao = " & DTS$
        If DtAssumiu$ = "" Then
            SQL$ = SQL$ & ", DtAssumiu = " & DTS$
        End If
    Case Else           'DESconclui a tarefa
        SQL$ = SQL$ & "Situacao = 2, DtConclusao = null "
End Select
SQL$ = SQL$ & " Where ID = " & ID$
ExecSql SQL$
End Sub

Private Sub GridLivre_Click()
'5.1.1 Permtir que escritório possa assumir e excluir tarefas
If cbMecanico.ListIndex > -1 Then

    '2.8.1 Melhorar a operacao de assumir a tarefa pela tela de tarefas
    btAssumir.Enabled = (GridLivre.SelBookmarks.Count < 2)
    'btAssumir.Enabled = (GridLivre.SelBookmarks.Count < 1)
    
    '5.1.1 Permtir que escritório possa assumir e excluir tarefas
    btExcluir.Enabled = btAssumir.Enabled
End If
End Sub

Private Sub GridLivre_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub GridLivre_SelChange(Cancel As Integer)
Dim sObs    As String
Dim sObsMec As String
Dim stxObs  As String

If GridLivre.SelBookmarks.Count = 1 Then
    sObs = SN(GridLivre.Columns(4).Value)
    sObsMec = SN(GridLivre.Columns(5).Value)
    If sObs > " " Then
        stxObs = "Observaçao: " & sObs & vbCrLf
    End If
    If sObsMec > " " Then
        stxObs = stxObs & "Observaçao do Mecânico: " & sObsMec
    End If
    txDet.Text = stxObs
Else
    txDet.Text = ""
End If
End Sub

Private Sub okOK_Click()
Dim X       As Long
Dim DtConcl As Date

Data2.Enabled = False
Dados.Recordset.MoveFirst
Do While Dados.Recordset.EOF = False

    '2.9.5 Carro na tela de tarefas
    DtConcl = SN(Dados.Recordset!DtConclusao, vbDate)
    'DtConcl = SN(Dados.Recordset.Fields(8).Value, vbDate)
        
    'Retornando pra Sem Mecânico
    '2.9.5 Carro na tela de tarefas
    If Dados.Recordset!Situacao = 1 Then
    '2.7.9 Conserto do erro ao concluir taréfa pela tela de taréfas
    'If Dados.Recordset.Fields(4).Value = 1 Then
    
        '2.8.6 Data da conclusão da tarefa
        ExecSql "Update Tarefas Set Mec = 0, DtConclusao = Null Where ID = " & Dados.Recordset.Fields(0)
        'ExecSql "Update Tarefas Set Mec = 0 Where ID = " & Dados.Recordset.Fields(0)
    
    '2.9.5 Carro na tela de tarefas
    'Esta como concluído
    ElseIf Dados.Recordset!Situacao = 3 Then
    '2.8.6 Data da conclusão da tarefa
    'ElseIf Dados.Recordset.Fields(4).Value = 3 Then
    
        'e esta sem data de conclusão
        If DtConcl = 0 Then
            ExecSql "Update Tarefas Set DtConclusao = Int(Now) Where ID = " & Dados.Recordset.Fields(0)
        End If
    Else
        If Not (DtConcl = 0) Then
        'Não esta como concluído, mas tem data de conclusão
            ExecSql "Update Tarefas Set DtConclusao = Null Where ID = " & Dados.Recordset.Fields(0)
        End If
    End If
    Dados.Recordset.MoveNext
Loop
Dados.Enabled = False
X = DoEvents()
Busca
End Sub

Private Sub txMec_KeyUp(KeyCode As Integer, Shift As Integer)
Dim Momento2 As String
Dim Nome     As String

If KeyCode = 13 Then

    '2.8.3 Impedir de dar mensagem se teclar ENTER no botão
    Momento2 = Format(Now, "HH:MM:SS")
    If Momento <> Momento2 Then
    
        '3.5.1 Não excluir fisicamente Mecânico
        Nome = Consulta("Select Nome From Mecanicos Where Senha = '" & txMec.Text & "' and Oper = 0 and Ativo = True")
        'Nome = Consulta("Select Nome From Mecanicos Where Senha = '" & txMec.Text & "' and Oper = 0 ")
        
        If Nome = "" Then
            MsgBox "Senha inválida", vbExclamation, "Informação do Mecânico"
            txMec.SetFocus
            Exit Sub
        End If
        cbMecanico.Visible = True
        txMec.Visible = False
        cbMecanico.Text = Nome
        cbMecanico.Locked = True
        
        Busca

        '2.8.0 Melhorar o log quanto as tarefas
        LogaAsTarefas

    End If

ElseIf KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Atualiza()
Dim SQL         As String
Dim VlrAPagar   As Variant
Dim VlrAReceber As Variant

'2.8.8 Mudança da crítica da liberação da comissão
Dim VlrFeito    As Variant
Dim VlrAssumido As Variant

'3.3.0 Atualiza valor total do orçamento caso o recalculo mostre que estava errado
Dim rsTotais    As Recordset
Dim QtdCarros   As Integer

'3.3.0 Impedir de ir para outro mecânico quanto já ter mostrado um na tela de tarefas
Dim sMec As String

'2.9.5 Carro na tela de tarefas
SQL = "SELECT Tarefas.ID, Tarefas.Vlr * " & sPercComis & " as Vlr , Orcamento.Carro, Carros.Modelo, tpConcertos.concerto, "
'2.8.6 Data da conclusão da tarefa
'SQL = "SELECT Tarefas.ID, Tarefas.Vlr * " & sPercComis & " as Vlr , Orcamento.Carro, tpConcertos.concerto, "

'3.3.8 Data que assumiu a assumiu as tarefas
SQL = SQL & " Tarefas.DtAssumiu, "

SQL = SQL & "Tarefas.Situacao, Orcamento.Orcamento, Orcamento.Obs, Orcamento.ObsMec, Tarefas.DtConclusao, Tarefas.Pago "
SQL = SQL & "FROM (((Tarefas "
SQL = SQL & "INNER JOIN Mecanicos ON Tarefas.Mec = Mecanicos.codi) "
SQL = SQL & "INNER JOIN Orcamento ON Tarefas.Orc = Orcamento.Orcamento) "
SQL = SQL & "INNER JOIN tpConcertos ON Tarefas.concerto = tpConcertos.tipo) "

'2.9.5 Carro na tela de tarefas
SQL = SQL & "INNER JOIN Carros ON Orcamento.Carro = Carros.Placa "

'3.2.4 Combo para selecionar o mecânico na tela de tarefas
SQL = SQL & "WHERE Mecanicos.Nome = '" & cbMecanico.Text & "' "
'SQL = SQL & "WHERE Mecanicos.Senha = '" & txMec.Text & "' "

SQL = SQL & "and Tarefas.Situacao > 0 "
SQL = SQL & "and Tarefas.Pago is null "

'2.9.1 Mostrar as tarefas sempre filtradas por placa quando for modo mecanico
If Placa > "" Then
    SQL = SQL & "and Orcamento.Carro = '" & Placa & "'"
End If

SQL = SQL & " ORDER BY Orcamento.Orcamento, Tarefas.id "
Dados.RecordSource = SQL
Dados.Enabled = True
Dados.Refresh
SQL = "SELECT Sum(Tarefas.Vlr) * " & sPercComis & " AS SomaDeVlr "
SQL = SQL & "FROM (((Tarefas INNER JOIN Mecanicos ON Tarefas.Mec = Mecanicos.codi) INNER JOIN Orcamento ON Tarefas.Orc = Orcamento.Orcamento) INNER JOIN tpConcertos ON Tarefas.concerto = tpConcertos.tipo) "

'3.2.4 Combo para selecionar o mecânico na tela de tarefas
SQL = SQL & "WHERE Mecanicos.Nome = '" & cbMecanico.Text & "' AND Tarefas.Pago Is Null "
'SQL = SQL & "WHERE Mecanicos.Senha = '" & txMec.Text & "' AND Tarefas.Pago Is Null "

'2.8.8 Mudança da crítica da liberação da comissão
VlrAssumido = Consulta(SQL)
MostraValor txVlrAssumido, VlrAssumido
'VlrAPagar = Consulta(SQL)
'MostraValor txVlrAPagar, VlrAPagar

'3.3.0 Atualiza valor total do orçamento caso o recalculo mostre que estava errado
SQL = "SELECT Sum(Tarefas.Vlr) AS Soma, Count(*) as Quant "
'2.8.8 Mudança da crítica da liberação da comissão
'SQL = "SELECT Sum(Tarefas.Vlr) AS Soma "
'SQL = "SELECT Sum(Tarefas.Vlr) * " & sPercComis & " AS SomaDeVlr "

SQL = SQL & ", Sum([Orcamento].[Total]) As Tot "

SQL = SQL & "FROM (((Tarefas INNER JOIN Mecanicos ON Tarefas.Mec = Mecanicos.codi) INNER JOIN Orcamento ON Tarefas.Orc = Orcamento.Orcamento) INNER JOIN tpConcertos ON Tarefas.concerto = tpConcertos.tipo) "

'3.2.4 Combo para selecionar o mecânico na tela de tarefas
SQL = SQL & "WHERE Mecanicos.Nome = '" & cbMecanico.Text & "' AND Tarefas.Pago Is Null "
'SQL = SQL & "WHERE Mecanicos.Senha = '" & txMec.Text & "' AND Tarefas.Pago Is Null "

'2.8.8 Mudança da crítica da liberação da comissão
SQL = SQL & "AND Tarefas.Situacao = 3 "

'3.3.0 Atualiza valor total do orçamento caso o recalculo mostre que estava errado
AbreTB rsTotais, SQL
VlrFeito = rsTotais!Soma
QtdCarros = rsTotais!Quant
txCarros.Text = Trim(STR(QtdCarros))
'VlrFeito = Consulta(SQL)

'3.1.0 Modo Balcão
If INI.ModoOperacao = tpEscritorio Then
'2.9.0 Mostrar o valor total das tarefas do mecânico, em modo escritório
'If INI.Restrito = False Then

    MostraValor txVlrBruto, VlrFeito
    
    '3.7.2 Mostrar os valores das peças dos orçamentos dos mecânicos
    MostraValor txTotal, rsTotais!Tot
    MostraValor txPecas, (rsTotais!Tot - VlrFeito)
    
    '3.3.0 Impedir de ir para outro mecânico quanto já ter mostrado um na tela de tarefas
    Mudando = True
    sMec = cbMecanico.Text
    cbMecanico.Clear
    cbMecanico.AddItem sMec
    cbMecanico.ListIndex = 0
    Mudando = False
    
End If

VlrAPagar = VlrFeito * PercComiss
MostraValor txVlrRec, VlrAPagar
'VlrAPagar = Consulta(SQL)
'MostraValor txVlrAPagar, VlrAPagar
'SQL = "SELECT Sum(Tarefas.Vlr) * " & sPercComis & " AS SomaDeVlr "
'SQL = SQL & "FROM (((Tarefas INNER JOIN Mecanicos ON Tarefas.Mec = Mecanicos.codi) INNER JOIN Orcamento ON Tarefas.Orc = Orcamento.Orcamento) INNER JOIN tpConcertos ON Tarefas.concerto = tpConcertos.tipo) "
'
''2.8.1 Conserto do calculo do valor disponivel na tela de tarefas
'SQL = SQL & "WHERE Mecanicos.Senha = '" & txMec.Text & "' AND Tarefas.Pago Is Null and Tarefas.Situacao = 3 "
''SQL = SQL & "WHERE Mecanicos.Senha = '" & txMec.Text & "' AND Tarefas.Pago Is Null and Tarefas.Situacao = 2 "
'VlrAReceber = Consulta(SQL)
'MostraValor txVlrRec, VlrAReceber

'3.3.0 Atualiza valor total do orçamento caso o recalculo mostre que estava errado
If VlrFeito > INI.VlrGatComiss And QtdCarros > INI.QtdCarrComiss Then
'2.8.8 Mudança da crítica da liberação da comissão
'If VlrFeito > INI.VlrGatComiss Then
'If VlrAReceber > INI.VlrGatComiss Then

    lbDisp.Caption = "Disponível"
    lbDisp.FontBold = True
Else
    lbDisp.Caption = "indisponível"
    lbDisp.FontBold = False
End If

'3.2.5 Loga os valores mostrados na tarefa
Loga "cbMecanico: " & cbMecanico.Text
Loga "Valor total das taréfas: " & txVlrAssumido.Text
Loga "Valor realizado: " & txVlrRec.Text & " " & lbDisp.Caption
Loga "Valor Bruto: " & txVlrBruto.Text

CarregaTarefasLivres
End Sub

Private Sub LogaAsTarefas()
Dim SQL As String

'2.8.0 Melhorar o log quanto as tarefas
If INI.Log Then
        
    '3.2.4 Combo para selecionar o mecânico na tela de tarefas
    SQL = "Select tpConcertos.concerto, Tarefas.Vlr, tpSituacao.situacao, '" & cbMecanico.Text & "' as Nome, Tarefas.Pago, Orcamento.Carro as Placa "
    'SQL = "Select tpConcertos.concerto, Tarefas.Vlr, tpSituacao.situacao, '" & lbMec.Caption & "' as Nome, Tarefas.Pago, Orcamento.Carro as Placa "
    
    SQL = SQL & "From Tarefas, tpConcertos, tpSituacao, Orcamento "
    SQL = SQL & "Where Tarefas.Mec = " & codMec
    SQL = SQL & " and tpConcertos.tipo = Tarefas.concerto"
    SQL = SQL & " and tpSituacao.tipo = Tarefas.Situacao"
    SQL = SQL & " and Orcamento.Orcamento = Tarefas.Orc"
    LogaTarefas SQL, True
End If
End Sub

Private Sub CarregaTarefasLivres()
Dim txCmpVlr As String
Dim SQL      As String

'3.1.0 Modo Balcão
If INI.ModoOperacao = tpEscritorio Then
'If INI.Restrito = False Then

    txCmpVlr = " Tarefas.Vlr "
Else
    txCmpVlr = " Tarefas.Vlr * " & sPercComis & " as Vlr "
End If

Data2.Enabled = False

'3.0.0 Mostrar o valor de comissão também para as tarefas livres, para os mecânicos
SQL = "SELECT Tarefas.id, " & txCmpVlr & ", Orcamento.Carro, TpConcertos.concerto, Orcamento.Orcamento, "
'2.9.0 Impedir que orçamentos pagos apareçam na grid de livres
'SQL = "SELECT Tarefas.id, Tarefas.Vlr, Orcamento.Carro, TpConcertos.concerto, Orcamento.Orcamento, "

SQL = SQL & "Orcamento.Obs , Orcamento.ObsMec "
SQL = SQL & "from Tarefas, Orcamento, TpConcertos "
SQL = SQL & "Where Tarefas.id > " & INI.Orc1
SQL = SQL & " AND Orcamento.Orcamento=Tarefas.Orc"
SQL = SQL & " AND Tarefas.Mec=0 "
SQL = SQL & " AND TpConcertos.tipo=Tarefas.concerto "

'2.9.1 Permitir mecânicos assumirem tarefas com orçamentos já pagos
'SQL = SQL & " AND Orcamento.Pagamento=0 "

'2.9.1 Mostrar as tarefas sempre filtradas por placa quando for modo mecanico
If Placa > "" Then
    SQL = SQL & "and Orcamento.Carro = '" & Placa & "' "
Else
    '3.7.0 Ajuste quanto a orçamentos com pagamento desfeito na visualização das tarefas livres
    SQL = SQL & "AND (Orcamento.Pagamento=0 or Orcamento.Pagamento is null) "
    '2.9.1 Permitir mecânicos assumirem tarefas com orçamentos já pagos
    'SQL = SQL & " AND Orcamento.Pagamento=0 "
End If

'3.1.6 Ajuste da visualização das tarefas livres
SQL = SQL & " and Tarefas.Pago Is Null "

SQL = SQL & "ORDER BY Orcamento.Orcamento, Tarefas.id "

Data2.RecordSource = SQL
Data2.Enabled = True
Data2.Refresh
End Sub

Private Sub Busca()
Dim Errado      As Boolean
Dim TbMec       As Recordset
Dim VlrAPagar   As Currency
Dim VlrAReceber As Currency
Dim SQL         As String

'3.5.1 Não excluir fisicamente Mecânico
SQL = "Select codi, Nome, PercComiss From Mecanicos Where Nome = '" & cbMecanico.Text & "' and Ativo = True "
'SQL = "Select codi, Nome, PercComiss From Mecanicos Where Nome = '" & cbMecanico.Text & "'"

'3.1.4 Deixar passar só mecânicos na tela de taréfas
SQL = SQL & " and Oper = 0 "

AbreTB TbMec, SQL
If TbMec.EOF = False Then

    '3.2.4 Combo para selecionar o mecânico na tela de tarefas
    'lbMec.Caption = TbMec!Nome
    
    '2.8.8 Mudança da crítica da liberação da comissão
    PercComiss = SN(TbMec!PercComiss / 100, vbSingle)
    sPercComis = VlrSql(STR(PercComiss))
    'sPercComis = VlrSql(SN(TbMec!PercComiss / 100, vbSingle))
    
    codMec = TbMec!codi
    Atualiza
    lbDisp.Visible = True
    okOK.Enabled = False
    
Else
    MsgBox "Mecânico não identificado", vbCritical, "OrCarro"
End If
TbMec.Close
End Sub

Public Property Get Placa() As String
Placa = lcPlaca
End Property

Public Property Let Placa(ByVal vNewValue As String)
'2.8.3 Impedir de dar mensagem se teclar ENTER no botão
Momento = Format(Now, "HH:MM:SS")

'3.1.0 Modo Balcão
If vNewValue = "" And INI.ModoOperacao = tpMecanico Then
'2.9.3 Mostrar as tarefas livres filtradas das tarefas se estiver no modo mecânico
'If vNewValue = "" And INI.Restrito = True Then
'2.9.2 Em modo mecânico não mostras as tarefas livres
'If vNewValue > "" Then

    Me.Height = 3450
    Label1(3).Visible = False
    GridLivre.Visible = False
    txDet.Visible = False
    btAssumir.Visible = False
End If
lcPlaca = vNewValue

'3.1.0 Modo Balcão
If INI.ModoOperacao = tpEscritorio Then
'If INI.Restrito = False Then

    GridLivre.Columns(1).Caption = "Valor"
    CarregaTarefasLivres
    
    '5.1.1 Permtir que escritório possa assumir e excluir tarefas
    '3.0.7 Adaptação da tela de taréfas pra quando não usa comissões
    'If INI.UtComissoes = 1 Then
    '    btAssumir.Caption = "Excluir"
    'End If
    
End If
End Property

