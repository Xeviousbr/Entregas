VERSION 5.00
Begin VB.Form frmProdutividade 
   Caption         =   "Valores Pagos e a Pagar"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   6570
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cbTarefas 
      Height          =   315
      ItemData        =   "frmProdutividade.frx":0000
      Left            =   4740
      List            =   "frmProdutividade.frx":000A
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   60
      Width           =   1755
   End
   Begin VB.ComboBox cbMecanico 
      Height          =   315
      ItemData        =   "frmProdutividade.frx":0021
      Left            =   900
      List            =   "frmProdutividade.frx":0023
      Sorted          =   -1  'True
      TabIndex        =   19
      Top             =   60
      Width           =   2415
   End
   Begin VB.TextBox txVlrPg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas concluídas e pagas no perídio"
      Top             =   1140
      Width           =   915
   End
   Begin VB.TextBox txDtFIM 
      Height          =   285
      Left            =   2820
      MaxLength       =   20
      TabIndex        =   16
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txVlrPg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas concluídas e pagas no perídio"
      Top             =   1140
      Width           =   915
   End
   Begin VB.TextBox txVlrPg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas concluídas e pagas no perídio"
      Top             =   1140
      Width           =   915
   End
   Begin VB.TextBox txVlrPg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas concluídas e pagas no perídio"
      Top             =   1140
      Width           =   915
   End
   Begin VB.TextBox txVlrPg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas concluídas e pagas no perídio"
      Top             =   1140
      Width           =   915
   End
   Begin VB.TextBox txVlrPg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas concluídas e pagas no perídio"
      Top             =   1140
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   315
      Left            =   5460
      TabIndex        =   9
      Top             =   480
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Consultar"
      Height          =   315
      Left            =   3900
      TabIndex        =   2
      Top             =   480
      Width           =   1035
   End
   Begin VB.TextBox txDtINI 
      Height          =   285
      Left            =   1020
      MaxLength       =   20
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tarefas:"
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   21
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Mecânico:"
      Height          =   255
      Index           =   11
      Left            =   60
      TabIndex        =   20
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Peças"
      Height          =   195
      Index           =   9
      Left            =   720
      TabIndex        =   18
      Top             =   840
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Data Final:"
      Height          =   195
      Index           =   8
      Left            =   2040
      TabIndex        =   15
      Top             =   540
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Eletricidade"
      Height          =   195
      Index           =   7
      Left            =   5520
      TabIndex        =   8
      Top             =   840
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Pintura"
      Height          =   195
      Index           =   6
      Left            =   4560
      TabIndex        =   7
      Top             =   840
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Guarnição"
      Height          =   195
      Index           =   5
      Left            =   3600
      TabIndex        =   6
      Top             =   840
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Chapeação"
      Height          =   195
      Index           =   4
      Left            =   2640
      TabIndex        =   5
      Top             =   840
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Mecânica"
      Height          =   195
      Index           =   3
      Left            =   1680
      TabIndex        =   4
      Top             =   840
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Valores"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   1140
      Width           =   525
   End
   Begin VB.Label Label1 
      Caption         =   "Data Inicial:"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   855
   End
End
Attribute VB_Name = "frmProdutividade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'4.2.6 Acabamento da Produtividade
'4.2.0 Tela de produtividade traz agora só os mecânicos

Option Explicit

Private Sub Command1_Click()
Dim a         As Integer
Dim Esse      As Long
Dim NrOrc     As Long
Dim sData     As String
Dim SQL       As String
Dim mensConlc As String
Dim Tarefa(5) As Currency
Dim Tarefas   As Currency
Dim Pecas     As Currency
Dim TotalOrc  As Currency
Dim Dados     As Recordset

If IsDate(txDtINI.Text) = False Or IsDate(txDtFIM.Text) = False Then
    msgboxL "Data Inválida"
Else
    sData = DTSqls(txDtINI.Text) & " and " & DTSqls(txDtFIM.Text, True)
        
    SQL = "SELECT First(ORC.Total) AS Total, ORC.Orcamento, Sum(Tarefas.Vlr) AS TT, Tarefas.concerto "
    SQL = SQL & "FROM Tarefas, Orcamento ORC "
    SQL = SQL & "WHERE ORC.Data Between " & sData
    SQL = SQL & " and Tarefas.Orc = ORC.Orcamento "
    If cbTarefas.ListIndex = 0 Then
        SQL = SQL & "and Tarefas.Situacao = 3 "
        mensConlc = "com tarefas concluidas "
    End If
    
    '4.2.6 Acabamento da Produtividade
    If cbMecanico.ListIndex > -1 Then
        SQL = SQL & " and Tarefas.Mec = " & Consulta("Select codi from mecanicos where nome = " & FA(cbMecanico.Text))
    End If
    
    SQL = SQL & " GROUP BY ORC.Orcamento, Tarefas.concerto "
    SQL = SQL & "Order BY ORC.Orcamento"
    
    AbreTB Dados, SQL, dbOpenSnapshot
    If Dados.EOF Then
        msgboxL "Não há orçamentos " & mensConlc & "neste intervalo de dadoas"
    Else
        Esse = Dados!Orcamento
        Do While Dados.EOF = False
            Tarefa(Dados!concerto) = Tarefa(Dados!concerto) + Dados("TT")
            NrOrc = Dados!Orcamento
            If Esse = NrOrc Then
                Tarefas = Tarefas + Dados("TT")
'                TotalOrc = Dados!Total
            Else
                
                'A cada registro
                'Somar os valores de tarefas
                'Já somar nas matrizes de tipos de tarefas
                
                'A cada novo nr de orçamento
                'Deduzir os valores de peças
                'Acumular os valores de totais de peças
                        
                Pecas = Dados!Total - Tarefas
'                Pecas = TotalOrc - Tarefas
                
                Tarefa(5) = Tarefa(5) + Pecas
                Tarefas = 0
                'Tarefas = Dados("TT")
                
                Esse = NrOrc
            End If
            Dados.MoveNext
        Loop
        For a = 0 To 5
            MostraValor txVlrPg(a), Tarefa(a)
        Next
    End If
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim SQL    As String
Dim tbMecs As Recordset
Dim DT     As Date

InicForm Me
DT = Consulta("SELECT Max(Orcamento.Data) FROM Orcamento")
txDtINI.Text = Format(DT - 30, "dd/mm/yyyy")
txDtFIM.Text = Format(DT, "dd/mm/yyyy")

'4.2.0 Tela de produtividade traz agora só os mecânicos
SQL = "Select Nome From Mecanicos Where Ativo = True And Oper = 0 and Nome > '' " & SQL & " Order by Nome "

AbreTB tbMecs, SQL, dbOpenDynaset
cbMecanico.AddItem "TODOS"
Do While tbMecs.EOF = False
    cbMecanico.AddItem tbMecs!Nome
    tbMecs.MoveNext
Loop
tbMecs.Close
cbMecanico.Text = "TODOS"

cbTarefas.ListIndex = 0
End Sub

'Private Sub btImpres_Click()
''2.4.8 Impressão da consulta de valores pagos e a pagar
'Load RelPagosPagar
'RelPagosPagar.txVlrPg1.Caption = txVlrPg(0).Text
'RelPagosPagar.txVlrPg2.Caption = txVlrPg(1).Text
'RelPagosPagar.txVlrPg3.Caption = txVlrPg(2).Text
'RelPagosPagar.txVlrPg4.Caption = txVlrPg(3).Text
'RelPagosPagar.txVlrPg5.Caption = txVlrPg(4).Text
'RelPagosPagar.txVlrNpgNoInt1.Caption = txVlrNpgNoInt(0).Text
'RelPagosPagar.txVlrNpgNoInt2.Caption = txVlrNpgNoInt(1).Text
'RelPagosPagar.txVlrNpgNoInt3.Caption = txVlrNpgNoInt(2).Text
'RelPagosPagar.txVlrNpgNoInt4.Caption = txVlrNpgNoInt(3).Text
'RelPagosPagar.txVlrNpgNoInt5.Caption = txVlrNpgNoInt(4).Text
'RelPagosPagar.txVlraPg1.Caption = txVlraPg(0).Text
'RelPagosPagar.txVlraPg2.Caption = txVlraPg(1).Text
'RelPagosPagar.txVlraPg3.Caption = txVlraPg(2).Text
'RelPagosPagar.txVlraPg4.Caption = txVlraPg(3).Text
'RelPagosPagar.txVlraPg5.Caption = txVlraPg(4).Text
'RelPagosPagar.lbDt.Caption = "Data inicial " & txDtINI.Text & " final " & txDtFIM.Text
'RelPagosPagar.Show
'Unload Me
'End Sub
'
'Private Sub Command1_Click()
'Dim SQL           As String
'Dim conAPagar     As Recordset
'Dim conPagas      As Recordset
'Dim conNaoPgnoInt As Recordset
'
''2.9.0 Ajuste dos valores pagos e a pagar quanto as tarefas dinamicas
'Dim a             As Integer
'Dim sData         As String
'
'If IsDate(txDtINI.Text) = False Or IsDate(txDtFIM.Text) = False Then
'    '2.7.2 Logar todas mensagens
'    msgboxL "Data Inválida"
'Else
'
'    '2.9.0 Ajuste dos valores pagos e a pagar quanto as tarefas dinamicas
'    For a = 0 To 4
'        txVlrPg(a).Text = ""
'        txVlrNpgNoInt(a).Text = ""
'        txVlraPg(a).Text = ""
'    Next
'    sData = DTSqls(txDtINI.Text) & " and " & DTSqls(txDtFIM.Text, True)
'
'    'Valores Pagos
'    SQL = "SELECT Sum(Tarefas.Vlr) AS SomaDeVlr, tpConcertos.tipo "
'    SQL = SQL & "from Tarefas, tpConcertos "
'    SQL = SQL & "Where Tarefas.Pago BetWeen " & sData
'    SQL = SQL & " and tpConcertos.tipo = Tarefas.concerto "
'    SQL = SQL & "GROUP BY tpConcertos.tipo "
'    AbreTB conPagas, SQL, dbOpenSnapshot
'    If conPagas.EOF = False Then
'        Do While conPagas.EOF = False
'            MostraValor txVlrPg(conPagas!Tipo), conPagas!SomaDeVlr
'            conPagas.MoveNext
'        Loop
'    End If
'
'    'Valores Concluídos mas não pagos
'    SQL = "SELECT Sum(Tarefas.Vlr) AS SomaDeVlr, tpConcertos.tipo "
'    SQL = SQL & "from Tarefas, tpConcertos "
'    SQL = SQL & "Where Tarefas.DtConclusao BetWeen " & sData
'    SQL = SQL & " and Tarefas.Pago is Null "
'    SQL = SQL & " and tpConcertos.tipo = Tarefas.concerto "
'    SQL = SQL & "GROUP BY tpConcertos.tipo "
'    AbreTB conNaoPgnoInt, SQL, dbOpenSnapshot
'    If conNaoPgnoInt.EOF = False Then
'        Do While conNaoPgnoInt.EOF = False
'            MostraValor txVlrNpgNoInt(conNaoPgnoInt!Tipo), conNaoPgnoInt!SomaDeVlr
'            conNaoPgnoInt.MoveNext
'        Loop
'    End If
'
'    'Assumidos mas não concluídos
'    SQL = "SELECT Sum(Tarefas.Vlr) AS SomaDeVlr, tpConcertos.tipo "
'    SQL = SQL & "from Tarefas, tpConcertos "
'    SQL = SQL & "Where Tarefas.Situacao = 2 "
'    SQL = SQL & "and tpConcertos.tipo = Tarefas.concerto "
'    SQL = SQL & "and tarefas.Orc in "
'    SQL = SQL & "( "
'    SQL = SQL & "Select Orcamento "
'    SQL = SQL & "from Orcamento "
'    SQL = SQL & "Where Pagamento BetWeen " & sData
'    SQL = SQL & "Order By Orcamento "
'    SQL = SQL & ") "
'    SQL = SQL & "GROUP BY tpConcertos.tipo"
'    AbreTB conAPagar, SQL, dbOpenSnapshot
'    If conAPagar.EOF = False Then
'        Do While conAPagar.EOF = False
'            MostraValor txVlraPg(conAPagar!Tipo), conAPagar!SomaDeVlr
'            conAPagar.MoveNext
'        Loop
'    End If
'
''    SQL = "SELECT Sum(Mecanica) AS SomaDeMecanica, Sum(Pintura) AS SomaDePintura, "
''    SQL = SQL & "Sum(Chapeação) AS SomaDeChapeação, Sum(Eletricidade) AS SomaDeEletricidade, "
''    SQL = SQL & "Sum(Guarnição) As SomaDeGuarnição "
''    SQL = SQL & "from Orcamento "
''    '2.4.8 Data final de intervalo no relatório "valores não pagos no intervalo"
''    SQL = SQL & "Where Pagamento BetWeen " & DTSqls(txDtINI.Text) & " and " & DTSqls(txDtFIM.Text, True)
''    AbreTB conPagas, SQL, dbOpenSnapshot
''
''    '2.4.7 Ajuste nos campos dos totais de valores
''    MostraValor txVlrPg(0), conPagas!SomaDeMecanica
''    MostraValor txVlrPg(1), conPagas!SomaDeChapeação
''    MostraValor txVlrPg(2), conPagas!SomaDeGuarnição
''    MostraValor txVlrPg(3), conPagas!SomaDePintura
''    MostraValor txVlrPg(4), conPagas!SomaDeEletricidade
''
''    '2.4.8 Mostrar valores não pagos no intervalo
''    SQL = "SELECT Sum(Mecanica) AS SomaDeMecanica, Sum(Pintura) AS SomaDePintura, "
''    SQL = SQL & "Sum(Chapeação) AS SomaDeChapeação, Sum(Eletricidade) AS SomaDeEletricidade, "
''    SQL = SQL & "Sum(Guarnição) As SomaDeGuarnição "
''    SQL = SQL & "from Orcamento "
''    SQL = SQL & "Where Data BetWeen " & DTSqls(txDtINI.Text) & " and " & DTSqls(txDtFIM.Text, True)
''    SQL = SQL & " and (Pagamento Is Null Or Pagamento < #1/1/2000#) "
''    AbreTB conNaoPgnoInt, SQL, dbOpenSnapshot
''    MostraValor txVlrNpgNoInt(0), conNaoPgnoInt!SomaDeMecanica
''    MostraValor txVlrNpgNoInt(1), conNaoPgnoInt!SomaDeChapeação
''    MostraValor txVlrNpgNoInt(2), conNaoPgnoInt!SomaDeGuarnição
''    MostraValor txVlrNpgNoInt(3), conNaoPgnoInt!SomaDePintura
''    MostraValor txVlrNpgNoInt(4), conNaoPgnoInt!SomaDeEletricidade
''
''    SQL = "SELECT Sum(Mecanica) AS SomaDeMecanica, Sum(Pintura) AS SomaDePintura, "
''    SQL = SQL & "Sum(Chapeação) AS SomaDeChapeação, Sum(Eletricidade) AS SomaDeEletricidade, "
''    SQL = SQL & "Sum(Guarnição) As SomaDeGuarnição "
''    SQL = SQL & "from Orcamento "
''    SQL = SQL & "Where Pagamento Is Null Or Pagamento < #1/1/2000# "
''    AbreTB conAPagar, SQL, dbOpenSnapshot
''    '2.4.8 Mostrar valores não pagos no intervalo
''    MostraValor txVlraPg(0), conAPagar!SomaDeMecanica - conNaoPgnoInt!SomaDeMecanica
''    MostraValor txVlraPg(1), conAPagar!SomaDeChapeação - conNaoPgnoInt!SomaDeChapeação
''    MostraValor txVlraPg(2), conAPagar!SomaDeGuarnição - conNaoPgnoInt!SomaDeGuarnição
''    MostraValor txVlraPg(3), conAPagar!SomaDePintura - conNaoPgnoInt!SomaDePintura
''    MostraValor txVlraPg(4), conAPagar!SomaDeEletricidade - conNaoPgnoInt!SomaDeEletricidade
'
'    '2.4.8 Impressão da consulta de valores pagos e a pagar
'    btImpres.Enabled = True
'End If
'End Sub
'
'Private Sub Command2_Click()
'Unload Me
'End Sub
'
'Private Sub Form_Load()
'Dim DT As Date
'
'InicForm Me
'DT = DtINIRel()
'txDtINI.Text = Format(DT, "dd/mm/yyyy")
'
''2.4.8 Data final de intervalo no relatório "valores não pagos no intervalo"
'txDtFIM.Text = Format(Now, "dd/mm/yyyy")
'End Sub
