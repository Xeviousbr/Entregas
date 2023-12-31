VERSION 5.00
Begin VB.Form frmVlrPagPagar 
   Caption         =   "Valores Pagos e a Pagar"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   7335
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btImpres 
      Caption         =   "Impressão"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5520
      TabIndex        =   29
      Top             =   60
      Width           =   1035
   End
   Begin VB.TextBox txVlrNpgNoInt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   6300
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas concluídas no período mas ainda não pagas"
      Top             =   1020
      Width           =   915
   End
   Begin VB.TextBox txVlrNpgNoInt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas concluídas no período mas ainda não pagas"
      Top             =   1020
      Width           =   915
   End
   Begin VB.TextBox txVlrNpgNoInt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   4380
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas concluídas no período mas ainda não pagas"
      Top             =   1020
      Width           =   915
   End
   Begin VB.TextBox txVlrNpgNoInt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   3420
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas concluídas no período mas ainda não pagas"
      Top             =   1020
      Width           =   915
   End
   Begin VB.TextBox txVlrNpgNoInt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   2460
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas concluídas no período mas ainda não pagas"
      Top             =   1020
      Width           =   915
   End
   Begin VB.TextBox txDtFIM 
      Height          =   285
      Left            =   3300
      MaxLength       =   20
      TabIndex        =   22
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txVlraPg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   6300
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas assumidas no período mas não concluídas"
      Top             =   1320
      Width           =   915
   End
   Begin VB.TextBox txVlraPg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas assumidas no período mas não concluídas"
      Top             =   1320
      Width           =   915
   End
   Begin VB.TextBox txVlraPg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   4380
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas assumidas no período mas não concluídas"
      Top             =   1320
      Width           =   915
   End
   Begin VB.TextBox txVlraPg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   3420
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas assumidas no período mas não concluídas"
      Top             =   1320
      Width           =   915
   End
   Begin VB.TextBox txVlraPg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   2460
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas assumidas no período mas não concluídas"
      Top             =   1320
      Width           =   915
   End
   Begin VB.TextBox txVlrPg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   6300
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas concluídas e pagas no perídio"
      Top             =   720
      Width           =   915
   End
   Begin VB.TextBox txVlrPg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas concluídas e pagas no perídio"
      Top             =   720
      Width           =   915
   End
   Begin VB.TextBox txVlrPg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   4380
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas concluídas e pagas no perídio"
      Top             =   720
      Width           =   915
   End
   Begin VB.TextBox txVlrPg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   3420
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas concluídas e pagas no perídio"
      Top             =   720
      Width           =   915
   End
   Begin VB.TextBox txVlrPg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   2460
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Tarefas concluídas e pagas no perídio"
      Top             =   720
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   315
      Left            =   3150
      TabIndex        =   10
      Top             =   1680
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Consultar"
      Height          =   315
      Left            =   4380
      TabIndex        =   2
      Top             =   60
      Width           =   1035
   End
   Begin VB.TextBox txDtINI 
      Height          =   285
      Left            =   1500
      MaxLength       =   20
      TabIndex        =   1
      Top             =   60
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Valores Não pagos no intervalo"
      Height          =   195
      Index           =   10
      Left            =   120
      TabIndex        =   23
      Top             =   1020
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Data Final:"
      Height          =   195
      Index           =   8
      Left            =   2520
      TabIndex        =   21
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Eletricidade"
      Height          =   195
      Index           =   7
      Left            =   6300
      TabIndex        =   9
      Top             =   420
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Pintura"
      Height          =   195
      Index           =   6
      Left            =   5580
      TabIndex        =   8
      Top             =   420
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Guarnição"
      Height          =   195
      Index           =   5
      Left            =   4500
      TabIndex        =   7
      Top             =   420
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Chapeação"
      Height          =   195
      Index           =   4
      Left            =   3480
      TabIndex        =   6
      Top             =   420
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Mecânica"
      Height          =   195
      Index           =   3
      Left            =   2580
      TabIndex        =   5
      Top             =   420
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Valores a serem pagos restantes"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Valores Pagos"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Data Inicial:"
      Height          =   195
      Index           =   0
      Left            =   540
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmVlrPagPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2.9.0 Ajuste dos valores pagos e a pagar quanto as tarefas dinamicas
'2.7.2 Logar todas mensagens
'2.4.8 Impressão da consulta de valores pagos e a pagar
'2.4.8 Mostrar valores não pagos no intervalo
'2.4.8 Data final de intervalo no relatório "valores não pagos no intervalo"
'2.4.7 Ajuste nos campos dos totais de valores
'2.4.6 Relação de totais de valores de Orçamentos

Private Sub btImpres_Click()
'2.4.8 Impressão da consulta de valores pagos e a pagar
Load RelPagosPagar
RelPagosPagar.txVlrPg1.Caption = txVlrPg(0).Text
RelPagosPagar.txVlrPg2.Caption = txVlrPg(1).Text
RelPagosPagar.txVlrPg3.Caption = txVlrPg(2).Text
RelPagosPagar.txVlrPg4.Caption = txVlrPg(3).Text
RelPagosPagar.txVlrPg5.Caption = txVlrPg(4).Text
RelPagosPagar.txVlrNpgNoInt1.Caption = txVlrNpgNoInt(0).Text
RelPagosPagar.txVlrNpgNoInt2.Caption = txVlrNpgNoInt(1).Text
RelPagosPagar.txVlrNpgNoInt3.Caption = txVlrNpgNoInt(2).Text
RelPagosPagar.txVlrNpgNoInt4.Caption = txVlrNpgNoInt(3).Text
RelPagosPagar.txVlrNpgNoInt5.Caption = txVlrNpgNoInt(4).Text
RelPagosPagar.txVlraPg1.Caption = txVlraPg(0).Text
RelPagosPagar.txVlraPg2.Caption = txVlraPg(1).Text
RelPagosPagar.txVlraPg3.Caption = txVlraPg(2).Text
RelPagosPagar.txVlraPg4.Caption = txVlraPg(3).Text
RelPagosPagar.txVlraPg5.Caption = txVlraPg(4).Text
RelPagosPagar.lbDt.Caption = "Data inicial " & txDtINI.Text & " final " & txDtFIM.Text
RelPagosPagar.Show
Unload Me
End Sub

Private Sub Command1_Click()
Dim SQL           As String
Dim conAPagar     As Recordset
Dim conPagas      As Recordset
Dim conNaoPgnoInt As Recordset

'2.9.0 Ajuste dos valores pagos e a pagar quanto as tarefas dinamicas
Dim a             As Integer
Dim sData         As String

If IsDate(txDtINI.Text) = False Or IsDate(txDtFIM.Text) = False Then
    '2.7.2 Logar todas mensagens
    msgboxL "Data Inválida"
Else

    '2.9.0 Ajuste dos valores pagos e a pagar quanto as tarefas dinamicas
    For a = 0 To 4
        txVlrPg(a).Text = ""
        txVlrNpgNoInt(a).Text = ""
        txVlraPg(a).Text = ""
    Next
    sData = DTSqls(txDtINI.Text) & " and " & DTSqls(txDtFIM.Text, True)

    'Valores Pagos
    SQL = "SELECT Sum(Tarefas.Vlr) AS SomaDeVlr, tpConcertos.tipo "
    SQL = SQL & "from Tarefas, tpConcertos "
    SQL = SQL & "Where Tarefas.Pago BetWeen " & sData
    SQL = SQL & " and tpConcertos.tipo = Tarefas.concerto "
    SQL = SQL & "GROUP BY tpConcertos.tipo "
    AbreTB conPagas, SQL, dbOpenSnapshot
    If conPagas.EOF = False Then
        Do While conPagas.EOF = False
            MostraValor txVlrPg(conPagas!Tipo), conPagas!SomaDeVlr
            conPagas.MoveNext
        Loop
    End If
    
    'Valores Concluídos mas não pagos
    SQL = "SELECT Sum(Tarefas.Vlr) AS SomaDeVlr, tpConcertos.tipo "
    SQL = SQL & "from Tarefas, tpConcertos "
    SQL = SQL & "Where Tarefas.DtConclusao BetWeen " & sData
    SQL = SQL & " and Tarefas.Pago is Null "
    SQL = SQL & " and tpConcertos.tipo = Tarefas.concerto "
    SQL = SQL & "GROUP BY tpConcertos.tipo "
    AbreTB conNaoPgnoInt, SQL, dbOpenSnapshot
    If conNaoPgnoInt.EOF = False Then
        Do While conNaoPgnoInt.EOF = False
            MostraValor txVlrNpgNoInt(conNaoPgnoInt!Tipo), conNaoPgnoInt!SomaDeVlr
            conNaoPgnoInt.MoveNext
        Loop
    End If
    
    'Assumidos mas não concluídos
    SQL = "SELECT Sum(Tarefas.Vlr) AS SomaDeVlr, tpConcertos.tipo "
    SQL = SQL & "from Tarefas, tpConcertos "
    SQL = SQL & "Where Tarefas.Situacao = 2 "
    SQL = SQL & "and tpConcertos.tipo = Tarefas.concerto "
    SQL = SQL & "and tarefas.Orc in "
    SQL = SQL & "( "
    SQL = SQL & "Select Orcamento "
    SQL = SQL & "from Orcamento "
    SQL = SQL & "Where Pagamento BetWeen " & sData
    SQL = SQL & "Order By Orcamento "
    SQL = SQL & ") "
    SQL = SQL & "GROUP BY tpConcertos.tipo"
    AbreTB conAPagar, SQL, dbOpenSnapshot
    If conAPagar.EOF = False Then
        Do While conAPagar.EOF = False
            MostraValor txVlraPg(conAPagar!Tipo), conAPagar!SomaDeVlr
            conAPagar.MoveNext
        Loop
    End If
    
'    SQL = "SELECT Sum(Mecanica) AS SomaDeMecanica, Sum(Pintura) AS SomaDePintura, "
'    SQL = SQL & "Sum(Chapeação) AS SomaDeChapeação, Sum(Eletricidade) AS SomaDeEletricidade, "
'    SQL = SQL & "Sum(Guarnição) As SomaDeGuarnição "
'    SQL = SQL & "from Orcamento "
'    '2.4.8 Data final de intervalo no relatório "valores não pagos no intervalo"
'    SQL = SQL & "Where Pagamento BetWeen " & DTSqls(txDtINI.Text) & " and " & DTSqls(txDtFIM.Text, True)
'    AbreTB conPagas, SQL, dbOpenSnapshot
'
'    '2.4.7 Ajuste nos campos dos totais de valores
'    MostraValor txVlrPg(0), conPagas!SomaDeMecanica
'    MostraValor txVlrPg(1), conPagas!SomaDeChapeação
'    MostraValor txVlrPg(2), conPagas!SomaDeGuarnição
'    MostraValor txVlrPg(3), conPagas!SomaDePintura
'    MostraValor txVlrPg(4), conPagas!SomaDeEletricidade
'
'    '2.4.8 Mostrar valores não pagos no intervalo
'    SQL = "SELECT Sum(Mecanica) AS SomaDeMecanica, Sum(Pintura) AS SomaDePintura, "
'    SQL = SQL & "Sum(Chapeação) AS SomaDeChapeação, Sum(Eletricidade) AS SomaDeEletricidade, "
'    SQL = SQL & "Sum(Guarnição) As SomaDeGuarnição "
'    SQL = SQL & "from Orcamento "
'    SQL = SQL & "Where Data BetWeen " & DTSqls(txDtINI.Text) & " and " & DTSqls(txDtFIM.Text, True)
'    SQL = SQL & " and (Pagamento Is Null Or Pagamento < #1/1/2000#) "
'    AbreTB conNaoPgnoInt, SQL, dbOpenSnapshot
'    MostraValor txVlrNpgNoInt(0), conNaoPgnoInt!SomaDeMecanica
'    MostraValor txVlrNpgNoInt(1), conNaoPgnoInt!SomaDeChapeação
'    MostraValor txVlrNpgNoInt(2), conNaoPgnoInt!SomaDeGuarnição
'    MostraValor txVlrNpgNoInt(3), conNaoPgnoInt!SomaDePintura
'    MostraValor txVlrNpgNoInt(4), conNaoPgnoInt!SomaDeEletricidade
'
'    SQL = "SELECT Sum(Mecanica) AS SomaDeMecanica, Sum(Pintura) AS SomaDePintura, "
'    SQL = SQL & "Sum(Chapeação) AS SomaDeChapeação, Sum(Eletricidade) AS SomaDeEletricidade, "
'    SQL = SQL & "Sum(Guarnição) As SomaDeGuarnição "
'    SQL = SQL & "from Orcamento "
'    SQL = SQL & "Where Pagamento Is Null Or Pagamento < #1/1/2000# "
'    AbreTB conAPagar, SQL, dbOpenSnapshot
'    '2.4.8 Mostrar valores não pagos no intervalo
'    MostraValor txVlraPg(0), conAPagar!SomaDeMecanica - conNaoPgnoInt!SomaDeMecanica
'    MostraValor txVlraPg(1), conAPagar!SomaDeChapeação - conNaoPgnoInt!SomaDeChapeação
'    MostraValor txVlraPg(2), conAPagar!SomaDeGuarnição - conNaoPgnoInt!SomaDeGuarnição
'    MostraValor txVlraPg(3), conAPagar!SomaDePintura - conNaoPgnoInt!SomaDePintura
'    MostraValor txVlraPg(4), conAPagar!SomaDeEletricidade - conNaoPgnoInt!SomaDeEletricidade
                
    '2.4.8 Impressão da consulta de valores pagos e a pagar
    btImpres.Enabled = True
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim DT As Date

InicForm Me
DT = DtINIRel()
txDtINI.Text = Format(DT, "dd/mm/yyyy")

'2.4.8 Data final de intervalo no relatório "valores não pagos no intervalo"
txDtFIM.Text = Format(Now, "dd/mm/yyyy")
End Sub
