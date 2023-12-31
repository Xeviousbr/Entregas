VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTarefas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'3.7.0 Retirar preparação para tarefas dinâmicas
'3.1.2 Ajuste na conversão das taréfas
'2.7.5 Taréfas Dinâmicas
'2.7.5 Retirada dos campos de taréfas da tabela de orçamentos
Private TbOrcs As Recordset
Private DtUm   As Date

'3.7.0 Retirar preparação para tarefas dinâmicas
'Public Sub PrepTarDin(Objeto As Label)
'Dim a   As Integer
'Dim b   As Integer
'Dim SQL As String
'Dim M(4) As String
'
'M(0) = "Mecânica"
'M(1) = "Chapeação"
'M(2) = "Guarnição"
'M(3) = "Pintura"
'M(4) = "Eletricidade"
'DtUm = DateSerial(2011, 10, 1)
'
'SQL = "SELECT Orcamento, Mecanica, Pintura, [Chapeação], Eletricidade, [Guarnição], Pagamento "
'SQL = SQL & "from Orcamento "
'SQL = SQL & "Where Mecanica > 0 Or Pintura > 0 Or [Chapeação] > 0 Or Eletricidade > 0 Or [Guarnição] > 0 "
'AbreTB TbOrcs, SQL, dbOpenForwardOnly
'Do While TbOrcs.EOF = False
'    a = a + 1
'    For b = 1 To 5
'        If TbOrcs.Fields(b).Value > 0 Then
'            InsereTarefa TbOrcs!Orcamento, b, TbOrcs.Fields(b).Value, SN(TbOrcs!Pagamento, vbDate)
'            Objeto.Caption = Trim(STR(a)) & " - " & M(b - 1)
'            Objeto.Refresh
'        End If
'    Next
'    TbOrcs.MoveNext
'Loop
'TbOrcs.Close
'INI.Versao = Trim(" " & App.Major & "." & App.Minor & "." & App.Revision)
'
''2.7.5 Retirada dos campos de taréfas da tabela de orçamentos
'ExecSql "ALTER TABLE Orcamento DROP COLUMN Mecanica"
'ExecSql "ALTER TABLE Orcamento DROP COLUMN Pintura"
'ExecSql "ALTER TABLE Orcamento DROP COLUMN [Chapeação]"
'ExecSql "ALTER TABLE Orcamento DROP COLUMN Eletricidade"
'ExecSql "ALTER TABLE Orcamento DROP COLUMN [Guarnição]"
'
''3.0.7 Ajuste na adaptação do programa para versão 3.0.7 pela base da versão 2.7.2.7
'ExecSql "Update Tarefas Set DtConclusao = Pago Where pago is not null "
'ExecSql "Update Tarefas Set DtConclusao = Now Where pago is null and situacao = 3 "
'End Sub

Private Sub InsereTarefa(NrOrc As Long, campo As Integer, Vlr As Currency, Pagto As Date)
Dim SQL As String

If Pagto > 0 Then
    SQL = "Insert Into Tarefas (Orc, Mec, Vlr, concerto, Situacao, Pago) Values( "
Else
    SQL = "Insert Into Tarefas (Orc, Mec, Vlr, concerto, Situacao) Values( "
End If
SQL = SQL & NrOrc & ","                                     'Orc
SQL = SQL & "0, "                                           'Mec
SQL = SQL & VlrSql(STR(Vlr)) & ","                          'Vlr,

'3.1.2 Ajuste na conversão das taréfas
SQL = SQL & (campo - 1) & ","                                   'concerto
'SQL = SQL & campo & ","                                     'concerto

If Pagto > 0 Then
    SQL = SQL & 3 & ","                                     'Situacao,
    SQL = SQL & DTSqls(Format(Pagto, "DD/MM/YYYY")) & ")"    'Pago,
Else
    SQL = SQL & 2 & ")"                                     'Situacao,
End If
ExecSql SQL
End Sub
