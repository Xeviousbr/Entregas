Attribute VB_Name = "Tarefas"
'3.5.9 Tarefas em modo sem comissão são sempre concluídas
'3.5.1 Não excluir fisicamente Mecânico
'3.4.3 Conserto da data das tarefas quando é editado o orçamento
'3.4.1 Conferência interna na operação de gravação das tarefas pelos mecânicos
'3.2.8 Deixar pagar automaticamente a tarefa auto-completavel
'3.1.6 Melhoria no tratamento da gravação das tarefas quanto as tarefas auto-concluíveis
'3.0.5 Transações na gravação das tarefas
'2.9.6 Melhoria no log das tarefas
'2.9.3 Tratamento de erro na gravação das tarefas
'2.8.6 Data da conclusão da tarefa
'2.8.1 Conserto da atualização de tarefas de um orcamento com alguma tarefa paga
'2.8.1 Permitir adicionar tarefas pelo orçamento em dois PCs ao mesmo tempo
'2.8.0 Melhorar o log quanto as tarefas
'2.7.5 Taréfas Dinâmicas

Option Explicit

Public Sub NovasTarefas()
'2.8.1 Permitir adicionar tarefas pelo orçamento em dois PCs ao mesmo tempo
Dim X As Long

If Consulta("Select Count(*) From TarefasTemp Where IDPC = " & INI.PC) > 0 Then
    ExecSql "Delete * From TarefasTemp Where IDPC = " & INI.PC
    X = DoEvents()
End If
ExecSql "Insert Into TarefasTemp (Situacao, IDPC) Values ('" & SemMec & "', " & INI.PC & ") "
'If Consulta("Select Count(*) From TarefasTemp") > 0 Then
'    ExecSql "Delete * From TarefasTemp "
'    x = DoEvents()
'End If
'ExecSql "Insert Into TarefasTemp (Situacao) Values ('" & SemMec & "') "
X = DoEvents()
End Sub

'2.8.6 Data da conclusão da tarefa
Public Sub CarregaTarefas(NrOrc As Long)
Dim X   As Long
Dim SQL As String

'2.8.1 Permitir adicionar tarefas pelo orçamento em dois PCs ao mesmo tempo
ExecSql "Delete * From TarefasTemp Where IDPC = " & INI.PC

'3.4.3 Conserto da data das tarefas quando é editado o orçamento
SQL = "Insert Into TarefasTemp (concerto,Vlr,Situacao,Nome, Pago, IDPC, DtConclusao, DtAssumiu) "

'3.5.9 Tarefas em modo sem comissão são sempre concluídas
If INI.UtComissoes = "1" Then
    SQL = SQL & "SELECT tpConcertos.concerto, Tarefas.Vlr, tpSituacao.situacao, Mecanicos.Nome, Tarefas.Pago, " & INI.PC & " ,Tarefas.DtConclusao, Tarefas.DtAssumiu "
    SQL = SQL & "from Mecanicos "
    SQL = SQL & "INNER JOIN (tpSituacao "
    SQL = SQL & "INNER JOIN (tpConcertos "
    SQL = SQL & "INNER JOIN Tarefas ON tpConcertos.tipo = Tarefas.concerto) "
    SQL = SQL & "ON tpSituacao.tipo = Tarefas.Situacao) "
Else
    SQL = SQL & "SELECT tpConcertos.concerto, Tarefas.Vlr, 'Concluída', Mecanicos.Nome, Tarefas.Pago, " & INI.PC & " ,Tarefas.DtConclusao, Tarefas.DtAssumiu "
    SQL = SQL & "from Mecanicos "
    SQL = SQL & "INNER JOIN (tpConcertos "
    SQL = SQL & "INNER JOIN Tarefas ON tpConcertos.tipo = Tarefas.concerto) "
End If
SQL = SQL & "ON Mecanicos.codi = Tarefas.Mec "
SQL = SQL & "Where Tarefas.Orc = " & NrOrc

'SQL = SQL & "SELECT tpConcertos.concerto, Tarefas.Vlr, tpSituacao.situacao, Mecanicos.Nome, Tarefas.Pago, " & INI.PC & " ,Tarefas.DtConclusao, Tarefas.DtAssumiu "
'SQL = SQL & "from Mecanicos "
'SQL = SQL & "INNER JOIN (tpSituacao "
'SQL = SQL & "INNER JOIN (tpConcertos "
'SQL = SQL & "INNER JOIN Tarefas ON tpConcertos.tipo = Tarefas.concerto) "
'SQL = SQL & "ON tpSituacao.tipo = Tarefas.Situacao) "
'SQL = SQL & "ON Mecanicos.codi = Tarefas.Mec "
'SQL = SQL & "Where Tarefas.Orc = " & NrOrc
ExecSql SQL
X = DoEvents()

'2.8.0 Melhorar o log quanto as tarefas
If INI.Log Then

    '2.8.1 Permitir adicionar tarefas pelo orçamento em dois PCs ao mesmo tempo
    LogaTarefas "Select * From TarefasTemp Where IDPC = " & INI.PC
    'LogaTarefas "Select * From TarefasTemp"

End If
End Sub

Public Sub LogaTarefas(SQL As String, Optional MostraPlaca As Boolean = False)
'2.8.0 Melhorar o log quanto as tarefas
Dim a         As Integer
Dim Aux       As String
Dim rsTarefas As Recordset

'2.9.6 Melhoria no log das tarefas
Dim Teste As String

'1) Eletricidade R$ 20,00 Em Andamento ""
'1) IMD6592 Mecânica R$123,00 Concluido 7/4/2013 Pago em 07/04/2013

AbreTB rsTarefas, SQL, dbOpenSnapshot

'2.9.6 Melhoria no log das tarefas
On Local Error GoTo SemTarefas
Teste = STR(rsTarefas!concerto)
On Local Error GoTo 0
    
Do While rsTarefas.EOF = False
    a = a + 1
    Aux = Trim(STR(a)) & ") "
    If MostraPlaca Then
        Aux = Aux & rsTarefas!Placa & " "                               'Placa
    End If
    Aux = Aux & rsTarefas!concerto                                      'Concerto
    Aux = Aux & " R$" & Format(rsTarefas!Vlr, "##,###,##0.00") & " "    'Vlr
    Aux = Aux & SN(rsTarefas!Situacao, vbString) & " "                  'Situacao
    If MostraPlaca = False Then
        Aux = Aux & SN(rsTarefas!Nome, vbString)                        'Mecânico
    End If
    If SN(rsTarefas!PAGO, vbDate) > 0 Then                              'Data de Pagamento
        Aux = Aux & " Pago em " & Format(rsTarefas!PAGO, "DD/MM/YYYY")
    End If
    If a = 1 Then
        Loga " ", lDBG
    End If
    Loga Aux, lDBG
    rsTarefas.MoveNext
Loop

'2.9.6 Melhoria no log das tarefas
Sai_LogaTarefas:

Loga " ", lDBG

'2.9.6 Melhoria no log das tarefas
Exit Sub

'2.9.6 Melhoria no log das tarefas
SemTarefas:
Loga "Sem Tarefas", lDBG
Resume Sai_LogaTarefas
End Sub

Public Function GravaTarefas(Orcam As Long) As Integer
Dim JaLogou     As Boolean
Dim nrErro      As Long
Dim lnErro      As Long
Dim cVlr        As Currency
Dim sConst      As String
Dim SQL         As String
Dim rsTarefTemp As Recordset

'2.9.3 Tarefas auto-concluíveis
Dim sSit          As String
Dim rsTpConcertos As Recordset
Dim DtPagto       As Date

'2.9.4 Ajuste nas Tarefas auto-concluíveis
Dim DtConcl As Date

'3.1.6 Melhoria no tratamento da gravação das tarefas quanto as tarefas auto-concluíveis
Dim MecDaTarefa As Integer

'3.4.3 Conserto da data das tarefas quando é editado o orçamento
Dim DtAssumiu As Date

'2.9.3 Tratamento de erro na gravação das tarefas
On Local Error GoTo Err_GravaTarefas

'3.0.5 Transações na gravação das tarefas
Loga "BeginTrans", lDBG
WK.BeginTrans

'db.Transactions
1000 ExecSql "Delete From Tarefas Where Orc = " & Orcam

     '2.8.1 Permitir adicionar tarefas pelo orçamento em dois PCs ao mesmo tempo
1010 SQL = "Select * From TarefasTemp Where IDPC = " & INI.PC
1020 SQL = SQL & " Order By ID"
     'SQL = "Select * From TarefasTemp Order By ID "

1030 AbreTB rsTarefTemp, SQL, dbOpenDynaset
1040 If rsTarefTemp.EOF = False Then
1050     rsTarefTemp.MoveFirst
1060     Do While rsTarefTemp.EOF = False
    
             '2.8.0 Melhorar o log quanto as tarefas
             If JaLogou = False Then
                 JaLogou = True
1070             If INI.Log Then
1080                 LogaTarefas SQL
                 End If
             End If
    
1090         sConst = SN(rsTarefTemp!concerto, vbString)
             If sConst > "" Then
             
                 '2.9.3 Tarefas auto-concluíveis
1100             DtPagto = SN(rsTarefTemp.Fields("Pago").Value, vbDate)

                 '2.9.4 Ajuste nas Tarefas auto-concluíveis
1115             DtConcl = SN(rsTarefTemp.Fields("DtConclusao").Value, vbDate)

                 '3.4.3 Conserto da data das tarefas quando é editado o orçamento
                 DtAssumiu = SN(rsTarefTemp.Fields("DtAssumiu").Value, vbDate)
                 
1110             SQL = "Select tipo, Mec from tpConcertos Where concerto = '" & rsTarefTemp!concerto & "'"
1120             AbreTB rsTpConcertos, SQL, dbOpenForwardOnly
             
                 'Orc
1130             SQL = Orcam & ", "

                 '3.1.6 Melhoria no tratamento da gravação das tarefas quanto as tarefas auto-concluíveis
1135             MecDaTarefa = SN(rsTpConcertos!Mec, vbInteger)
                                  
                 'Mec
                 '3.1.6 Melhoria no tratamento da gravação das tarefas quanto as tarefas auto-concluíveis
1140             If MecDaTarefa > 0 Then
1150                SQL = SQL & MecDaTarefa & ", "
                  '2.9.3 Tarefas auto-concluíveis
'1140             If rsTpConcertos!Mec > 0 Then
'1150                SQL = SQL & rsTpConcertos!Mec & ", "

                 Else
1160                If SN(rsTarefTemp!Nome, vbString) = "" Then
                        SQL = SQL & " 0, "
                    Else
                    
                        '3.5.1 Não excluir fisicamente Mecânico
1170                    SQL = SQL & Consulta("Select codi from Mecanicos Where Nome = '" & rsTarefTemp!Nome & "'") & ", "
                        'SQL = SQL & Consulta("Select codi from Mecanicos Where Nome = '" & rsTarefTemp!Nome & "'") & ", "
                        
                    End If
                 End If
                 
                 'Vlr
1180             cVlr = SN(rsTarefTemp!Vlr, vbCurrency)
                 If cVlr = 0 Then
                     GravaTarefas = 2
1190                 rsTarefTemp.Close

                     '3.0.5 Transações na gravação das tarefas
                     Loga "Rollback", lDBG
                     WK.Rollback
                     
                     Exit Function
                 End If
1200             SQL = SQL & VlrSql(STR(cVlr)) & ", "

                 '2.9.3 Tarefas auto-concluíveis
                 'concerto
1210             SQL = SQL & rsTpConcertos!Tipo & ", "
                 'SQL = SQL & Consulta("Select tipo from tpConcertos Where concerto = '" & rsTarefTemp!concerto & "'") & ", "

                 '2.9.3 Tarefas auto-concluíveis
                 'Situacao
                 
                 '3.1.6 Melhoria no tratamento da gravação das tarefas quanto as tarefas auto-concluíveis
1220             If MecDaTarefa > 0 Then
'1220             If rsTpConcertos!Mec > 0 Then
                    SQL = SQL & "3"
                    
                    '2.9.4 Ajuste nas Tarefas auto-concluíveis
                    If DtConcl = 0 Then
                    
                        '3.2.8 Deixar pagar automaticamente a tarefa auto-completavel
                        'DtPagto = Now
                                                
                        '3.2.8 Deixar pagar automaticamente a tarefa auto-completavel
                        DtConcl = Now
                        
                        '2.9.4 Ajuste nas Tarefas auto-concluíveis
                        'DtConcl = DtPagto
                        
                    End If
                    
                    '5.0.0 Gravar corretamente a data que assumiu caso seja tarefa tipo eletrica
                    DtAssumiu = Now
                    
                 Else
1230                sSit = IIf(SN(rsTarefTemp!Situacao) = "", SemMec, rsTarefTemp!Situacao)
1240                SQL = SQL & Consulta("Select tipo from tpSituacao Where Situacao = '" & sSit & "'")
                 End If
                 
                 '2.9.3 Tarefas auto-concluíveis
1250             If DtPagto > 0 Then
                 'If SN(rsTarefTemp.Fields("Pago").Value, vbDate) > 0 Then

                     '2.9.4 Ajuste nas Tarefas auto-concluíveis
1260                 SQL = "Insert Into Tarefas (Orc, Mec, Vlr, concerto, Situacao, DtConclusao, Pago) Values (" & SQL
'1260                 SQL = "Insert Into Tarefas (Orc, Mec, Vlr, concerto, Situacao, Pago) Values (" & SQL

                     '2.9.4 Ajuste nas Tarefas auto-concluíveis
1265                 SQL = SQL & "," & DTSqls(Format(DtConcl, "DD/MM/YYYY"))
                
                     '2.9.3 Tarefas auto-concluíveis
1270                 SQL = SQL & "," & DTSqls(Format(DtPagto, "DD/MM/YYYY"))
                     '2.8.1 Conserto da atualização de tarefas de um orcamento com alguma tarefa paga
                     'SQL = SQL & "," & DTSqls(Format(rsTarefTemp!PAGO, "DD/MM/YYYY")) & ")"
                
                 '2.9.4 Ajuste nas Tarefas auto-concluíveis
                 ElseIf DtConcl > 0 Then
                 
                     '3.4.3 Conserto da data das tarefas quando é editado o orçamento
1273                 SQL = "Insert Into Tarefas (Orc, Mec, Vlr, concerto, Situacao, DtConclusao, DtAssumiu) Values (" & SQL
                     SQL = SQL & "," & DTSqld$(DtConcl) & "," & DTSqld$(DtAssumiu)

                 Else
                 
                    If DtAssumiu > 0 Then
                        '3.4.3 Conserto da data das tarefas quando é editado o orçamento
                        SQL = "Insert Into Tarefas (Orc, Mec, Vlr, concerto, Situacao, DtAssumiu) Values(" & SQL
                        SQL = SQL & "," & DTSqld$(DtAssumiu)

                    Else
1280                    SQL = "Insert Into Tarefas (Orc, Mec, Vlr, concerto, Situacao) Values(" & SQL
                    End If
                 End If
                 
1290             ExecSql SQL & ")"

                 '2.9.3 Tarefas auto-concluíveis
1300             rsTpConcertos.Close
                 
             End If
1310         rsTarefTemp.MoveNext
         Loop
     End If
     
1320 rsTarefTemp.Close
     GravaTarefas = 0
     
     '3.0.5 Transações na gravação das tarefas
     Loga "CommitTrans", lDBG
     WK.CommitTrans
     
     On Local Error GoTo 0
     Exit Function
     
Err_GravaTarefas:
    '3.0.5 Transações na gravação das tarefas
    nrErro = Err
    lnErro = Erl
    Loga "Rollback", lDBG
    WK.Rollback
    msgboxL "Erro de tipo: " & Error(nrErro) & vbCrLf & "na linha " & lnErro & vbCrLf & vbCrLf & "Tente gravar novamente" & vbCrLf & "ou cancele a gravação", vbCritical, "Erro na gravação de tarefas"
    '2.9.3 Tratamento de erro na gravação das tarefas
    'msgboxL "Erro de tipo: " & Error(Err) & vbCrLf & "na linha " & Erl, vbCritical, "Erro na gravação de tarefas"
    
    On Local Error GoTo 0
    GravaTarefas = 1
End Function

'3.4.1 Conferência interna na operação de gravação das tarefas pelos mecânicos
Public Sub GravaTarefaMecanicos(ID As Long, Sit As Integer, Mec As Integer)
Dim SQLu$, SQLc$, Data$, campo$
Dim MomentoU As Date, MomentoC As Date
Dim vez!
Dim OK As Boolean
Dim X&

MomentoU = Now
Data$ = DTSqls(Format(MomentoU, "DD/MM/YYYY HH:MM:SS"))
SQLu$ = "Update Tarefas "
SQLu$ = SQLu$ & "Set Situacao = " & Sit
If Sit = 2 Then
    SQLu$ = SQLu$ & ", Mec = " & Mec
    campo$ = "DtAssumiu"
Else
    campo$ = "DtConclusao"
End If
SQLu$ = SQLu$ & "," & campo$ & " = " & Data$
SQLu$ = SQLu$ & " Where ID = " & ID
SQLc$ = "Select " & campo$ & " from Tarefas Where ID = " & ID
Do While OK = False
    ExecSql SQLu$
    X& = DoEvents()
    Sleep 100 * (vez! + 1)
    MomentoC = Consulta(SQLc$)
    If MomentoC = MomentoU Then
        Exit Do
    End If
    vez! = vez! + 1
    If vez! = 10 Then
        MsgBox "Contacte o programador", vbCritical, "Erro ao gravar a tarefa"
    End If
Loop
End Sub
