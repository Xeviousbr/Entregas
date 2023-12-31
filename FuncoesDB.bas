Attribute VB_Name = "FuncoesDB"
'4.2.3 Mais informação na impressão da observação
'3.4.3 Conserto da data das tarefas quando é editado o orçamento
'3.3.8 Data que assumiu a assumiu as tarefas
'2.7.8 Diferenciação para comandos de alteração no log
'2.2.8 Ajuste na seleção da data final no relatório de totais

Option Explicit

Public Sub AbreTB(ByRef rsTB As Recordset, SQL As String, Optional Tipo As RecordsetTypeEnum = dbOpenForwardOnly, Optional NaoLoga As Boolean = False)
'2.0.5 Implantação do Log
If NaoLoga = False Then
    Loga SQL, lSQL
End If
Set rsTB = db.OpenRecordset(SQL, Tipo)
End Sub

Public Function ExecSql(SQL As String) As Long

'If InStr(SQL, "Insert Into Orcamento") Then Stop

'2.7.8 Diferenciação para comandos de alteração no log
'2.0.5 Implantação do Log
Loga SQL, lUPD
On Local Error GoTo RetornaErro
db.Execute SQL
ExecSql = 0

Sai_ExecSql:
On Local Error GoTo 0
Exit Function

RetornaErro:
ExecSql = Err
Loga Error(Err)
Resume Sai_ExecSql
End Function

Public Sub AdapTamCampo(Txt As String, campo As Field, Tb As Recordset, nmTabela As String)
'2.0.2 Aumentar os campos automaticamente
Dim nmCampo     As String
Dim AuxBookMark As Variant
Dim TamTxt      As Integer

TamTxt = Len(Txt)
If TamTxt > campo.Size Then
    AuxBookMark = Tb.Bookmark
    nmCampo = campo.SourceField
    nmTabela = Tb.Name
    Tb.Close
    
    '2.0.5 Implantação do Log
    ExecSql "Alter Table " & nmTabela & " Alter Column " & nmCampo & " Text(" & (TamTxt + 1) & ")"
    'db.Execute "Alter Table " & nmTabela & " Alter Column " & nmCampo & " Text(" & (TamTxt + 1) & ")"
    
    '2.0.5 Implantação do Log
    AbreTB Tb, nmTabela, dbOpenTable
    'Set Tb = db.OpenRecordset(nmTabela)
    
    Tb.Bookmark = AuxBookMark
End If
End Sub

Public Function CompactaBD(LocBase As String)
Dim NmTemp As String

NmTemp = Left(LocBase, Len(LocBase) - 3) + "tmp"
On Error GoTo bdAberto
DBEngine.CompactDatabase LocBase, NmTemp
Kill LocBase
ContinuabdAberto:
Name NmTemp As LocBase
Exit Function

bdAberto:
Resume ContinuabdAberto
End Function

Public Function Consulta(sSQL As String, Optional NaoLoga As Boolean = False) As Variant
Dim rsAux As Recordset

On Local Error GoTo Nulo

'2.0.5 Implantação do Log
AbreTB rsAux, sSQL, dbOpenSnapshot, NaoLoga
'Set rsAux = db.OpenRecordset(sSQL, dbOpenSnapshot)

If IsNull(rsAux(0)) Then GoTo Nulo
If rsAux.BOF = True And rsAux.EOF = True Then GoTo Nulo
Consulta = rsAux(0)
rsAux.Close
Set rsAux = Nothing
Exit Function

Nulo:
On Local Error GoTo 0
Select Case rsAux(0).Type
    Case 0, 2, 3, 5
      Consulta = 0
    Case 202
      Consulta = ""
'    Case Else
'      Loga "rsAux(0).Type=" & rsAux(0).Type
'      Stop
End Select
Set rsAux = Nothing
End Function

Public Function VeSeTemTb(Tabela As String) As Boolean
Dim a   As Integer
Dim TDs As TableDefs

'Verifica se exste a tabela Etiquetas
Set TDs = db.TableDefs
VeSeTemTb = False
For a = 0 To (TDs.Count - 1)
   Debug.Print TDs(a).Name
   If TDs(a).Name = Tabela Then
      VeSeTemTb = True
      Exit For
   End If
Next
End Function

Public Function VlrSql(Inf As String) As String
Dim STR As String

If Inf = "" Then
    '2.1.5 Prever orçamento sem valor
    VlrSql = "0"
Else
    STR = Valo(dado:=Inf)
    STR = Replace(STR, ",", ".")
    VlrSql = STR
End If
End Function

'3.4.3 Conserto da data das tarefas quando é editado o orçamento
Public Function DTSqld(DT As Date, Optional Fim As Boolean) As String
Dim DTAux  As String
Dim Result As String

DTAux$ = Format(DT, "DD/MM/YYYY HH:MM:SS")
Result$ = "#" & Mid$(DTAux$, 7, 4) + Mid$(DTAux$, 3, 4) + Left$(DTAux$, 2) & Right(DTAux$, 9)
If Fim Then
    Result$ = Result$ & " 23:59:59"
End If
DTSqld$ = Result$ & "#"
End Function

Function DTSqls(DT As String, Optional Fim As Boolean) As String
'2.2.8 Ajuste na seleção da data final no relatório de totais
Dim DTAux As String

'3.3.8 Data que assumiu a assumiu as tarefas
Select Case Len(DT)
    Case 6
        DTSqls = "#" & "20" + Right(DT, 2) + "/" + Mid(DT, 3, 2) + "/" + Left(DT, 2)
    Case 10
        DTAux = Left(DT, 10)
        DTSqls = "#" & Mid$(DTAux, 4, 3) + Left$(DTAux, 3) + Right$(DTAux, 4)
        
    '4.2.3 Mais informação na impressão da observação
    Case 16
        DTAux = Left(DT, 10)
        DTSqls = "#" & Mid$(DTAux, 4, 3) + Left$(DTAux, 3) + Right$(DTAux, 4) & " " & Right(DT, 5)
    
    Case 16, 19
        DTAux = Left(DT, 10)
        DTSqls = "#" & Mid$(DTAux, 4, 3) + Left$(DTAux, 3) + Right$(DTAux, 4) & " " & Right(DT, 8)
End Select

If Fim Then
    DTSqls = DTSqls & " 23:59:59"
End If
DTSqls = DTSqls & "#"
End Function

Public Function FA(Palavra As String) As String
'FA = Faz Aspas
FA = Chr(34) + Replace(Trim(Palavra), Chr(34), "'") + Chr(34)
End Function
