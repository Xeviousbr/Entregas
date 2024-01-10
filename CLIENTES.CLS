VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'4.5.6 Não permitir novo orçamento para cliente em atraso em modo balcão
'4.5.4 Cliente em atraso
'3.7.8 Troca de carros entre os clientes
'3.1.5 Previsao para alteracao do nome indevidamente depois de ter localizado o cliente no combo
'3.0.6 Aumento de log para operaçao de colocar os carros no combo do orçamento
'3.0.6 Aprimoramento do tratamento de erro para operação de colocar os carros no combo do orçamento
'3.0.0 Alteração na obtenção dos dados de telefone e endereço do cliente
'2.8.4 Otimizar a Obtenção do Nr do Cliente
'2.7.2 Colocar o telefone no orçamento [A4]
'2.7.2 Logar todas mensagens
'2.4.3 Concerto da situação em que as vezes não aparecia o endereço
'2.0.5 Implantação do Log
'2.0.0 Troca do engine de DAO 3.5 para DAO 3.6

Private NomeCli()   As String * 30
Private QuaCli      As Integer
Private VarCliAtual As String * 30
Private l_NrCli     As Integer

'2.8.4 Otimizar a Obtenção do Nr do Cliente
Private l_Nome As String

'4.5.4 Cliente em atraso
Private l_EmAtraso As Boolean

Public Sub CorPlaca(Carro As Integer, Cor As String, Placa As String, sCarro As String)
Dim Abriu As Boolean
Dim SQL   As String
Dim CP    As Recordset

'SQL = "SELECT DISTINCTROW Carros.Placa, Carros.Cor, Carros.Modelo "
'SQL = SQL + "FROM Clientes INNER JOIN Carros ON Clientes.NrCli = Carros.NrCli "
'If Placa > "" Then
'    SQL = SQL + "WHERE ((Carros.Placa='" + Placa + "') AND (Clientes.Nome='" + Trim(CliAtual) + "'));"
'Else
'    SQL = SQL + "WHERE ((Clientes.Nome='" + CliAtual + "')); "
'End If

SQL = "SELECT DISTINCTROW Placa, Cor, Modelo "
SQL = SQL & "from Carros "
If Placa > "" Then
    SQL = SQL & "WHERE Placa='" & Placa & "' AND NrCli=" & l_NrCli
Else
    SQL = SQL & "WHERE NrCli=" & l_NrCli
End If

'3.0.6 Aprimoramento do tratamento de erro para operação de colocar os carros no combo do orçamento
'AbreTB CP, SQL, dbOpenDynaset
AbreTBCP_Antes:
Loga SQL, lSQL
On Local Error GoTo Err_CorPlaca
Set CP = db.OpenRecordset(SQL, dbOpenDynaset)
Abriu = True

AbreTBCP_Depois:
On Local Error GoTo Err_CorPlaca
If CP.RecordCount > 0 Then
    CP.MoveFirst
    If Placa = "" Then
        For Cont% = 1 To Carro
            CP.MoveNext
        Next
    End If
    Placa = CP(0)
    Cor = CP(1)
    sCarro = CP(2)
End If
CP.Close

'3.0.6 Aumento de log para operaçao de colocar os carros no combo do orçamento
Loga "Saiu do CorPlaca", lDBG
Exit Sub

'3.0.6 Aprimoramento do tratamento de erro para operação de colocar os carros no combo do orçamento
Err_CorPlaca:
Select Case msgboxL("Houve erro interno no programa de tipo:" & vbCrLf & Error(Err), vbInformation + vbYesNo, "Deseja tentar novamente?")
    Case vbYes
        If Abriu Then
            Abriu = False
            CP.Close
            Resume AbreTBCP_Antes
        Else
            Resume AbreTBCP_Depois
        End If
    Case vbNo
        msgboxL "Logo apos tente refazer a operação novamente", vbApplicationModal, "O programa será fechado agora"
        End
End Select
End Sub

Private Sub LeClientes()
'2.0.0 Troca do engine de DAO 3.5 para DAO 3.6
Dim QuantRegCli As Integer
Dim SQL       As String
Dim TbCli       As Recordset

On Error GoTo DeuPau

SQL = "Select DISTINCTROW Nome "
SQL = SQL & "From Clientes "
SQL = SQL & "WHERE ((Nome>" + Chr(34) + Chr$(34) + ")) "

'4.5.6 Não permitir novo orçamento para cliente em atraso em modo balcão
If INI.ModoOperacao <> tpEscritorio Then
    SQL = SQL & "And EmAtraso = 0 "
End If

SQL = SQL & "ORDER BY Nome; "

'2.0.5 Implantação do Log
AbreTB TbCli, SQL, dbOpenSnapshot
'Set TbCli = db.OpenRecordset(SQL)

TbCli.MoveFirst
TbCli.MoveLast
QuaCli = 0
QuantRegCli = TbCli.RecordCount
TbCli.MoveFirst
ReDim NomeCli(QuantRegCli) As String * 30
For Cont% = 1 To QuantRegCli
    If Not IsNull(TbCli("Nome")) Then
        QuaCli = QuaCli + 1
        NomeCli(QuaCli) = TbCli("Nome")
    End If
    TbCli.MoveNext
Next
TbCli.Close
Continua:
Exit Sub

DeuPau:
If Err = 3021 Then
   '2.7.2 Logar todas mensagens
   msgboxL "O arquivo de clientes esta vazio", 0, "Erro"
Else
   '2.7.2 Logar todas mensagens
   msgboxL Error$(Err) + " na abertura do tabela clientes", 0, "Erro"
End If
TbCli.Close
Resume Continua
End Sub

Public Sub CarCliCB(cb As ComboBox, Optional execao As String)
For a% = 1 To QuaCli

    '3.7.8 Troca de carros entre os clientes
    If execao <> Trim(NomeCli(a%)) Then
    
        cb.AddItem NomeCli(a%)
    End If
Next
End Sub

'3.1.5 Previsao para alteracao do nome indevidamente depois de ter localizado o cliente no combo
Public Function SetaClienteAtual(sCli As String) As Boolean
Dim rsAux As Recordset

CliAtual = sCli

'4.5.4 Cliente em atraso
'NrCli = Consulta("Select NrCli from clientes Where Nome = '" & sCli & "'")
AbreTB rsAux, "Select NrCli, EmAtraso from clientes Where Nome = '" & sCli & "'", dbOpenSnapshot
NrCli = rsAux!NrCli
EmAtraso = rsAux!EmAtraso

SetaClienteAtual = (NrCli > 0)
End Function

Public Property Get CliAtual()
CliAtual = VarCliAtual
End Property

'3.1.5 Previsao para alteracao do nome indevidamente depois de ter localizado o cliente no combo
Private Property Let CliAtual(vNewValue)
'Public Property Let CliAtual(vNewValue)
VarCliAtual = vNewValue
End Property

Public Property Get Telefone()
'3.0.0 Alteração na obtenção dos dados de telefone e endereço do cliente
Telefone = SN(Consulta("Select Telefone From Clientes Where NrCli = " & NrCli), vbString)

'2.7.2 Colocar o telefone no orçamento [A4]
'Dim SQL As String
'
'SQL = "SELECT Telefone "
'SQL = SQL & "From Clientes "
'SQL = SQL & "WHERE Nome='" + CliAtual + "' "
'Telefone = Consulta(SQL)
'If Telefone = "" Then
'    SQL = "SELECT Telefone "
'    SQL = SQL & "From Clientes "
'    SQL = SQL & "WHERE Nome Like '" + CliAtual + "*' "
'    Telefone = Consulta(SQL)
'End If
End Property

Public Property Get Endereco()
'3.0.0 Alteração na obtenção dos dados de telefone e endereço do cliente
Endereco = SN(Consulta("Select Ender From Clientes Where NrCli = " & NrCli), vbString)

'Dim SQL As String
'
''2.4.3 Concerto da situação em que as vezes não aparecia o endereço
'SQL = "SELECT Ender "
'SQL = SQL & "From Clientes "
'SQL = SQL & "WHERE Nome='" + CliAtual + "' "
'Endereco = Consulta(SQL)
'If Endereco = "" Then
'    SQL = "SELECT Ender "
'    SQL = SQL & "From Clientes "
'    SQL = SQL & "WHERE Nome Like '" + CliAtual + "*' "
'    Endereco = Consulta(SQL)
'End If
End Property

Public Sub CarregaCarros(cb As ComboBox)
Dim CAR As Recordset
'Dim Selec As String
Dim SQL As String

cb.Clear

'Selec = "SELECT DISTINCTROW Carros.Modelo "
'Selec = Selec + "FROM Clientes INNER JOIN Carros ON Clientes.NrCli = Carros.NrCli "
'Selec = Selec + "WHERE ((Clientes.Nome='" + Trim(CliAtual) + "')); "

'If l_NrCli = 0 Then Stop

SQL = "Select Modelo From Carros Where NrCli = " & NrCli

'2.0.5 Implantação do Log
AbreTB CAR, SQL, dbOpenDynaset
'Set CAR = db.OpenRecordset(Selec, dbOpenDynaset)

On Error GoTo DeuGuru
CAR.MoveFirst
CAR.MoveLast
CAR.MoveFirst
nrCar = CAR.RecordCount
For a% = 1 To nrCar
    cb.AddItem CAR(0)
    CAR.MoveNext
Next
CAR.Close
Sai:
Exit Sub

DeuGuru:
If Err <> 3021 Then
    '2.7.2 Logar todas mensagens
    msgboxL Error$(Err)
End If
Resume Sai
End Sub

Private Sub Class_Initialize()
l_EmAtraso = False

LeClientes

'2.8.4 Otimizar a Obtenção do Nr do Cliente
l_Nome = "-"
End Sub

Public Property Get NrCli() As Integer
NrCli = l_NrCli
End Property

Public Property Let NrCli(ByVal vNewValue As Integer)
l_NrCli = vNewValue
End Property

Public Function NrDoNome(Nome As String) As Long
'2.8.4 Otimizar a Obtenção do Nr do Cliente
If Nome <> l_Nome Then
    l_NrCli = Consulta("Select NrCli From Clientes Where Nome = '" & Nome & "'", True)
    l_Nome = Nome
End If
NrDoNome = l_NrCli
End Function

Public Property Get Nome() As String
'2.8.4 Otimizar a Obtenção do Nr do Cliente
Nome = l_Nome
End Property

Public Property Let Nome(ByVal vNewValue As String)
'2.8.4 Otimizar a Obtenção do Nr do Cliente
l_Nome = vNewValue
End Property

'4.5.4 Cliente em atraso
Public Property Get EmAtraso() As Boolean
EmAtraso = l_EmAtraso
End Property

Public Property Let EmAtraso(ByVal vNewValue As Boolean)
l_EmAtraso = vNewValue
End Property
