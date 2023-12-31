VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'5.0.7 Gravação do Ini em C:\OrCarro
'5.0.4 Correções de bugs
'5.0.1 Ajuste na impressão USB
'4.9.8 Selecionar a impressora ao imprimir
'4.9.7 Tela para gravar se a impressora é tipo fita ou não
'4.9.5 Definir o nr do tamanho da fonte na impressão USB
'4.9.3 Ajuste na impressão USB
'4.4.0 No Windows Seven o modo passa a ser definido pelo INI
'4.3.5 Tipo de impressora passa a ser configurado pelo INI
'4.1.6 Ajuste na gravação da impressora para o Windows 8
'4.1.3 Gravação da configuração da impressora no hd alternativamente
'3.8.N Guardar as opções do relatório de pagamentos
'3.3.0 Critério de quantidade de carros para liberar as comissões
'3.2.3 Gravação automática dos itens
'3.2.1 Pesquisa de tarefas realizadas
'3.1.0 Modo Balcão
'3.0.2 Log em Rede
'2.9.6 Informação da verão do aplicativo no log
'2.8.1 Permitir adicionar tarefas pelo orçamento em dois PCs ao mesmo tempo
'2.7.8 Mostrar a versão no log
'2.7.2-7 Criação automática do campo de CGC
'2.7.4-5 Campo de CGC na configuracao
'2.7.4 Opção para utilizar ou não as comissões
'2.7.3 Configuração para valor de gatilho para comissão
'2.7.2 Deixar de imprimir valores em zero, para impressão A4
'2.5.6 Modo Restrito
'2.0.6 Alteração do funcionamento interno das variáveis de configuração
'2.0.5 Implantação do Log
'2.0.3 Conserto da gravação da informação da localização da impressora

Option Explicit

Private ChaveAberta  As Boolean
Private l_UtComissoe As Integer
Private l_TpImpress  As Integer
Private l_LinhasApos As Integer
Private l_Garantia   As Integer
Private l_Key        As Long
Private l_Cor        As Long
Private l_Orc1       As Long
Private l_Versao     As String * 5
Private l_CaminChave As String
Private l_Empresa    As String
Private l_Endereco   As String
Private l_Fones      As String

'2.8.1 Permitir adicionar tarefas pelo orçamento em dois PCs ao mesmo tempo
Private lcPC         As Integer

'3.0.2 Log em Rede
Private lcLogEmRede As Integer

Private Const KeyValType = 0

'Private lcRestrito     As Boolean
Private lcVlrGatComiss As Double

'3.1.0 Modo Balcão
Public Enum tModoOperacao
  tpEscritorio = 0
  tpBalcao = 1
  tpMecanico = 2
End Enum
Private l_ModoOperacao As tModoOperacao

'Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal LPString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'2.7.4-5 Campo de CGC na configuracao
Private lcCGC As String

'3.3.0 Critério de quantidade de carros para liberar as comissões
Private lcQtdCarr As Integer

'4.1.6 Ajuste na gravação da impressora para o Windows 8
Private SO As String

Private Function VSGetPrivateProfileString(ByVal argFileName As String, _
                                           ByVal argSection As String, _
                                           ByVal argKey As String, _
                                           Optional ByVal Defa As String) As String

Dim varSlave As String
Dim varCount As Byte

varSlave = Space(254)
GetPrivateProfileString argSection, argKey, vbNullString, varSlave, 254, argFileName
For varCount = 1 To 254
   If Asc(Mid(varSlave, varCount, 1)) = 0 Then
      varSlave = Left(varSlave, varCount - 1)
      Exit For
   End If
Next
If Len(varSlave) = 0 Then
    varSlave = Defa
End If
VSGetPrivateProfileString = varSlave
End Function

Public Property Let Impressora(ByVal vData As String)
Dim rc    As Long
Dim Teste As String

'5.0.7 Gravação do Ini em C:\OrCarro
'4.1.3 Gravação da configuração da impressora no hd alternativamente
'rc = RegSetValueEx(Key, "Impressora", 0, REG_SZ, ByVal vData, Len(vData))

'Teste = Impressora
'If Teste <> vData Then

    '5.0.7 Gravação do Ini em C:\OrCarro
    WritePrivateProfileString "Impressao", "Caminho", vData, "C:\OrCarro\Orcarro.ini"
    '4.1.6 Ajuste na gravação da impressora para o Windows 8
    'WritePrivateProfileString "Impressao", "Caminho", vData, "C:\Program Files\OrCarro\Orcarro.ini"
    'WritePrivateProfileString "Impressao", "Caminho", vData, AplicDirNat & "\Orcarro.ini"
    
'End If
End Property

Public Sub FechaChave()
Dim rc As Long

rc = RegCloseKey(l_Key)
ChaveAberta = False
End Sub

Public Property Get Impressora() As String
'5.0.7 Gravação do Ini em C:\OrCarro
'Dim tmpVal     As String
Dim Aux        As String
'Dim KeyValSize As Long
'Dim rc         As Long
'Dim l_Key      As Long

'5.0.7 Gravação do Ini em C:\OrCarro
'tmpVal = String$(1024, 0)
'KeyValSize = 1024
'l_Key = INI.Key
'rc = RegQueryValueEx(l_Key, "Impressora", 0, KeyValType, tmpVal, KeyValSize)
'Impressora = Left(tmpVal, KeyValSize - 1)
'If Asc(Impressora) = 0 Then

    '5.0.7 Gravação do Ini em C:\OrCarro
    Aux = VSGetPrivateProfileString("C:\OrCarro\Orcarro.ini", "Impressao", "Caminho")
    '4.1.6 Ajuste na gravação da impressora para o Windows 8
    'Aux = VSGetPrivateProfileString("C:\Program Files\OrCarro\Orcarro.ini", "Impressao", "Caminho")
    '4.1.3 Gravação da configuração da impressora no hd alternativamente
    'Aux = VSGetPrivateProfileString(AplicDirNat & "\Orcarro.ini", "Impressao", "Caminho")
    
    Impressora = Aux
    If Impressora = "" Then
        Impressora = "LPT1"
    End If
    'Impressora = "LPT1"
    
'End If
End Property

Public Property Get UtTempIni() As Integer
'
End Property

Public Property Let UtTempIni(ByVal vNewValue As Integer)
'
End Property

Public Property Get TempIni() As Variant
'
End Property

Public Property Let TempIni(ByVal vNewValue As Variant)
'
End Property

Public Property Get Log() As Boolean
'2.0.5 Implantação do Log
Dim AbertoLoc  As Boolean
Dim tmpVal     As String
Dim Aux        As String
Dim sLog       As String
Dim KeyValSize As Long
Dim rc         As Long

If ChaveAberta = False Then
    AbertoLoc = True
    AbreChave
End If
tmpVal = String$(1024, 0)
KeyValSize = 1024
rc = RegQueryValueEx(l_Key, "Log", 0, KeyValType, tmpVal, KeyValSize)
sLog = Left(tmpVal, KeyValSize - 1)
If Asc(sLog) = 0 Then
    Log = False
Else
    Log = (sLog = "1")
End If
Logar = Log
If AbertoLoc Then
    FechaChave
End If
End Property

Public Property Let Log(ByVal vNewValue As Boolean)
'2.0.5 Implantação do Log
Dim rc    As Long
Dim Ret   As Long
Dim vData As String

Logar = vNewValue
vData = IIf(vNewValue, "1", "0")
rc = RegSetValueEx(l_Key, "Log", 0, REG_SZ, ByVal vData, Len(vData))
End Property

Public Sub AbreChave()
Dim rc   As Long
Dim hKey As Long
Dim Ret  As Long

rc = RegOpenKeyEx(Key, l_CaminChave, 0, KEY_ALL_ACCESS, hKey)
If rc <> 0 Then
    rc = RegCreateKey(Key, l_CaminChave, Ret)
End If
ChaveAberta = True
End Sub

Public Property Get Key() As Long
'Dim Aux As String

If l_Key = 0 Then

    '4.1.6 Ajuste na gravação da impressora para o Windows 8
    SO = SistemaOperacional
    If SO = "Windows NT" Or SO = "Windows 2K" Or SO = "Windows XP" Then
    'Aux = SistemaOperacional
    'If Aux = "Windows NT" Or Aux = "Windows 2K" Or Aux = "Windows XP" Then
    
       l_CaminChave = "Software\ATC Informatica Ltda"
       l_Key = HKEY_CURRENT_USER
    Else
       l_CaminChave = "Config\ATC"
       l_Key = HKEY_LOCAL_MACHINE
    End If
End If
Key = l_Key
End Property

Public Property Get CaminhoChave() As String
CaminhoChave = l_CaminChave
End Property

'Private Function PegaInfoReg(Info As String) As String
'Dim AbertoLoc  As Boolean
'Dim tmpVal     As String
'Dim Aux        As String
'Dim sLog       As String
'Dim KeyValSize As Long
'Dim rc         As Long
'
'If ChaveAberta = False Then
'    AbertoLoc = True
'    AbreChave
'End If
'tmpVal = String$(1024, 0)
'KeyValSize = 1024
'rc = RegQueryValueEx(l_Key, Info, 0, KeyValType, tmpVal, KeyValSize)
'If KeyValSize = 0 Then
'    PegaInfo = ""
'Else
'    PegaInfo = Left(tmpVal, KeyValSize - 1)
'    If Asc(PegaInfo) = 0 Then
'        PegaInfo = ""
'    End If
'End If
'If AbertoLoc Then
'    FechaChave
'End If
'End Function
'
'Public Sub GravaInfoReg(Campo As String, Valor As String)
'Dim AbertoLoc  As Boolean
'Dim rc         As Long
'Dim sValor     As String * 1024
'
'If ChaveAberta = False Then
'    AbertoLoc = True
'    AbreChave
'End If
'sValor = Valor
'rc = RegSetValueEx(l_Key, Campo, 0, REG_SZ, ByVal sValor, Len(Valor))
'If AbertoLoc Then
'    FechaChave
'End If
'End Sub

Public Property Get Versao() As String
If l_Versao = "0.0.0" Then
    l_Versao = Consulta("Select Versao From Config")
End If
'l_Versao = "2.7.4"

'2.9.6 Informação da verão do aplicativo no log
'Loga "VersãoBD = " & l_Versao
'Loga "Versão = " & l_Versao
'2.7.8 Mostrar a versão no log


Versao = l_Versao
End Property

Public Property Let Versao(ByVal vNewValue As String)
'GravaInfoReg "Versao", vNewValue
l_Versao = vNewValue
ExecSql "Update Config Set Versao = '" & l_Versao & "'"
End Property

Private Sub Class_Initialize()
'2.5.6 Modo Restrito
Dim tmpVal     As String
Dim Aux        As String
Dim KeyValSize As Long
Dim rc         As Long
Dim l_Key      As Long

tmpVal = String$(1024, 0)
KeyValSize = 1024
l_Key = INI.Key

'3.1.0 Modo Balcão
'rc = RegQueryValueEx(l_Key, "ModoOperacao", 0, KeyValType, tmpVal, KeyValSize)
''rc = RegQueryValueEx(l_Key, "Restrito", 0, KeyValType, tmpVal, KeyValSize)
'
'Aux = Left(tmpVal, KeyValSize - 1)
'
''4.4.0 No Windows Seven o modo passa a ser definido pelo INI
'If SO = "Windows Oito" Or SO = "Seven" Then
'If SO = "Windows Oito" Then

    '5.0.7 Gravação do Ini em C:\OrCarro
    Aux = VSGetPrivateProfileString("C:\OrCarro\Orcarro.ini", "Config", "Modo")
    'Aux = VSGetPrivateProfileString(AplicDirNat & "\Orcarro.ini", "Config", "Modo")
'End If

'3.1.0 Modo Balcão
l_ModoOperacao = Val(Aux)

'If Asc(Aux) = 0 Then
'    lcRestrito = False
'Else
'    lcRestrito = (Aux = "1")
'End If

l_Versao = "0.0.0"

'2.7.3 Configuração para valor de gatilho para comissão
lcVlrGatComiss = -1

'2.7.4 Opção para utilizar ou não as comissões
l_UtComissoe = -1

'2.7.4-5 Campo de CGC na configuracao
lcCGC = "-"

l_Orc1 = -1

'3.0.2 Log em Rede
lcLogEmRede = -1

'3.3.0 Critério de quantidade de carros para liberar as comissões
lcQtdCarr = -1

'3.3.4 Nr de linhas apos a impressao passa a ser informação local
l_LinhasApos = -1

'4.3.5 Tipo de impressora passa a ser configurado pelo INI
l_TpImpress = -1

End Sub

Public Property Get Empresa() As String
Empresa = l_Empresa
End Property

Public Property Let Empresa(ByVal vNewValue As String)
l_Empresa = vNewValue
End Property

Public Property Get Endereco() As String
Endereco = l_Endereco
End Property

Public Property Let Endereco(ByVal vNewValue As String)
l_Endereco = vNewValue
End Property

Public Property Get Fones() As String
Fones = l_Fones
End Property

Public Property Let Fones(ByVal vNewValue As String)
l_Fones = vNewValue
End Property

Public Property Get Cor() As Long
If l_Cor = 0 Then
    l_Cor = &H80000005
End If
Cor = l_Cor
End Property

Public Property Let Cor(ByVal vNewValue As Long)
l_Cor = vNewValue
End Property

'4.3.5 Tipo de impressora passa a ser configurado pelo INI
Public Property Get TpImpress() As Integer
Dim Aux$

'5.0.1 Ajuste na impressão USB
'If l_LinhasApos = -1 Then

    '5.0.7 Gravação do Ini em C:\OrCarro
    Aux$ = VSGetPrivateProfileString("C:\OrCarro\Orcarro.ini", "Impressao", "TpImpress")
    'Aux$ = VSGetPrivateProfileString(AplicDirNat & "\Orcarro.ini", "Impressao", "TpImpress")
    If Aux$ = "" Then
        l_TpImpress = 1
    Else
        l_TpImpress = Val(Aux$)
    End If
'End If

'4.9.3 Ajuste na impressão USB
If l_TpImpress = 0 Then
    TamFita = 55
ElseIf l_TpImpress = 2 Then
    TamFita = 31
End If

TpImpress = l_TpImpress
End Property

'4.3.5 Tipo de impressora passa a ser configurado pelo INI
Public Property Let TpImpress(ByVal vNewValue As Integer)
Dim Aux$

l_TpImpress = vNewValue
Aux$ = Trim(STR(l_TpImpress))

'5.0.7 Gravação do Ini em C:\OrCarro
WritePrivateProfileString "Impressao", "TpImpress", Aux$, "C:\OrCarro\Orcarro.ini"
'WritePrivateProfileString "Impressao", "TpImpress", Aux$, AplicDirNat & "\Orcarro.ini"
End Property

Public Property Get LinhasApos() As Integer
Dim Aux$

'3.3.4 Nr de linhas apos a impressao passa a ser informação local
If l_LinhasApos = -1 Then
    If Left(App.Path, 2) = "\\" Then
    
        '5.0.7 Gravação do Ini em C:\OrCarro
        Aux$ = VSGetPrivateProfileString("C:\OrCarro\Orcarro.ini", "Impressao", "LinhasApos")
        'Aux$ = VSGetPrivateProfileString(AplicDirNat & "\Orcarro.ini", "Impressao", "LinhasApos")
        
    End If
    If Aux$ = "" Then
        l_LinhasApos = Consulta("Select LinhasApos From Config")
    Else
        l_LinhasApos = Val(Aux$)
    End If
End If
LinhasApos = l_LinhasApos
End Property

Public Property Let LinhasApos(ByVal vNewValue As Integer)
Dim Aux$

'3.3.4 Nr de linhas apos a impressao passa a ser informação local
If Left(App.Path, 2) = "\\" Then
    Aux$ = Trim(STR(vNewValue))
    
    '5.0.7 Gravação do Ini em C:\OrCarro
    WritePrivateProfileString "Impressao", "LinhasApos", Aux$, "C:\OrCarro\Orcarro.ini"
    'WritePrivateProfileString "Impressao", "LinhasApos", Aux$, AplicDirNat & "\Orcarro.ini"
    
Else
    ExecSql "Update Config Set LinhasApos = " & vNewValue
End If
l_LinhasApos = vNewValue
End Property

Public Property Get Garantia() As Integer
Garantia = l_Garantia
End Property

Public Property Let Garantia(ByVal vNewValue As Integer)
l_Garantia = vNewValue
End Property

Public Property Get CaminhoBkp() As String

'5.0.7 Gravação do Ini em C:\OrCarro
CaminhoBkp = VSGetPrivateProfileString("C:\OrCarro\Orcarro.ini", "Backup", "Caminho")
'CaminhoBkp = VSGetPrivateProfileString(AplicDirNat & "\Orcarro.ini", "Backup", "Caminho")

End Function

Public Property Let CaminhoBkp(ByVal vNewValue As String)

'5.0.7 Gravação do Ini em C:\OrCarro
WritePrivateProfileString "Backup", "Caminho", vNewValue, "C:\OrCarro\Orcarro.ini"
'WritePrivateProfileString "Backup", "Caminho", vNewValue, AplicDirNat & "\Orcarro.ini"

End Property

'3.1.0 Modo Balcão
'Public Property Get Restrito() As Boolean
''2.5.6 Modo Restrito
'Restrito = lcRestrito
'End Property
'
'Public Property Let Restrito(ByVal vNewValue As Boolean)
''2.5.6 Modo Restrito
'Dim rc    As Long
'Dim vData As String
'Dim l_Key As Long
'
'vData = IIf(vNewValue = True, "1", "0")
'lcRestrito = vData
'l_Key = INI.Key
'rc = RegSetValueEx(l_Key, "Restrito", 0, REG_SZ, ByVal vData, Len(vData))
'End Property

Public Property Get ImprEmBrano() As Integer

'5.0.7 Gravação do Ini em C:\OrCarro
If VSGetPrivateProfileString(AplicDirNat & "\Orcarro.ini", "Impressao", "ImprEmBranco") = "SIM" Then
'2.7.2 Deixar de imprimir valores em zero, para impressão A4
'If VSGetPrivateProfileString(AplicDirNat & "\Orcarro.ini", "Impressao", "ImprEmBranco") = "SIM" Then
    ImprEmBrano = 1
Else
    ImprEmBrano = 0
End If
End Property

Public Property Let ImprEmBrano(ByVal vNewValue As Integer)
'2.7.2 Deixar de imprimir valores em zero, para impressão A4
Dim STRNewValue As String

If vNewValue = 1 Then
    STRNewValue = "SIM"
Else
    STRNewValue = "NAO"
End If
WritePrivateProfileString "Impressao", "ImprEmBranco", STRNewValue, AplicDirNat & "\Orcarro.ini"
End Property

Public Property Get VlrGatComiss() As Double
'2.7.3 Configuração para valor de gatilho para comissão
If lcVlrGatComiss = -1 Then
    VlrGatComiss = Consulta("Select VlrGatComiss From Config")
    lcVlrGatComiss = VlrGatComiss
End If
End Property

Public Property Let VlrGatComiss(ByVal vNewValue As Double)
'2.7.3 Configuração para valor de gatilho para comissão
ExecSql "Update Config Set VlrGatComiss = " & VlrSql(STR(vNewValue))
End Property

Public Property Get UtComissoes() As Integer
'2.7.4 Opção para utilizar ou não as comissões
If l_UtComissoe = -1 Then
    If SN(Consulta("Select UtComissoes From Config"), vbBoolean) Then
        l_UtComissoe = 1
    Else
        l_UtComissoe = 0
    End If
End If
UtComissoes = l_UtComissoe
End Property

Public Property Let UtComissoes(ByVal vNewValue As Integer)
'2.7.4 Opção para utilizar ou não as comissões
l_UtComissoe = vNewValue
ExecSql "Update Config Set UtComissoes = " & l_UtComissoe
End Property

Public Property Get CGC() As String
'2.7.4-5 Campo de CGC na configuracao
If lcCGC = "-" Then
    lcCGC = SN(Consulta("Select CGC From Config"), vbString)
    
    '2.7.2-7 Criação automática do campo de CGC
'    If lcCGC = "CAMPO INEXISTENTE" Then
'        ExecSql "Alter Table Config Add Column CGC Text(20)"
'        lcCGC = ""
'    End If
End If
CGC = lcCGC
End Property

Public Property Let CGC(ByVal vNewValue As String)
'2.7.4-5 Campo de CGC na configuracao
lcCGC = vNewValue
ExecSql "Update Config Set CGC = '" & lcCGC & "'"
End Property

Public Property Get Orc1() As Long
'2.7.4-5 Campo de CGC na configuracao
If l_Orc1 = -1 Then
    l_Orc1 = SN(Consulta("Select Orc1 From Config"), vbLong)
End If
Orc1 = l_Orc1
End Property

Public Property Get PC() As Integer
'2.8.1 Permitir adicionar tarefas pelo orçamento em dois PCs ao mesmo tempo
If lcPC = 0 Then
    lcPC = PegaNrPC()
End If
PC = lcPC
End Property

Private Function PegaNrPC() As Integer
'2.8.1 Permitir adicionar tarefas pelo orçamento em dois PCs ao mesmo tempo
Dim Sair   As Boolean
Dim Nr     As Integer
Dim EssePC As String

EssePC = NameOfPC()
Do While Sair = False
    Nr = Consulta("Select IDPC From PCs Where Nome = '" & EssePC & "'")
    If Nr > 0 Then
        Sair = True
    Else
        ExecSql "Insert Into PCs (Nome) Values ('" & EssePC & "')"
    End If
Loop
PegaNrPC = Nr
End Function

Public Property Get LogEmRede() As Boolean
'3.0.2 Log em Rede
If lcLogEmRede = -1 Then
    lcLogEmRede = SN(Consulta("Select LogEmRede From Config"), vbByte)
End If
LogEmRede = lcLogEmRede
End Property

Public Property Let LogEmRede(ByVal vNewValue As Boolean)
'3.0.2 Log em Rede
lcLogEmRede = vNewValue
ExecSql "Update Config Set LogEmRede = " & vNewValue
End Property

Public Property Get ModoOperacao() As tModoOperacao
'3.1.0 Modo Balcão
ModoOperacao = l_ModoOperacao
'0 = Escritório
'1 = Balcão
'2 = Mecânico
End Property

Public Property Let ModoOperacao(ByVal vNewValue As tModoOperacao)
'3.1.0 Modo Balcão
Dim rc    As Long
Dim vData As String
Dim l_Key As Long

'4.4.0 No Windows Seven o modo passa a ser definido pelo INI
'If SO = "Windows Oito" Or SO = "Seven" Then
'If SO = "Windows Oito" Then

    Dim AuxS As String
    AuxS = Trim(STR(Val(vNewValue)))
    WritePrivateProfileString "Config", "Modo", AuxS, AplicDirNat & "\Orcarro.ini"
'Else
'    l_ModoOperacao = vNewValue
'    vData = Trim(STR(l_ModoOperacao))
'    l_Key = INI.Key
'    rc = RegSetValueEx(l_Key, "ModoOperacao", 0, REG_SZ, ByVal vData, Len(vData))
'End If
End Property

'3.2.1 Pesquisa de tarefas realizadas
Public Property Get DtPesqTarefas() As Integer
Dim AuxS As String

AuxS = VSGetPrivateProfileString(AplicDirNat & "\Orcarro.ini", "Tarefas", "Data")
If AuxS = "" Then
    DtPesqTarefas = 1
Else
    DtPesqTarefas = Val(AuxS)
End If
End Property

'3.2.1 Pesquisa de tarefas realizadas
Public Property Let DtPesqTarefas(ByVal vNewValue As Integer)
Dim AuxS As String

AuxS = Trim(STR(vNewValue))
WritePrivateProfileString "Tarefas", "Data", AuxS, AplicDirNat & "\Orcarro.ini"
End Property

'3.2.1 Pesquisa de tarefas realizadas
Public Property Get TarefasConclu() As Integer
Dim AuxS As String

AuxS = VSGetPrivateProfileString(AplicDirNat & "\Orcarro.ini", "Tarefas", "Concluida")
If AuxS = "" Then
    TarefasConclu = 1
Else
    TarefasConclu = Val(AuxS)
End If
End Property

'3.2.1 Pesquisa de tarefas realizadas
Public Property Let TarefasConclu(ByVal vNewValue As Integer)
Dim AuxS As String

AuxS = Trim(STR(vNewValue))
WritePrivateProfileString "Tarefas", "Concluida", AuxS, AplicDirNat & "\Orcarro.ini"
End Property

'3.2.1 Pesquisa de tarefas realizadas
Public Property Get TarefasNaoConclu() As Integer
Dim AuxS As String

AuxS = VSGetPrivateProfileString(AplicDirNat & "\Orcarro.ini", "Tarefas", "NaoConcluida")
If AuxS = "" Then
    TarefasNaoConclu = 1
Else
    TarefasNaoConclu = Val(AuxS)
End If
End Property

'3.2.1 Pesquisa de tarefas realizadas
Public Property Let TarefasNaoConclu(ByVal vNewValue As Integer)
Dim AuxS As String

AuxS = Trim(STR(vNewValue))
WritePrivateProfileString "Tarefas", "NaoConcluida", AuxS, AplicDirNat & "\Orcarro.ini"
End Property

'3.2.3 Gravação automática dos itens
Public Property Get GravaAutom() As Integer
Dim AuxS As String

AuxS = VSGetPrivateProfileString(AplicDirNat & "\Orcarro.ini", "Orcam", "GravaAutom")
GravaAutom = Val(AuxS)
End Property

'3.2.3 Gravação automática dos itens
Public Property Let GravaAutom(ByVal vNewValue As Integer)
Dim AuxS As String

AuxS = Trim(STR(vNewValue))
WritePrivateProfileString "Orcam", "GravaAutom", AuxS, AplicDirNat & "\Orcarro.ini"
End Property

Public Property Get QtdCarrComiss() As Integer
'3.3.0 Critério de quantidade de carros para liberar as comissões
If lcQtdCarr = -1 Then
    lcQtdCarr = SN(Consulta("Select QtdCarrComiss From Config"), vbInteger)
End If
QtdCarrComiss = lcQtdCarr
End Property

Public Property Let QtdCarrComiss(ByVal vNewValue As Integer)
'3.3.0 Critério de quantidade de carros para liberar as comissões
lcQtdCarr = vNewValue
ExecSql "Update Config Set QtdCarrComiss = " & lcQtdCarr
End Property

'3.8.N Guardar as opções do relatório de pagamentos
Public Property Get OpcRelPag() As Integer
Dim AuxS As String

AuxS = VSGetPrivateProfileString(AplicDirNat & "\Orcarro.ini", "RelPag", "OpcRelPag")
OpcRelPag = Val(AuxS)
End Property

'3.8.N Guardar as opções do relatório de pagamentos
Public Property Let OpcRelPag(ByVal vNewValue As Integer)
Dim AuxS As String

AuxS = Trim(STR(vNewValue))
WritePrivateProfileString "RelPag", "OpcRelPag", AuxS, AplicDirNat & "\Orcarro.ini"
End Property

Public Property Get Resolucao() As Integer
Dim AuxS As String

AuxS = VSGetPrivateProfileString(AplicDirNat & "\Orcarro.ini", "Config", "AdapRes")
Resolucao = Val(AuxS)
End Property

Public Property Let Resolucao(ByVal vNewValue As Integer)
Dim AuxS As String

AuxS = Trim(STR(vNewValue))
WritePrivateProfileString "Config", "AdapRes", AuxS, AplicDirNat & "\Orcarro.ini"
End Property

'4.9.5 Definir o nr do tamanho da fonte na impressão USB
Public Property Get NrFonte() As Integer
Dim NrFnteTmp As Integer
Dim AuxS      As String

AuxS = VSGetPrivateProfileString(AplicDirNat & "\Orcarro.ini", "Impressao", "NrFonte")
NrFnteTmp = Val(AuxS)
If NrFnteTmp Then
    NrFonte = NrFnteTmp
Else
    NrFonte = 10
End If
End Property

'4.9.5 Definir o nr do tamanho da fonte na impressão USB
Public Property Let NrFonte(ByVal vNewValue As Integer)
Dim AuxS As String

AuxS = Trim(STR(vNewValue))
WritePrivateProfileString "Impressao", "NrFonte", AuxS, AplicDirNat & "\Orcarro.ini"
End Property

'4.9.7 Tela para gravar se a impressora é tipo fita ou não
Public Property Get ImprFitaNr(Posic As Integer) As Integer
Dim AuxS   As String
Dim nrfita As String

nrfita = "Fita" & Trim(STR(Posic))
AuxS = VSGetPrivateProfileString(AplicDirNat & "\Orcarro.ini", "Impressao", nrfita)
ImprFitaNr = Val(AuxS)
End Property

'4.9.8 Selecionar a impressora ao imprimir
Public Property Get ImprEscolhida() As Integer
Dim AuxS   As String

AuxS = VSGetPrivateProfileString(AplicDirNat & "\Orcarro.ini", "Impressao", "ImprEscolhida")
ImprEscolhida = Val(AuxS)
End Property

'4.9.8 Selecionar a impressora ao imprimir
Public Property Let ImprEscolhida(ByVal vNewValue As Integer)
Dim AuxS   As String

AuxS = Trim(STR(vNewValue))
WritePrivateProfileString "Impressao", "ImprEscolhida", AuxS, AplicDirNat & "\Orcarro.ini"
End Property

Public Property Get OrcGrav() As Integer
'5.0.4 Correções de bugs
Dim AuxS   As String

AuxS = VSGetPrivateProfileString(AplicDirNat & "\Orcarro.ini", "Orcam", "OrcGrav")
OrcGrav = Val(AuxS)
End Property

