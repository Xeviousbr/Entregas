Attribute VB_Name = "Funcoes"
'2.7.2 Logar todas mensagens
'2.6.4 Impressão do recibo do mecânico
'2.6.1 Mecânicos não devem ver valores
'2.6.0 Troca do título do modo restrito para 'modo Mecânico'
'2.5.6 Modo Restrito
'2.3.3 Possibilitar retornar o carro ao cliente
'2.1.5 Prever orçamento sem valor
'2.0.5 Implantação do Log
'2.0.2 Aumentar os campos automaticamente
'2.0.0 Alteração da indicação do diretório dos dados

Option Explicit

Public AppVersao      As String * 5
Public Base           As String
Public LinhaDeComando As String
Public CaminhoBkp     As String
Public db             As Database
Public OrcasRecordset As Recordset
Public INI            As New clsReg
Public clsCLi         As New clsClientes

Public CorSelec As Long

'2.3.3 Possibilitar retornar o carro ao cliente
Public GCliente As String
Public GGPlaca  As String

Type ItOrc
    Peca As String
    Quant As Single
    Valor As Currency
End Type

'--- Declarações para descobrir a versão do Windows
Private strWinVersion  As String
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion      As Long
  dwMinorVersion      As Long
  dwBuildNumber       As Long
  dwPlatformId        As Long
  szCSDVersion        As String * 128
End Type
Public Enum Enum_OperatingPlatform
  Platform_Windows_32 = 0
  Platform_Windows_95_98_ME = 1
  Platform_Windows_NT_2K_XP = 2
End Enum
Public Enum TFiltraNum
   TFSoNumero = 0
   TFValores = 1
   TFComSinais = 2
   TFAsc = 3
End Enum
Declare Function GetVersionExA Lib "KERNEL32" (lpVersionInformation As OSVERSIONINFO) As Long

'Declarações para lidar com registro
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const REG_SZ = 1
Public Const REG_DWord = 2
Public Const REG_BINARY = 3
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CURRENT_USER = &H80000001
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
Public Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
'--

'Declarações para as funções de Combo
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_FINDSTRING = &H14C
Private Const CB_FINDSTRINGEXACT = &H158
Private Const LB_FINDSTRING = &H18F
Private Const LB_FINDSTRINGEXACT = &H1A2
'--
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

'Função originada do Estoq em VBDOS
'(06/08/95) Fun‡Æo desmembrada da fun‡Æo VeValor
'(08/08/95) Ler corretamente os valores dos textos
'Retorna o valor de uma string escrita com ponto ou virgula
Function Valo(Optional Ctrl As Control, Optional dado As String = "") As Double
Attribute Valo.VB_Description = "Retorna o valor de uma string escrita com ponto ou virgula"
Dim Aux2 As Integer
Dim Aux3 As String * 1
Dim Aux As String

If dado = "" Then
    If TypeOf Ctrl Is Label Then
        Aux = Ctrl.Caption
    Else
        Aux = Ctrl.Text
    End If
Else
    Aux = dado
End If

'(31/01/96) Consertado erro na leitura de um n£mero com percentual
If InStr(Aux, "%") > 0 Then
    Aux = Left$(Aux, InStr(Aux, "%") - 1)
End If

TiraLetra:
Aux3 = Left$(Aux, 1)
If Val(Aux3) = 0 And Aux3 <> "0" And Aux3 <> "." Then
    If Len(Aux) < 2 Then
        Aux = "0"
        'Vlr = 0
        Exit Function
    End If
    Aux = Mid$(Aux, 2)
    GoTo TiraLetra
End If
If InStr(Aux, ",") > 0 Then
MILHAR:
    Aux2 = InStr(Aux, ".")
    If Aux2 > 0 Then
        Aux = Left$(Aux, Aux2 - 1) + Right$(Aux, Len(Aux) - Aux2)
        GoTo MILHAR
    End If
    Mid$(Aux, InStr(Aux, ","), 1) = "."
End If
Valo = Val(Aux)
End Function

Public Sub SetaDirBackup()
If Len(CaminhoBkp) < 3 Then CaminhoBkp = ""
If CaminhoBkp > "" Then
    FrmMenu.BkpDisk.Caption = "Backup para " & CaminhoBkp
    FrmMenu.BkpDisk.Enabled = True
    FrmMenu.MnuRestDisk.Caption = "Restauração apartir de " & CaminhoBkp
    FrmMenu.MnuRestDisk.Enabled = True
End If
End Sub

Public Sub Inicializa()
'Dim tpImpr As Integer
'Dim linhas As Integer
'Dim DefCor As Long
'Dim SQL    As String
'Dim NmCofg As String
'Dim EnCofg As String
'Dim TlCofg As String

Base = App.Path & "\OrCarro.mdb"

CaminhoBkp = INI.CaminhoBkp
SetaDirBackup

Abre:
On Local Error GoTo SemOrCarroMDB
Set db = OpenDatabase(Base)
On Error GoTo 0
Exit Sub

SemOrCarroMDB:
'2.7.2 Logar todas mensagens
msgboxL "Não localizei o arquivo OrCarro.MDB" & Chr(13) & Error(Err), vbCritical, "Impossível iniciar"
End
End Sub

Public Function CarregaConfig()
Dim TbConfig As Recordset

'2.0.6 Alteração do funcionamento interno das variáveis de configuração
AbreTB TbConfig, "Config", dbOpenTable
With TbConfig
    INI.Empresa = !Empresa
    INI.Endereco = !Endereco
    INI.Fones = !Fones
    INI.Cor = !Cor
    INI.TpImpress = !TpImpress
    INI.LinhasApos = !LinhasApos
    INI.Garantia = !Garantia
    CarregaConfig = SN(!Imagem, vbString)
    .Close
End With
CorSelec = 16777215 - INI.Cor
End Function

Public Sub Seleciona()
Attribute Seleciona.VB_Description = "Seleciona um texto."
On Local Error Resume Next
If Len(Screen.ActiveControl.Text) > 0 Then
   Screen.ActiveControl.SelStart = 0
   Screen.ActiveControl.SelLength = Len(RTrim$(Screen.ActiveControl.Text))
End If
On Local Error GoTo 0
End Sub

Public Sub TrataCombo(STR As String, cb As ComboBox, KeyAscii As Integer)
Dim Proc$, Tam%
Dim Cont%, Item%

If KeyAscii <> 16 Then
   If KeyAscii = 40 Or KeyAscii = 38 Then
      STR = ""
   ElseIf KeyAscii = 13 Then
      SendKeys "{Tab}"
   Else
      STR = STR + Chr$(KeyAscii)
      Proc$ = UCase$(STR)
      Tam% = Len(STR)
Procurar:
      For Cont% = 0 To cb.ListCount - 1
         If UCase$(Left$(cb.List(Cont%), Tam%)) = Proc$ Then
            Item% = Cont%
            Exit For
         End If
      Next
      If Item% = 0 Then
         If Tam% = 0 Then
            Exit Sub
         Else
            Tam% = Tam% - 1
            Proc$ = Left$(STR, Tam%)
            GoTo Procurar
         End If
      Else
         cb.ListIndex = Item%
         cb.Text = cb.List(Item%)
      End If
      KeyAscii = 0
   End If
End If
End Sub

'Função que coloca valores nos controles
'(27/07/96) Originado do RECIBO, Cobran
Public Sub MostraValor(Ctrl As Control, Optional dado, Optional Moeda)
'Static Virgula As Boolean
Dim EhNulo As Boolean
Dim Aux    As Double
Dim Aux2   As String
Dim Masc   As String
Dim Nada   As Variant

If IsMissing(dado) Then
    EhNulo = True
    If TypeOf Ctrl Is TextBox Then
        Aux = Valo(Ctrl)
    Else
        Aux = Valo(Ctrl)
    End If
Else
    If IsNull(dado) Then
        EhNulo = True
    Else
        Aux = dado
    End If
End If
If IsMissing(Moeda) Then Moeda = True
If Aux = 0 And EhNulo = True Then
    Aux2 = ""
ElseIf Moeda Then

SetaVirgula:
On Local Error GoTo TrocaVirgula
    Aux2 = Format(Aux, "##,###,###,##0.00") + " "
    On Error GoTo 0
Else
    Aux2 = STR$(dado)
End If
If TypeOf Ctrl Is Label Then
    Ctrl.Caption = Aux2
Else
    Ctrl.Text = Aux2
End If
Exit Sub

TrocaVirgula:
'Virgula = Not Virgula
Resume SetaVirgula
Resume
End Sub

Public Function Extenso(nValor As Single) As String
  
'Faz a validação do argumento
If IsNull(nValor) Or nValor <= 0 Or nValor > 9999999.99 Then
  Exit Function
End If

'Declara as variáveis da função
Dim nContador, nTamanho As Integer
Dim cValor, cParte, cFinal As String
ReDim aGrupo(4), aTexto(4) As String

'Define matrizes com extensos parciais
ReDim aUnid(19) As String
aUnid(1) = "Um ": aUnid(2) = "Dois ": aUnid(3) = "Tres "
aUnid(4) = "Quatro ": aUnid(5) = "Cinco ": aUnid(6) = "Seis "
aUnid(7) = "Sete ": aUnid(8) = "Oito ": aUnid(9) = "Nove "
aUnid(10) = "Dez ": aUnid(11) = "Onze ": aUnid(12) = "Doze "
aUnid(13) = "Treze ": aUnid(14) = "Quatorze ": aUnid(15) = "Quinze "
aUnid(16) = "Dezesseis ": aUnid(17) = "Dezessete ": aUnid(18) = "Dezoito "
aUnid(19) = "Dezenove "

ReDim aDezena(9) As String
aDezena(1) = "Dez ": aDezena(2) = "Vinte ": aDezena(3) = "Trinta "
aDezena(4) = "Quarenta ": aDezena(5) = "Cinquenta "
aDezena(6) = "Sessenta ": aDezena(7) = "Setenta ": aDezena(8) = "Oitenta "
aDezena(9) = "Noventa "

ReDim aCentena(9) As String
aCentena(1) = "Cento ":  aCentena(2) = "Duzentos "
aCentena(3) = "Trezentos ": aCentena(4) = "Quatrocentos "
aCentena(5) = "Quinhentos ": aCentena(6) = "Seiscentos "
aCentena(7) = "Setecentos ": aCentena(8) = "Oitocentos "
aCentena(9) = "Novecentos "

'Divide o valor em vários grupos
cValor = Format$(nValor, "0000000000.00")
aGrupo(1) = Mid$(cValor, 2, 3)
aGrupo(2) = Mid$(cValor, 5, 3)
aGrupo(3) = Mid$(cValor, 8, 3)
aGrupo(4) = "0" + Mid$(cValor, 12, 2)

'Processa cada grupo
For nContador = 1 To 4
  cParte = aGrupo(nContador)
  nTamanho = Switch(Val(cParte) < 10, 1, Val(cParte) < 100, 2, Val(cParte) < 1000, 3)
  If nTamanho = 3 Then
    If Right$(cParte, 2) <> "00" Then
      aTexto(nContador) = aTexto(nContador) + aCentena(Left(cParte, 1)) + "e "
      nTamanho = 2
    Else
      aTexto(nContador) = aTexto(nContador) + IIf(Left$(cParte, 1) = "1", "Cem ", aCentena(Left(cParte, 1)))
    End If
  End If
  If nTamanho = 2 Then
    If Val(Right(cParte, 2)) < 20 Then
      aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 2))
    Else
      aTexto(nContador) = aTexto(nContador) + aDezena(Mid(cParte, 2, 1))
      If Right$(cParte, 1) <> "0" Then
        aTexto(nContador) = aTexto(nContador) + "e "
        nTamanho = 1
      End If
    End If
  End If
  If nTamanho = 1 Then
    aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 1))
  End If
Next

'Gera o formato final do texto
If Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 0 And Val(aGrupo(4)) <> 0 Then
  cFinal = aTexto(4) + IIf(Val(aGrupo(4)) = 1, "Centavo", "Centavos")
Else
  cFinal = ""
  cFinal = cFinal + IIf(Val(aGrupo(1)) <> 0, aTexto(1) + IIf(Val(aGrupo(1)) > 1, "Milhões ", "Milhão "), "")
  If Val(aGrupo(2) + aGrupo(3)) = 0 Then
    cFinal = cFinal + "de "
  Else
    cFinal = cFinal + IIf(Val(aGrupo(2)) <> 0, aTexto(2) + "Mil ", "")
  End If
  cFinal = cFinal + aTexto(3) + IIf(Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 1, "Real ", "Reais ")
  cFinal = cFinal + IIf(Val(aGrupo(4)) <> 0, "e " + aTexto(4) + IIf(Val(aGrupo(4)) = 1, "Centavo", "Centavos"), "")
End If
Extenso = cFinal
End Function


'Public Function CompactaBD(LocBase As String)
'Dim NmTemp As String
'
'NmTemp = Left(LocBase, Len(LocBase) - 3) + "tmp"
'On Error GoTo bdAberto
'DBEngine.CompactDatabase LocBase, NmTemp
'Kill LocBase
'ContinuabdAberto:
'Name NmTemp As LocBase
'Exit Function
'
'bdAberto:
'Resume ContinuabdAberto
'End Function

Public Function TamanhoArq(Arq As String) As Long
Dim Buf As Integer

Buf = FreeFile
Open Arq For Binary As #Buf
TamanhoArq = LOF(Buf)
Close #Buf
If TamanhoArq = 0 Then
   Kill Arq
End If
End Function

Public Function FileExists(FileName As String) As Boolean
'Estoq 5.9.9 Conserto do funcionamento do INI localmente
On Error Resume Next
FileExists = (Dir$(FileName) <> "")
End Function

Function ComplStr(dado As String, Tam As Integer, Prench As String, Tipo As Integer) As String
'Dado é a string
'Tam é o tamanho final da string
'Preenc é o caractere de preenchimento, para preencher com espaçoes use " "
'Tipo é o alinhamento 0=Esquerda, 1=Central e 2=Direita
Dim Aux As Integer
Dim Aux2 As Integer
Dim Aux3 As String
Aux = Len(dado)
If Aux >= Tam Then
    ComplStr = Left$(dado, Tam)
Else
    Select Case Tipo
        Case Is = 0
            'A esquerda
            If Len(dado) < Tam Then
                dado = dado + String$(Tam - Len(dado), Prench)
            End If
        Case Is = 1
            'Central
            Aux2 = (Tam - Aux) / 2
            Aux3 = String$(Aux2, Prench) + dado + String$(Aux2, Prench)
            Aux = Len(Aux3)
            If Aux < Tam Then
                Aux3 = Aux3 + String$(Tam - Aux, Prench)
            End If
            If Aux > Tam Then
                Aux3 = Left$(Aux3, Tam)
            End If
            dado = Aux3
        Case Is = 2
            'A Direita
            If Len(Aux) < Tam Then
                dado = String$(Tam - Len(dado), Prench) + dado
            End If
        End Select
    ComplStr = dado
End If
End Function

'Public Function Consulta(sSQL As String) As Variant
'Dim rsAux As Recordset
'
'On Local Error GoTo Nulo
'
''2.0.5 Implantação do Log
'AbreTB rsAux, sSQL, dbOpenSnapshot
''Set rsAux = db.OpenRecordset(sSQL, dbOpenSnapshot)
'
'If IsNull(rsAux(0)) Then GoTo Nulo
'If rsAux.BOF = True And rsAux.EOF = True Then GoTo Nulo
'Consulta = rsAux(0)
'rsAux.Close
'Set rsAux = Nothing
'Exit Function
'
'Nulo:
'On Local Error GoTo 0
'Select Case rsAux(0).Type
'    Case 0, 3, 5
'      Consulta = 0
'    Case 202
'      Consulta = ""
''    Case Else
''      Loga "rsAux(0).Type=" & rsAux(0).Type
''      Stop
'End Select
'Set rsAux = Nothing
'End Function

Public Function SistemaOperacional() As String
Dim lpVersionInformation As OSVERSIONINFO

  lpVersionInformation.dwOSVersionInfoSize = Len(lpVersionInformation)
  Call GetVersionExA(lpVersionInformation)

  If (lpVersionInformation.dwPlatformId = Platform_Windows_32) Then
    SistemaOperacional = "Windows 32"
  ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_95_98_ME) And (lpVersionInformation.dwMinorVersion = 0) Then
    SistemaOperacional = "Windows 95"
  ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_95_98_ME) And (lpVersionInformation.dwMinorVersion = 10) Then
    SistemaOperacional = "Windows 98"
  ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_95_98_ME) And (lpVersionInformation.dwMinorVersion = 90) Then
    SistemaOperacional = "Windows ME"
  ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion < 5) Then
    SistemaOperacional = "Windows NT"
  ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 5) And (lpVersionInformation.dwMinorVersion = 0) Then
    SistemaOperacional = "Windows 2K"
  ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 5) And (lpVersionInformation.dwMinorVersion = 1) Then
    SistemaOperacional = "Windows XP"
  End If
End Function

'Public Function VeSeTemTb(Tabela As String) As Boolean
'Dim TDs As TableDefs
'
''Verifica se exste a tabela Etiquetas
'Set TDs = db.TableDefs
'VeSeTemTb = False
'For a = 0 To (TDs.Count - 1)
'   Debug.Print TDs(a).Name
'   If TDs(a).Name = Tabela Then
'      VeSeTemTb = True
'      Exit For
'   End If
'Next
'End Function

Public Function SN(Aux As Variant, Optional Tipo As VbVarType) As Variant
If VarType(Aux) = vbError Then
   If Tipo = vbString Then
      Aux = ""
   Else
      Aux = Null
   End If
End If
If IsNull(Aux) = True Then
   If IsNull(Tipo) Then
      SN = ""
   Else
      If Tipo = vbInteger Then
         SN = 0
      ElseIf Tipo = vbString Then
         SN = ""
      End If
   End If
Else
    On Error GoTo Erro_SN
   If Tipo = vbString Then
      If Aux = "" Then
         Aux = " "
      End If
   End If
   On Error GoTo Erro_SN
   SN = Aux
   On Error GoTo 0
End If
Sai_SN:
Exit Function

Erro_SN:
If Tipo = vbCurrency Then
    SN = 0
ElseIf Tipo = vbString Then
    SN = ""
End If
Resume Sai_SN
End Function

Public Sub Colorizar(F As Form)
Dim C As Control

On Local Error Resume Next
For Each C In F
   C.BackColor = INI.Cor
   'If TypeOf C Is TextBox And opt Then
   '   C.ForeColor = CorSelec
   'End If
Next
F.BackColor = INI.Cor
On Local Error GoTo 0
End Sub

Public Function PegaNmCliente() As String
Dim rc         As Long
Dim hKey       As Long
Dim KeyValSize As Long
Dim sCliente   As String
Dim tmpVal     As String
Dim Informacao As String

Const KeyValType = 0
KeyValSize = 1024
tmpVal = String$(KeyValSize, 0)

'Key = RetKey()
rc = RegOpenKeyEx(INI.Key, INI.CaminhoChave, 0, KEY_ALL_ACCESS, hKey)
rc = RegQueryValueEx(hKey, "Cliente", 0, KeyValType, tmpVal, KeyValSize)
Informacao = Left(tmpVal, KeyValSize - 1)
If Asc(Informacao) = 0 Then
    PegaNmCliente = ""
Else
    PegaNmCliente = Informacao
End If
End Function

Public Sub AchaOrc(Selec As String, Optional Pesq$)
Dim CadeOrc As Recordset

'2.0.5 Implantação do Log
AbreTB CadeOrc, Selec, dbOpenDynaset
'Set CadeOrc = db.OpenRecordset(Selec, dbOpenDynaset)

If CadeOrc.EOF Then
    '2.7.2 Logar todas mensagens
    msgboxL "Não localizei o Orçamento"
Else
    CadeOrc.MoveLast
    If CadeOrc.RecordCount = 1 Then
        Load frmOrc
        frmOrc.Leitura = CadeOrc("Orcamento")
        frmOrc.Show
    Else
      'Mostrar um browse para os orçamentos
      Load Orcas
      If Pesq$ = "" Then
         Pesq$ = Chr$(34) + Trim$(CadeOrc("Cliente")) + Chr$(34)
      End If
      Orcas.NmCliente = Pesq$
      Orcas.Show
    End If
End If
CadeOrc.Close
End Sub

Public Sub VeValor(Aux As String, Vlr As Double, ctr As Control, Escr As Integer)
'2.0.3 Cálculo algébrico para a entrada de valores dos itens do orçamento
'Programador Arnaldo, função criada em 1984
'Aux se a origem da informação esta numa string
'Vlr retorna o valor
'crt controle se for para pegar direto
'Escr=1 para escrever o resultado no controle
Dim Aux2    As Integer
Dim a       As Integer
Dim Algebra As Integer
Dim PoeCtr  As Integer
Dim Lugar   As Integer
Dim Aux3    As String
Dim Ped1    As String
Dim Ped2    As String
Dim Oper    As String
Dim Vlr1    As Double
Dim Vlr2    As Double

If Vlr Then
   On Error Resume Next
   PoeCtr = Vlr
   On Error GoTo 0
End If
If Aux = "" Then
    Aux = ctr.Text
End If
If InStr(Aux, "%") > 0 Then
   Aux = Left$(Aux, InStr(Aux, "%") - 1)
End If
If Len(Aux) > 0 Then
   If EhNumero(Right(Aux, 1)) = False Then
      Aux = Left(Aux, Len(Aux) - 1)
   End If
End If
Vlr1 = Val(Aux)
Conta:
If InStr(Aux, "+") > 0 Or InStr(Aux, "-") > 0 Or InStr(Aux, "*") > 0 Or InStr(Aux, "/") > 0 Then
   Algebra = 1
   If Left$(Aux, 1) = "-" Then Aux = "0" + Aux
      For a = 1 To Len(Aux)
         Aux3 = Mid$(Aux, a, 1)
         If Aux3 = "+" Or Aux3 = "-" Or Aux3 = "*" Or Aux3 = "/" Then
            If Oper = "" Then
               Ped1 = Left$(Aux, a - 1)
               Oper = Aux3
               Lugar = a
            Else
               Ped2 = Mid$(Aux, Lugar + 1, a - Lugar - 1)
               Exit For
            End If
         End If
      Next
      If Ped2 = "" Then
         Ped2 = Mid$(Aux, Lugar + 1)
      End If
      Vlr1 = 0
      VeValor Ped1, Vlr1, ctr, 0
      VeValor Ped2, Vlr2, ctr, 0
      Select Case Oper
         Case Is = "+"
            Vlr1 = Vlr1 + Vlr2
         Case Is = "-"
            Vlr1 = Vlr1 - Vlr2
         Case Is = "*"
            Vlr1 = Vlr1 * Vlr2
         Case Is = "/"
            Vlr1 = Vlr1 / Vlr2
      End Select
      If a < Len(Aux) Then
         Aux = STR$(Vlr1) + Mid$(Aux, a)
         Oper = ""
         Ped2 = ""
         GoTo Conta
      End If
End If
If Algebra Then
   Aux = STR$(Vlr1)
   Vlr = Vlr1
Else
   Vlr = Valo(ctr, Aux)
End If
If Escr Then
    MostraValor ctr, Vlr
End If
End Sub

Private Function EhNumero(letra As Variant) As Boolean
'2.0.3 Cálculo algébrico para a entrada de valores dos itens do orçamento
If Val(letra) = 0 And letra <> "0" Then
   EhNumero = False
Else
   EhNumero = True
End If
End Function

Public Sub Atraso(tt As Single)
'Função responsável por um atraso
Dim UmSeg As Double
UmSeg = Now + (tt / 24 / 60 / 60)
Do While Now < (UmSeg)
Loop
End Sub

Public Function VlrCrit(Valor As Variant, VlrMin As Variant, VlrMax As Variant, VlrDefault As Variant) As Variant
If Valor < VlrMin Or Valor > VlrMax Then
    VlrCrit = VlrDefault
Else
    VlrCrit = Valor
End If
End Function

'Public Sub PoemCombo(ct As Control, dado$())
'Dim a As Integer
'
'Do While dado$(a) > " "
'    If Val(Left$(dado$(a), 4)) > 0 Then
'       TextCombo Mid$(dado$(a), 5), ct, 1
'    Else
'        TextCombo dado$(a), ct, 1
'    End If
'    a = a + 1
'Loop
'End Sub

Public Function TextCombo(Texto As String, Combo As Control, Colocar As Integer, Optional ItemData As Integer) As Integer
'Objetivo: Colocar um texto no combo e se não tiver coloca conforme o caso
Dim TamTexto As Integer
Dim Aux As String

Aux = RTrim$(Texto)
TamTexto = Len(Aux)

TextCombo = AcharTextoNaLista(Combo, Texto)
If TextCombo >= 0 And Combo.ListCount > 0 Then
    If TextCombo = 0 And LCase(Aux) = "a" Then
        If LCase(Mid(Combo.List(0), 1, 1)) = "a" Then
            TextCombo = 1
        End If
    End If
    Combo.ListIndex = TextCombo
Else
    If Colocar > 0 Then
        If Trim(Texto) > "" Then
            Combo.AddItem Texto
            If ItemData > 0 Then
               Combo.ItemData(Combo.ListCount - 1) = ItemData
            End If
            If Colocar = 1 Then
                TamTexto = TextCombo(Texto, Combo, 0)
            End If
        End If
    End If
End If
End Function

Private Function AcharTextoNaLista(ByRef rctlLista As Control, ByRef rstrTexto As String, Optional ByVal vblnAcharTextoExato As Boolean) As Long
'Função retirada do site http://www.fatorx.com.br em 18/03/2001

Dim lngMensagem As Long

' Verifica se o controle passado é do tipo 'ComboBox'.
If TypeOf rctlLista Is ComboBox Then
   ' Verifica se a função deve retornar um item da lista cujo
   ' valor corresponda exatamente ou parcialmente ao texto
   ' passado para a função.
   If vblnAcharTextoExato Then
      lngMensagem = CB_FINDSTRINGEXACT
   Else
      lngMensagem = CB_FINDSTRING
   End If
' Verifica se o controle passado é do tipo 'ListBox'.
ElseIf TypeOf rctlLista Is ListBox Then
   ' Verifica se a função deve retornar um item da lista cujo
   ' valor corresponda exatamente ou parcialmente ao texto
   ' passado para a função.
   If vblnAcharTextoExato Then
      lngMensagem = LB_FINDSTRINGEXACT
   Else
      lngMensagem = LB_FINDSTRING
   End If
Else
   ' O controle passado não é nem um 'ComboBox' nem um
   ' 'ListBox', portanto, a função não pode continuar.
   '2.7.2 Logar todas mensagens
   msgboxL "A função 'AcharTextoNaLista' não pode ser usada com um controle do tipo: """ & _
      TypeName(rctlLista) & """.", vbInformation
   Exit Function
End If

' Se a lista não possui nenhum item, a função não deve continuar.
If rctlLista.ListCount = 0 Then Exit Function
' Faz uma chamada à função 'SendMessage', indicando o hWnd do
' controle passado assim como a mensagem desejada (procura de
' texto parcial ou exata) e o texto que se deseja achar na lista.
' O valor -1 passado no parâmetro 'wParam' indica que a busca
' será feita em toda a lista desde o início da mesma.
AcharTextoNaLista = SendMessage( _
   rctlLista.hWnd, lngMensagem, -1, ByVal rstrTexto)
End Function

Public Function DtINIRel() As Date
Dim Ano As Integer
Dim Mes As Integer
Dim Aux As Integer
Dim Dia As Integer
Dim DT  As Date

DT = Int(Now)
Ano = Year(DT)
Mes = Month(DT) - 1
Dia = 1

If Mes < 1 Then
   Mes = 12
   Ano = Ano - 1
End If

FazData:
DT = DateSerial(Ano, Mes, Dia)
If Day(DT) < Dia Then
   Dia = Dia - 1
   GoTo FazData
End If

DtINIRel = DT
End Function

Public Sub PoemTituloAplicacao(F As Form)
With App
    AppVersao = .Major & "." & .Minor & "." & .Revision
    F.Caption = "Gerenciamento de Orçamentos " & AppVersao
    '2.5.6 Modo Restrito
    If INI.Restrito Then
        
        '2.6.0 Troca do título do modo restrito para 'modo Mecânico'
        F.Caption = F.Caption & " Modo Mecânico"
        'F.Caption = F.Caption & " Modo Restrito"
        
    End If
End With
End Sub

Public Sub ModoRestrito(Modo As Boolean)
Modo = Not Modo
'2.5.6 Modo Restrito
FrmMenu.MnuOrc(0).Visible = Modo
FrmMenu.MnuOrc(4).Visible = Modo
FrmMenu.MnuPagPagar.Visible = Modo
FrmMenu.MnuCli(0).Visible = Modo
FrmMenu.mnu.Visible = Modo

'2.6.1 Mecânicos não devem ver valores
FrmMenu.MnuOrc(1).Visible = Modo
FrmMenu.MnuOrc(2).Visible = Modo
FrmMenu.MnuOpcoes.Visible = Modo

'2.6.4 Impressão do recibo do mecânico
FrmMenu.Mnu_Mec.Visible = Modo
End Sub

Public Sub DesabilitaForm(F As Form)
'2.5.6 Modo Restrito
Dim C As Control

On Local Error Resume Next
For Each C In F
    If TypeOf C Is TextBox Then
        C.ReadOnly = True
    Else
        If Not (TypeOf C Is VScrollBar) Then
            C.Enabled = False
        End If
    End If
Next
On Local Error GoTo 0
End Sub

Public Function StringBoa(STR As String, Optional AceitaPonto As Boolean, Optional AceitaNulo As Boolean = False) As String
'Retorna apenas caracteres válidos
Dim a As Integer
Dim Tam As Integer
Dim L As Integer
Dim naBoa As Boolean
Dim LL As String * 1

Const PBoas = "0ÂÁÉÍÓÚÜÃÕ& -Ç()'"

naBoa = True
Tam = Len(STR)
For a = 1 To Tam
    LL = Mid(STR, a, 1)
    L = Asc(LL)
    If (L < 48 Or L > 122) And InStr(PBoas, LL) = 0 Then
        Select Case L
            Case 44, 46
                If AceitaPonto = False Then
                    naBoa = False
                End If
            Case Else
                If AceitaNulo = False Then
                    naBoa = False
                    Exit For
                End If
        End Select
    End If
Next
If naBoa = True Then
    StringBoa = STR
Else
    StringBoa = ""
End If
End Function

Public Function Busca_ItemData(Combo As ComboBox, Mec As Integer) As Integer
'2.7.4 Mecânicos dos Orçamentos
Dim a As Integer
Dim b As Integer

b = -1
For a = 0 To Combo.ListCount
    If Mec = Combo.ItemData(a) Then
        b = a
        Exit For
    End If
Next
Busca_ItemData = b
End Function

