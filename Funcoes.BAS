Attribute VB_Name = "Funcoes"
'4.8.4 ReImpressão dos Vales
'4.7.0 Retorno do menu para orçamentos mas sem acionar o orçamento
'4.6.9 Retirada de permissões de menus no perfil de mecânico
'4.6.5 Retirada da configuração do modo de impressão a nivel de banco de dados
'4.6.3 Tratamento para estouro de divisão em VeValor
'4.3.3 Previsão para o Windows Seven
'4.2.5 Recolocação da validação do CPNj
'4.2.4 Imprimir também o balconísta que criou o orçamento na impressão da observação
'4.2.3 Mais informação na impressão da observação
'4.2.2 Ajuste na Impressão da observação
'4.2.1 Impressão da observação
'4.1.9 Detecção para Sistema Operacional Windows Oito
'4.1.4 Campo para CGC e CPF no cadastro de clientes
'3.8.5 Previsão para o objeto PictureBox e TDBGrid na adaptação da resolução
'3.8.5 Retorno do botão de tarefas no modo mecânico
'3.8.4 Adaptação da resolução
'3.7.0 Ferramentas
'3.6.5 Identificação do modo no log
'3.5.1 Não excluir fisicamente Mecânico
'3.4.4 Vale Transporte
'3.3.5 Alteração da senha geral
'3.3.4 Nr de linhas apos a impressao passa a ser informação local
'3.2.1 Pesquisa de tarefas realizadas
'3.1.2 Permitir só receber pagamento de balconistas ou escritório
'3.0.5 Transações na gravação das tarefas
'2.9.6 Informação da verão do aplicativo no log
'2.8.1 Permitir adicionar tarefas pelo orçamento em dois PCs ao mesmo tempo
'2.7.5 Taréfas Dinâmicas
'2.7.4 Taréfas dos mecânicos
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

'Declarações para as funções de Combo
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_FINDSTRING = &H14C
Private Const CB_FINDSTRINGEXACT = &H158
Private Const LB_FINDSTRING = &H18F
Private Const LB_FINDSTRINGEXACT = &H1A2

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
Private Declare Function GetVersionExA Lib "KERNEL32" (lpVersionInformation As OSVERSIONINFO) As Long

'2.8.1 Permitir adicionar tarefas pelo orçamento em dois PCs ao mesmo tempo
Private Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

'3.8.4 Adaptação da resolução
Private Declare Function GetSystemMetrics Lib "USER32" (ByVal nIndex As Long) As Long

'Public Function VlrSql(Inf As String) As String
'Dim STR As String
'
'If Inf = "" Then
'    '2.1.5 Prever orçamento sem valor
'    VlrSql = "0"
'Else
'    STR = Valo(dado:=Inf)
'    STR = Replace(STR, ",", ".")
'    VlrSql = STR
'End If
'End Function

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
Base = App.Path & "\OrCarro.mdb"

CaminhoBkp = INI.CaminhoBkp
SetaDirBackup

'3.0.5 Transações na gravação das tarefas
Set WK = DBEngine.Workspaces(0)

Abre:
On Local Error GoTo SemOrCarroMDB
Set db = WK.OpenDatabase(Base)
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
    
    '4.6.5 Retirada da configuração do modo de impressão a nivel de banco de dados
    'INI.TpImpress = !TpImpress
    
    '3.3.4 Nr de linhas apos a impressao passa a ser informação local
    'INI.LinhasApos = !LinhasApos
    
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

'3.3.0 Critério de quantidade de carros para liberar as comissões
If Aux = 0 Or EhNulo = True Then 'não vi realmente motivos de mudar aqui, só achei que fiz errado da primeira vez
'If Aux = 0 And EhNulo = True Then

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
'Tipo é o alinhamento
'   0=Esquerda,
'   1=Central
'   2=Direita
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

'lpVersionInformation.dwPlatformId = 2
'lpVersionInformation.dwMajorVersion = 6
'lpVersionInformation.dwMinorVersion = 1

'4.3.3 Previsão para o Windows Seven
Select Case lpVersionInformation.dwPlatformId
    Case Is = Platform_Windows_32
        SistemaOperacional = "Windows 32"
    Case Is = Platform_Windows_95_98_ME
        Select Case lpVersionInformation.dwMinorVersion
            Case Is = 0
                SistemaOperacional = "Windows 95"
            Case Is = 10
                SistemaOperacional = "Windows 98"
            Case Is = 90
                SistemaOperacional = "Windows ME"
        End Select
    Case Is = Platform_Windows_NT_2K_XP
        Select Case lpVersionInformation.dwMajorVersion
            Case Is < 5
                SistemaOperacional = "Windows NT"
            Case Is = 5
                Select Case lpVersionInformation.dwMinorVersion
                    Case Is = 0
                        SistemaOperacional = "Windows 2K"
                    Case Is = 1
                        SistemaOperacional = "Windows XP"
                End Select
            Case Is = 6
                Select Case lpVersionInformation.dwMinorVersion
                    Case Is = 1
                        SistemaOperacional = "Seven"
                    Case Else
                        SistemaOperacional = "Windows Oito"
                End Select
        End Select
End Select
If SistemaOperacional = "" Then
    MsgBox "PlatformId = " & lpVersionInformation.dwPlatformId & vbNewLine & "MajorVersion = " & lpVersionInformation.dwMajorVersion & vbNewLine & "MinorVersion = " & lpVersionInformation.dwMinorVersion, vbOKOnly, "Sistema Operacional não previsto"
End If
'If (lpVersionInformation.dwPlatformId = Platform_Windows_32) Then
'  SistemaOperacional = "Windows 32"
'ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_95_98_ME) And (lpVersionInformation.dwMinorVersion = 0) Then
'  SistemaOperacional = "Windows 95"
'ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_95_98_ME) And (lpVersionInformation.dwMinorVersion = 10) Then
'  SistemaOperacional = "Windows 98"
'ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_95_98_ME) And (lpVersionInformation.dwMinorVersion = 90) Then
'  SistemaOperacional = "Windows ME"
'ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion < 5) Then
'  SistemaOperacional = "Windows NT"
'ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 5) And (lpVersionInformation.dwMinorVersion = 0) Then
'  SistemaOperacional = "Windows 2K"
'ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 5) And (lpVersionInformation.dwMinorVersion = 1) Then
'  SistemaOperacional = "Windows XP"
'
''4.1.9 Detecção para Sistema Operacional Windows Oito
'ElseIf (lpVersionInformation.dwPlatformId = Platform_Windows_NT_2K_XP) And (lpVersionInformation.dwMajorVersion = 6) And (lpVersionInformation.dwMinorVersion = 2) Then
'  SistemaOperacional = "Windows Oito"
'Else
'  MsgBox "PlatformId = " & lpVersionInformation.dwPlatformId & vbNewLine & "MajorVersion = " & lpVersionInformation.dwMajorVersion & vbNewLine & "MinorVersion = " & lpVersionInformation.dwMinorVersion, vbOKOnly, "Sistema Operacional não previsto"
'End If
  
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
        '2.7.5 Taréfas Dinâmicas
        Select Case Tipo
            Case vbInteger  '2
                SN = 0
            Case vbCurrency '6
                SN = 0
            Case vbDate     '7
                SN = 0
            Case vbString   '8
                SN = ""
            Case Else
                SN = ""
        End Select
'      If Tipo = vbInteger Then
'         SN = 0
'       ElseIf Tipo = vbString Then
'         SN = ""
'      End If

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

'3.8.4 Adaptação da resolução
Public Sub InicForm(F As Form, Optional SemAdapRes As Boolean = True, Optional MudaTodos As Boolean = False)
Colorizar F

If INI.Resolucao = 1 And SemAdapRes = True Then
   AdapForm F, 800, MudaTodos
End If
End Sub

'3.8.4 Adaptação da resolução
Private Sub Colorizar(F As Form)
Dim C As Control

On Local Error Resume Next
For Each C In F
    C.BackColor = INI.Cor
   
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

Public Sub AchaOrc(Selec As String, Optional Pesq$, Optional Placa As String)
Dim X       As Long
Dim SQL     As String
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
    
        '2.7.5 Taréfas Dinâmicas
        CarregaTarefas CadeOrc("Orcamento")
    
        Load frmOrc
        frmOrc.NrOrcamento = CadeOrc("Orcamento")
        frmOrc.Show
    Else
      'Mostrar um browse para os orçamentos
      Load Orcas
      If Pesq$ = "" Then
      
        '3.7.8 Troca de carros entre os clientes
         If Placa > "" Then
            Orcas.Placa = Placa
         Else
         
            Pesq$ = Chr$(34) + Trim$(CadeOrc("Cliente")) + Chr$(34)
        End If
      End If
      
      '3.7.8 Troca de carros entre os clientes
      If Placa = "" Then
      
        Orcas.nmCliente = Pesq$
      End If
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
         
            '4.6.3 Tratamento para estouro de divisão em VeValor
            On Local Error GoTo DivZero
            Vlr1 = Vlr1 / Vlr2
            On Local Error GoTo 0
            
      End Select
      If a < Len(Aux) Then
         Aux = STR$(Vlr1) + Mid$(Aux, a)
         Oper = ""
         Ped2 = ""
         GoTo Conta
      End If
End If

'4.6.3 Tratamento para estouro de divisão em VeValor
continuaVeValor:

If Algebra Then
   Aux = STR$(Vlr1)
   Vlr = Vlr1
Else
   Vlr = Valo(ctr, Aux)
End If
If Escr Then
    MostraValor ctr, Vlr
End If
Exit Sub

'4.6.3 Tratamento para estouro de divisão em VeValor
DivZero:
Vlr1 = 0
Resume continuaVeValor
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
    #If V272 > 0 Then
        AppVersao = "272-3"
    #Else
        AppVersao = .Major & "." & .Minor & "." & .Revision
    #End If
    
    #If Beta = 1 Then
        F.Caption = F.Caption & " " & AppVersao & ".1"
    #ElseIf Beta = 2 Then
        F.Caption = F.Caption & " " & AppVersao & ".2 "
    #ElseIf Beta = 3 Then
        F.Caption = F.Caption & " " & AppVersao & ".3 "
    #ElseIf Beta = 4 Then
        F.Caption = F.Caption & " " & AppVersao & ".4 "
    #ElseIf Beta = 5 Then
        F.Caption = F.Caption & " " & AppVersao & ".5 "
    #ElseIf Beta = 6 Then
        F.Caption = "ORCARRO Beta 6 "
    #Else
        F.Caption = "Gerenciamento de Orçamentos " & AppVersao
    #End If
    
    '3.1.0 Modo Balcão
    Select Case INI.ModoOperacao
        Case tpMecanico
            F.Caption = F.Caption & " Modo Mecânico"
        Case tpBalcao
            F.Caption = F.Caption & " Modo Balcão"
    End Select
'    '2.5.6 Modo Restrito
'    If INI.Restrito Then
'
'        '2.6.0 Troca do título do modo restrito para 'modo Mecânico'
'        F.Caption = F.Caption & " Modo Mecânico"
'        'F.Caption = F.Caption & " Modo Restrito"
'
'    End If
    
    #If Beta = 1 Then
        F.Caption = F.Caption & " - Para testes"
    #End If
    
End With

'2.9.6 Informação da verão do aplicativo no log
Loga "VersãoEXE: " & AppVersao, lDBG
End Sub

'3.1.0 Modo Balcão
Public Sub DefineMenus(Modo As Integer)

'Public Sub ModoRestrito(Modo As Boolean)
Select Case INI.ModoOperacao
    Case 0  'ESCRITÓRIO
    
        '3.6.5 Identificação do modo no log
        Loga "Modo Escritório"
        
        FrmMenu.Mnu_Mec.Caption = "Operadores"
        FrmMenu.Mnu_Mec.Visible = True
        FrmMenu.btTarefa.Visible = False
        FrmMenu.MnuOpcoes.Visible = True
        FrmMenu.mnu.Visible = True
        FrmMenu.MnuOrc(4).Visible = True
        FrmMenu.MnuPagPagar.Visible = True
        
        '3.2.1 Pesquisa de tarefas realizadas
        FrmMenu.MnuMec(4).Visible = True
        
        '3.2.7 Gravação do valor da mão de obra
        FrmMenu.mnuTabela.Visible = False
        
        '3.4.4 Vale Transporte
        FrmMenu.mnuRec(1).Enabled = INI.UtComissoes
        
        '3.7.0 Ferramentas
        FrmMenu.MnuFerramentas.Visible = True
        
        FrmMenu.MnuCli(0).Visible = True
        
    Case 1  'BALCÃO
    
        '3.6.5 Identificação do modo no log
        Loga "Modo Balcão"
        
        FrmMenu.Mnu_Mec.Visible = False
        FrmMenu.btTarefa.Visible = False
        FrmMenu.MnuOpcoes.Visible = False
        FrmMenu.mnu.Visible = False
        FrmMenu.MnuPagPagar.Visible = False
        FrmMenu.MnuOrc(0).Visible = True
        FrmMenu.MnuOrc(1).Visible = True
        FrmMenu.MnuOrc(2).Visible = True
        FrmMenu.MnuOrc(4).Visible = False
        
        '3.2.1 Pesquisa de tarefas realizadas
        FrmMenu.MnuMec(4).Visible = False
        
        '3.2.7 Gravação do valor da mão de obra
        FrmMenu.mnuTabela.Visible = True
        
        '3.7.0 Ferramentas
        FrmMenu.MnuFerramentas.Visible = False
        
    Case 2  'MECÂNICO
    
        '3.6.5 Identificação do modo no log
        Loga "Mecânico"
        FrmMenu.Mnu_Mec.Caption = "Mecânicos"
        
        FrmMenu.MnuPagPagar.Visible = False
        FrmMenu.MnuCli(0).Visible = False
        FrmMenu.mnu.Visible = False
        FrmMenu.MnuOrc(0).Visible = False
        FrmMenu.MnuOrc(1).Visible = False
        FrmMenu.MnuOrc(2).Visible = False
        FrmMenu.MnuOrc(4).Visible = False
        FrmMenu.MnuOpcoes.Visible = False
        FrmMenu.MnuMec(0).Visible = False
        FrmMenu.MnuMec(1).Visible = False
        FrmMenu.MnuMec(2).Visible = False
        FrmMenu.btTarefa.Visible = True
        FrmMenu.mnu.Visible = False
        
        '3.2.7 Gravação do valor da mão de obra
        FrmMenu.mnuTabela.Visible = False
        
        '4.6.9 Retirada de permissões de menus no perfil de mecânico
        FrmMenu.MnuRelComiss.Visible = False
        FrmMenu.MnuProdutividade.Visible = False
        
        '4.7.0 Retorno do menu para orçamentos mas sem acionar o orçamento
        'FrmMenu.Mnu_Orc.Visible = False
        
        '3.7.0 Ferramentas
        FrmMenu.MnuFerramentas.Visible = False
        
        '3.8.5 Retorno do botão de tarefas no modo mecânico
        FrmMenu.btTarefa.Visible = True
        
End Select
'Modo = Not Modo
''2.5.6 Modo Restrito
'FrmMenu.MnuOrc(0).Visible = Modo
'FrmMenu.MnuOrc(4).Visible = Modo
'FrmMenu.MnuPagPagar.Visible = Modo
'FrmMenu.MnuCli(0).Visible = Modo
'FrmMenu.mnu.Visible = Modo
'
''2.6.1 Mecânicos não devem ver valores
'FrmMenu.MnuOrc(1).Visible = Modo
'FrmMenu.MnuOrc(2).Visible = Modo
'FrmMenu.MnuOpcoes.Visible = Modo
'
''2.7.4 Taréfas dos mecânicos
'FrmMenu.MnuMec(0).Visible = Modo
'FrmMenu.MnuMec(1).Visible = Modo
'FrmMenu.MnuMec(2).Visible = Modo
'2.6.4 Impressão do recibo do mecânico
'FrmMenu.Mnu_Mec.Visible = Modo
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

'4.8.4 ReImpressão dos Vales
Public Function MesExtenso(Data As Date) As String
'Public Function MesExtenso() As String
Dim iMes As Integer

'4.8.4 ReImpressão dos Vales
iMes = Month(Data)
'iMes = Month(Now)

Select Case iMes
    Case 1
        MesExtenso = "Janeiro"
    Case 2
        MesExtenso = "Fevereiro"
    Case 3
        MesExtenso = "Março"
    Case 4
        MesExtenso = "Abril"
    Case 5
        MesExtenso = "Maio"
    Case 6
        MesExtenso = "Junho"
    Case 7
        MesExtenso = "Julho"
    Case 8
        MesExtenso = "Agosto"
    Case 9
        MesExtenso = "Setembro"
    Case 10
        MesExtenso = "Outubro"
    Case 11
        MesExtenso = "Novembro"
    Case Else
        MesExtenso = "Dezembro"
End Select
End Function

Public Function NameOfPC() As String
'2.8.1 Permitir adicionar tarefas pelo orçamento em dois PCs ao mesmo tempo
Dim PCName As String
Dim P As Long
Dim NameSize As Long
Dim X As Long

PCName = Space$(16)
NameSize = Len(PCName)
X = GetComputerName(PCName, NameSize)
PCName = Trim(PCName)
PCName = Replace(Trim(PCName), Chr(0), "")
NameOfPC = PCName
End Function

Public Function LetraMenosUm(Palavra As String) As String
LetraMenosUm = Left(Palavra, Len(Palavra) - 1)
End Function

'3.1.2 Permitir só receber pagamento de balconistas ou escritório
Public Function BuscaMec(Codigo As String, Optional Tipo As String = "") As Recordset
'2.7.4 Registro facilitado das taréfas dos mecânicos
Dim SQL As String

'3.5.1 Não excluir fisicamente Mecânico
SQL = "Select codi, Nome From Mecanicos Where Senha = '" & Codigo & "'" & Tipo & " and Ativo = True "
'SQL = "Select codi, Nome From Mecanicos Where Senha = '" & codigo & "'" & Tipo

AbreTB BuscaMec, SQL, dbOpenDynaset
End Function

Public Function NumDiaSem() As String
'3.3.5 Alteração da senha geral
NumDiaSem = Trim(STR(Weekday(Now)))
End Function

Public Function CriticaData(TextBox As String, Data As Date) As Integer
'Adaptado de Alexandre Balester - MavData - TRC.MAK
'<<< 100 Passos na condição normal, incluindo Sub's
Dim tamAux As Integer, Mensagem As String, crlf As String
Dim auxiliar$, auxiliar1$, auxiliar11$, auxiliar12$, Caracter$
Dim Desloc As Integer, separador_1 As Integer, separador_2 As Integer
Dim X As VbMsgBoxResult
Dim flag As Integer

Dim Dia As Integer
Dim Mes As Integer
Dim Ano As Integer

CriticaData = -1
flag = -1
auxiliar$ = TextBox

tamAux = Len(auxiliar$)

'* Numeros
For Desloc = 1 To tamAux
        Caracter$ = Mid$(auxiliar$, Desloc, 1)
        'filtra lixo
        If (Caracter$ < "0" Or Caracter$ > "9") And Caracter$ <> "." And Caracter$ <> "/" And Caracter$ <> "-" And Caracter$ <> "," Then
                Mensagem = crlf + "Digite apenas NUMEROS !"
                GoSub PROBLEMAS
                Exit For
        End If
        'identifica separadores (/,.-)
        If Caracter$ = "." Or Caracter$ = "/" Or Caracter$ = "-" Or Caracter$ = "," Then
                If separador_1 = 0 Then
                        separador_1 = Desloc
                ElseIf separador_2 = 0 Then
                        separador_2 = Desloc
                        If separador_2 - separador_1 = 2 Then
                                Mid$(auxiliar$, Desloc, 1) = Mid$(auxiliar$, Desloc - 1, 1)
                                Mid$(auxiliar$, Desloc - 1, 1) = "0"
                                Caracter$ = Mid$(auxiliar$, Desloc, 1)
                                If Len(auxiliar1$) > 0 Then
                                        Mid$(auxiliar1$, Len(auxiliar1$), 1) = "0"
                                End If
                        End If
                Else
                        Mensagem = crlf + "Formato invalido !" + crlf + "(Formato DDMMAA)"
                        GoSub PROBLEMAS
                        Exit For
                End If
        End If
        'filtra so numeros
        If Caracter$ >= "0" And Caracter$ <= "9" Then
                auxiliar1$ = auxiliar1$ + Caracter$
        End If
Next Desloc
If Not flag Then Exit Function

        '* controle de preenchimento
tamAux = Len(auxiliar1$)
If tamAux <> 8 Then
    CriticaData = 0
Else
    '* Dia, mes e ano
    Dia = Val(Left$(auxiliar1$, 2))
    Mes = Val(Mid$(auxiliar1$, 3, 2))
    Ano = Val(Right$(auxiliar1$, 4))
    If Ano < 1900 Or Ano > 2099 Then
    
            Mensagem = crlf + "Ano invalido !" + crlf + "(Formato DDMMAAAA)"
            GoSub PROBLEMAS
            Exit Function
    End If
    If Mes < 1 Or Mes > 12 Then
            Mensagem = crlf + "Mes invalido !" + crlf + "(Formato DDMMAAAA)"
            GoSub PROBLEMAS
            Exit Function
    End If
    If Dia < 1 Or Dia > MaxDiaMes(Mes, Ano) Then
            Mensagem = crlf + "Dia inexistente"
            GoSub PROBLEMAS
            Exit Function
    End If
    Data = DateSerial(Ano, Mes, Dia)
End If
Exit Function

PROBLEMAS:
Beep
X = MsgBox(Mensagem, 48, "ERRO")
CriticaData = 0
flag = 0
'TextBox.SetFocus
Return
End Function

Private Function MaxDiaMes(Mes As Integer, Ano As Integer)
Dim Aux As Integer, Aux2 As Integer

If Mes < 1 Or Mes > 12 Then
        MaxDiaMes = 30
        Exit Function
End If
Aux = Mes
Aux2 = Ano
Aux = Aux + 1
If Aux = 13 Then
        Aux = 1
        Aux2 = Aux2 + 1
End If
MaxDiaMes = Day(DateSerial(Aux2, Aux, 1) - 1)
End Function

'3.8.4 Adaptação da resolução
Private Function AdapForm(F As Form, Ideal As Integer, Optional MudaTodos As Boolean = False)
Dim xRes As Integer
Dim a    As Integer
Dim Prop As Single
Dim PropOrig As Single
Dim C As Control
Dim Alterar As Integer

Loga "Entrou em AdapForm"
xRes = GetSystemMetrics(0)
Loga "GetSystemMetrics"
If xRes > Ideal Then
   PropOrig = xRes / Ideal
   
    '8.2.7 Previsão para monitor em wide-screen, na adaptação do tamanho da tela
    If xRes > 1024 Then
        PropOrig = PropOrig / (xRes / 1024)
    End If
         
   For Each C In F
   
      Debug.Print C.Name
            
      Prop = PropOrig
      
      If MudaTodos Then
        Alterar = 1
      Else
        Alterar = 0
      End If
      
      If TypeOf C Is Image Then Alterar = 1
      #If SemProgressBar = 0 Then
         If TypeOf C Is ProgressBar Then Alterar = 1
      #End If
      If TypeOf C Is VScrollBar Then Alterar = 1
      If TypeOf C Is Label Then Alterar = 2
          
      '7.2.8 Log na tela
      If TypeOf C Is ListBox Then
        Alterar = 1
        Prop = Prop * 1.2
      End If
            
      '4.3.4 Realçado o link da página e mail no sobre
      #If SemLabelLink = 0 Then
'         If TypeOf C Is LabelLink Then
'            Alterar = 2
'         End If
      #End If
            
      If TypeOf C Is TextBox Then Alterar = 2
      If TypeOf C Is ComboBox Then Alterar = 2
      If TypeOf C Is CommandButton Then Alterar = 2
      If TypeOf C Is OptionButton Then Alterar = 2
      If TypeOf C Is CheckBox Then Alterar = 2
      If TypeOf C Is Frame Then Alterar = 2
'      If TypeOf C Is StatusBar Then Alterar = 1

      If TypeOf C Is Toolbar Then Alterar = 2
      
      #If SemMSFlexGrid = 0 Then
         If TypeOf C Is MSFlexGrid Then Alterar = 2
      #End If
      '#If SemSSTab = 0 Then
         'If TypeOf C Is SSTab Then Alterar = 2
      '#End If
            
      If TypeOf C Is DataControl Then
        Alterar = 1
      End If
      
      '3.8.5 Previsão para o objeto PictureBox e TDBGrid na adaptação da resolução
      If TypeOf C Is PictureBox Then Alterar = 1
      If TypeOf C Is TDBGrid Then Alterar = 1
      
      If Alterar Then
         C.Width = C.Width * PropOrig
         C.Left = C.Left * PropOrig
         
         C.Top = C.Top * PropOrig
         
         '4.4.5 Correção do redimensionamento do objeto combo
         If Alterar = 2 Then
         
            If C.Tag <> "NãoAlterar" Then
                C.FontSize = C.FontSize * PropOrig
            End If
         End If
         On Error GoTo NaoVai
         C.Height = C.Height * Prop
'         If Alterar = 2 Then
'            C.FontSize = C.FontSize * Prop
'         End If
Continua_AdapForm:
      End If
   Next
   F.Width = F.Width * Prop
   F.Height = F.Height * Prop
End If
Loga "Saiu do AdapForm"
Exit Function

NaoVai:
Resume Continua_AdapForm
End Function

Public Function CriticaEmail(sEMail As String) As Boolean
Dim ValidEMail As Boolean
Dim nCharacter As Integer
Dim Count      As Integer
Dim sLetra     As String

ValidEMail = False

'Verifica se o e-mail tem no MÍNIMO 5 caracteres (a@b.c)
If Len(sEMail) < 5 Then
    'O e-mail é inválido, pois tem menos de 5 caracteres
    MsgBox "O e-mail digitado tem menos de 5 caracterec!!!"
    Exit Function
End If

'Verificar a existencia de arrobas (@) no e-mail
For nCharacter = 1 To Len(sEMail)
    If Mid(sEMail, nCharacter, 1) = "@" Then
        Count = Count + 1
    End If
Next
'Verifica o número de arrobas.
If Count <> 1 Then
    'O e-mail é inválido, pois tem 0 ou mais de 1 arroba
    MsgBox "O nº de arrobas (@) do e-mail é inválido!!!"
    Exit Function
Else
    'O e-mail tem 1 arroba.
    'Verificar a posição da arroba

    If InStr(sEMail, "@") = 1 Then
        'O e-mail é inválido, pois começa com uma @
        MsgBox "O e-mail foi iniciado com uma arroba (@)!!!"
        Exit Function

    ElseIf InStr(sEMail, "@") = Len(sEMail) Then
        'O e-mail é inválido, pois termina com uma @
        MsgBox "O e-mail termina com uma arroba (@)!!!"
        Exit Function
   End If
End If

nCharacter = 0
Count = 0
'Verificar a existencia de pontos (.) no e-mail

For nCharacter = 1 To Len(sEMail)
    If Mid(sEMail, nCharacter, 1) = "." Then
        Count = Count + 1
    End If
Next
'Verifica o número de pontos.
'TEM que ter PELO MENOS UM ponto.

If Count < 1 Then
    'O e-mail é inválido, pois não tem pontos.
    MsgBox "O e-mail é inválido, pois não contém pontos (.)!!!"
    Exit Function
Else
    'O e-mail tem pelo menos 1 ponto.
    'Verificar a posição do ponto:

    If InStr(sEMail, ".") = 1 Then
        'O e-mail é inválido, pois começa com um ponto
        MsgBox "O e-mail foi iniciado com um ponto (.)!!!"
        Exit Function
    ElseIf InStr(sEMail, ".") = Len(sEMail) Then
        'O e-mail é inválido, pois termina com um ponto.
        MsgBox "O e-mail termina com um ponto (.)!!!"
        Exit Function
    ElseIf InStr(InStr(sEMail, "@"), sEMail, ".") = 0 Then
        'O e-mail é inválido, pois termina com um ponto.
        MsgBox "O e-mail não tem nenhum ponto (.) após a arroba (@)!!!"
        Exit Function
    End If
End If
nCharacter = 0
Count = 0
'Verifica se o e-mail não tem pontos
'consecutivos (..) após a arroba (@).

If InStr(sEMail, "..") > InStr(sEMail, "@") Then
    'O e-mail é inválido, tem pontos consecutivos após o @.
    MsgBox "O e-mail contém pontos consecutivos (..) após o arroba (@)!!!"
    Exit Function
End If

'Verifica se o e-mail tem caracteres inválidos
For nCharacter = 1 To Len(sEMail)
    sLetra = Mid$(sEMail, nCharacter, 1)
    If Not (LCase(sLetra) Like "[a-z]" Or sLetra = "@" Or sLetra = "." Or sLetra = "-" Or sLetra = "_" Or IsNumeric(sLetra)) Then
        'O e-mail é inválido, pois tem caracteres inválidos
        MsgBox "Foi digitado um caracter inválido no e-mail!!!"
        Exit Function
    End If
Next
nCharacter = 0
'Bem, se a verificação chegou até aqui é porque o e-mail é válido, então...
CriticaEmail = True
End Function

Public Function OperadorQueGravou(ByRef Nome As String, Optional Escrit As Boolean = False) As Integer
Dim Result As Boolean
Dim Resp   As String
Dim Oper$
Dim rsMec  As Recordset

Load frmSenha
frmSenha.Tipo = 1
frmSenha.Show 1
If frmSenha.Resultado = True Then
    Resp = frmSenha.Senha
End If
If Resp > "" Then
    If Escrit Then
        Oper$ = " and Oper = 2 "
    Else
        Oper$ = " and Oper > 0 "
    End If
    Set rsMec = BuscaMec(Resp, Oper$)
    If rsMec.EOF Then
        MsgBox "Operador não identificado"
        Exit Function
    Else
        Nome = rsMec!Nome
        OperadorQueGravou = rsMec!codi
    End If
End If
Unload frmSenha
End Function

''4.1.4 Campo para CGC e CPF no cadastro de clientes
'Function CheckCGCCPF(Vl_CgcCpf As String) As Boolean
'' Esta Rotina Devolverá True  Se o Cgc/Cpf Informado For valido
''                    ou False Se o Cgc/Cpf Não For Correto
'' Para Chamar esta Rotina de Consistência
'' 1 ) Atribuir o valor do CgcCpf a uma Variavel String
'' 2 ) Chamar a Rotina com : Fu_consistir_CgcCpf (Variavel)
'' Uma Forma Simples de fazer a Consistencia
' ' é Copiando as linhas abaixo (exemplo)
'' para dentro do Programa
'' Dim Vl_CgcCpf As String
'' Vl_CgcCpf = Me.CgcCpf.Text
'' If Fu_consistir_CgcCpf(Vl_CgcCpf) = False then
''  MsgBox "( Cgc/Cpf Informado Não é um Cgc/Cpf Correto )"
''  Me.CgcCpf.SetFocus
''  Exit Sub
'' End if
'
'Static Numero(15)       As Integer
'Dim VA_Resto            As Integer
'Dim VA_Resultado        As Integer
'Dim VA_SomaDigito10     As Integer
'Dim VA_resto1           As Integer
'Dim a                   As Integer
'Dim Soma                As Integer
'Dim VA_CgcCpf           As String
'Dim VA_Digito           As String
'CheckCGCCPF = False
'
'Vl_CgcCpf = Replace(Vl_CgcCpf, ".", "")
'Vl_CgcCpf = Replace(Vl_CgcCpf, "-", "")
'Vl_CgcCpf = Replace(Vl_CgcCpf, "/", "")
'
'VA_CgcCpf = Format(Vl_CgcCpf, "@@@@@@@@@@@@@@")
'VA_Digito = Mid(VA_CgcCpf, 13, 2)
'
'Numero(1) = Val(Mid(VA_CgcCpf, 1, 1))
'Numero(2) = Val(Mid(VA_CgcCpf, 2, 1))
'Numero(3) = Val(Mid(VA_CgcCpf, 3, 1))
'Numero(4) = Val(Mid(VA_CgcCpf, 4, 1))
'Numero(5) = Val(Mid(VA_CgcCpf, 5, 1))
'Numero(6) = Val(Mid(VA_CgcCpf, 6, 1))
'Numero(7) = Val(Mid(VA_CgcCpf, 7, 1))
'Numero(8) = Val(Mid(VA_CgcCpf, 8, 1))
'Numero(9) = Val(Mid(VA_CgcCpf, 9, 1))
'Numero(10) = Val(Mid(VA_CgcCpf, 10, 1))
'Numero(11) = Val(Mid(VA_CgcCpf, 11, 1))
'Numero(12) = Val(Mid(VA_CgcCpf, 12, 1))
'Numero(13) = Val(Mid(VA_CgcCpf, 13, 1))
'Numero(14) = Val(Mid(VA_CgcCpf, 14, 1))
'
'For a = 1 To 14
'    Soma = Soma + Numero(a)
'Next
'If Soma = 0 Then
'    Exit Function
'End If
'
'  ' Cpf
'    VA_Resultado = (Numero(4) * 1) + _
'     (Numero(5) * 2) + (Numero(6) * 3) _
'     + (Numero(7) * 4) + (Numero(8) * 5) _
'     + (Numero(9) * 6) + (Numero(10) * 7) _
'     + (Numero(11) * 8) + (Numero(12) * 9)
'    VA_Resto = VA_Resultado Mod 11
'    If VA_Resto > 9 Then
'        VA_resto1 = VA_Resto - 10
'    Else
'        VA_resto1 = VA_Resto
'    End If
'    If VA_resto1 <> Numero(13) Then
'        GoTo CPF_Errado
'    End If
'
'    VA_Resultado = (Numero(5) * 1) _
'     + (Numero(6) * 2) + (Numero(7) * 3) _
'     + (Numero(8) * 4) + (Numero(9) * 5) + _
'     (Numero(10) * 6) + (Numero(11) * 7) + _
'     (Numero(12) * 8) + (VA_Resto * 9)
'    VA_Resto = VA_Resultado Mod 11
'    If VA_Resto > 9 Then
'        VA_resto1 = VA_Resto - 10
'    Else
'        VA_resto1 = VA_Resto
'    End If
'    If VA_resto1 <> Numero(14) Then
'        GoTo CPF_Errado
'    End If
'
'CheckCGCCPF = True
'Exit Function
'
'CPF_Errado:
'If CNPJVALIDO(Vl_CgcCpf) Then
'    CheckCGCCPF = True
'End If
'End Function

Public Function CNPJVALIDO(NUMERO_CNPJ As String) As Boolean
Dim CNPJ As String
Dim a As Integer
Dim J As Integer
Dim i As Integer
Dim D1 As Integer
Dim D2 As Integer
Dim X As Integer
Dim texte As String

X = 1
Do While X <= Len(Trim(NUMERO_CNPJ)) 'Faça enquanto x for menor ou igual ao tanto da caixa de texto.
    texte = Mid(NUMERO_CNPJ, X, 1) 'A variavel vai receber cada valor da caixa de texto.
    If IsNumeric(texte) Then 'Se a variavel texte for número.
        CNPJ = CNPJ & texte 'CNPJ recebe, ela mesma mais a variavel texte.
    End If 'Fim do se.
    X = X + 1 'A variavel x vai aumentando mais 1.
Loop 'Fim do enquanto.

If Len(CNPJ) = 8 And Val(CNPJ) > 0 Then
    a = 0
    J = 0
    D1 = 0
    For i = 1 To 7
        a = Val(Mid(CNPJ, i, 1))
        If (i Mod 2) <> 0 Then a = a * 2
        If a > 9 Then
            J = J + Int(a / 10) + (a Mod 10)
        Else
            J = J + a
        End If
    Next i
    D1 = IIf((J Mod 10) <> 0, 10 - (J Mod 10), 0)
    If D1 = Val(Mid(CNPJ, 8, 1)) Then
        CNPJVALIDO = True
    Else
        CNPJVALIDO = False
    End If
Else
    If Len(CNPJ) = 14 And Val(CNPJ) > 0 Then
        a = 0
        i = 0
        D1 = 0
        D2 = 0
        J = 5
        For i = 1 To 12 Step 1
            a = a + (Val(Mid(CNPJ, i, 1)) * J)
            J = IIf(J > 2, J - 1, 9)
        Next i
        a = a Mod 11
        D1 = IIf(a > 1, 11 - a, 0)
        a = 0
        i = 0
        J = 6
        For i = 1 To 13 Step 1
            a = a + (Val(Mid(CNPJ, i, 1)) * J)
            J = IIf(J > 2, J - 1, 9)
        Next i
        a = a Mod 11
        D2 = IIf(a > 1, 11 - a, 0)
        If (D1 = Val(Mid(CNPJ, 13, 1)) And D2 = Val(Mid(CNPJ, 14, 1))) Then
            CNPJVALIDO = True
        Else
            CNPJVALIDO = False
        End If
    Else
        CNPJVALIDO = False
    End If
End If
End Function

Function calculacpf(CPF As String) As Boolean
'Esta rotina foi adaptada da revista Fórum Access
'peguei do Marcoratti => http://www.macoratti.net/val_cpf.htm
On Error GoTo Err_CPF
Dim i As Integer 'utilizada nos FOR... NEXT
Dim strcampo As String 'armazena do CPF que será utilizada para o cálculo
Dim strCaracter As String 'armazena os digitos do CPF da direita para a esquerda
Dim intNumero As Integer 'armazena o digito separado para cálculo (uma a um)
Dim intMais As Integer 'armazena o digito específico multiplicado pela sua base
Dim lngSoma As Long 'armazena a soma dos digitos multiplicados pela sua base(intmais)
Dim dblDivisao As Double 'armazena a divisão dos digitos*base por 11
Dim lngInteiro As Long 'armazena inteiro da divisão
Dim intResto As Integer 'armazena o resto
Dim intDig1 As Integer 'armazena o 1º digito verificador
Dim intDig2 As Integer 'armazena o 2º digito verificador
Dim strConf As String 'armazena o digito verificador

lngSoma = 0
intNumero = 0
intMais = 0
strcampo = Left(CPF, 9)

'Inicia cálculos do 1º dígito
For i = 2 To 10
    strCaracter = Right(strcampo, i - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * i
    lngSoma = lngSoma + intMais
Next i
dblDivisao = lngSoma / 11

lngInteiro = Int(dblDivisao) * 11
intResto = lngSoma - lngInteiro
If intResto = 0 Or intResto = 1 Then
    intDig1 = 0
Else
    intDig1 = 11 - intResto
End If

strcampo = strcampo & intDig1 'concatena o CPF com o primeiro digito verificador
lngSoma = 0
intNumero = 0
intMais = 0
'Inicia cálculos do 2º dígito
For i = 2 To 11
    strCaracter = Right(strcampo, i - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * i
    lngSoma = lngSoma + intMais
Next i
dblDivisao = lngSoma / 11
lngInteiro = Int(dblDivisao) * 11
intResto = lngSoma - lngInteiro
If intResto = 0 Or intResto = 1 Then
    intDig2 = 0
Else
    intDig2 = 11 - intResto
End If
strConf = intDig1 & intDig2
'Caso o CPF esteja errado dispara a mensagem
If strConf <> Right(CPF, 2) Then
    
    '4.2.5 Recolocação da validação do CPNj
    calculacpf = CNPJVALIDO(CPF)
    'calculacpf = False
    
Else
    calculacpf = True
End If
Exit Function

Exit_CPF:
    Exit Function
Err_CPF:
    MsgBox Error$
    Resume Exit_CPF
End Function

'4.2.4 Imprimir também o balconísta que criou o orçamento na impressão da observação
Public Sub ImpressaoObs(Modelo As String, Cor As String, Placa As String, Obs As String, Cliente As String, Data As Date, Balconista As String)
'4.2.3 Mais informação na impressão da observação
'Public Sub ImpressaoObs(Modelo As String, Cor As String, Placa As String, Obs As String, Cliente As String, Data As Date)
'4.2.1 Impressão da observação
'Public Sub ImpressaoObs(Modelo As String, Cor As String, Placa As String, Obs As String)

'Vend

Dim TemMais As Boolean
Dim Tam     As Integer
Dim PosE    As Integer
Dim Pedaco  As String
Dim sData   As String

ImprBuferizada_Inicializa
LPRINT Chr$(15)

'4.2.3 Mais informação na impressão da observação
sData = Format(Data, "dd/mm/yyyy hh:mm")
If Right(sData, 5) = "00:00" Then
    sData = Left(sData, 10)
End If
LPRINT "DATA: " & sData

'4.2.4 Imprimir também o balconísta que criou o orçamento na impressão da observação
If Balconista > "" Then
    LPRINT "BALCONISTA: " & Balconista
End If

LPRINT "CLIENTE: " & Cliente
LPRINT "PLACA: " & Placa
LPRINT "CARRO: " & Modelo & " " & Cor
'LPRINT Format(Now, "dd/mm/yyyy")
'LPRINT Modelo & " " & Cor & " " & Placa

LPRINT " "
TemMais = True
Do
    Pedaco = Left(Obs, TamFita)
    PosE = InStr(Pedaco, Chr(13))
    If PosE Then
        Pedaco = Left(Pedaco, PosE - 1)
        Obs = Mid(Obs, PosE + 2)
        TemMais = True
    Else
        Pedaco = Obs
        Obs = Mid(Obs, 1, TamFita)
        TemMais = False
    End If
    LPRINT Pedaco
    Tam = Len(Obs)
    
'4.2.2 Ajuste na Impressão da observação
Loop While TemMais = True
'Loop While Tam > TamFita

'4.2.3 Mais informação na impressão da observação
LPRINT " "
LPRINT String(TamFita, "-")
'LPRINT

ImprBuferizada_Finaliza
End Sub
