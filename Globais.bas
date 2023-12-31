Attribute VB_Name = "Globais"
'5.0.4 Correções de bugs
'4.2.1 Impressão da observação
'4.0.6 Acabamento da alteração da identificação do operador que lança a peça, no tocando a deleção de peças
'4.0.2 Identificar quem editou o item do orçamento
'3.8.8 Data dos itens de orçamento
'3.0.5 Transações na gravação das tarefas
'2.8.4 Gravar no log o tempo de carregamento da tela de orçamento
'2.7.5 Taréfas Dinâmicas

Option Explicit

Public AppVersao      As String * 5
Public Base           As String
Public LinhaDeComando As String
Public CaminhoBkp     As String
Public AplicDirNat    As String
Public db             As Database
Public OrcasRecordset As Recordset
Public INI            As New clsReg
Public clsCLi         As New clsClientes

Public CorSelec As Long

'2.3.3 Possibilitar retornar o carro ao cliente
Public GCliente As String
Public GGPlaca  As String

'2.7.5 Taréfas Dinâmicas
Public Const SemMec = "Sem Mecânico"

Type ItOrc
    Peca As String
    Quant As Single
    Valor As Currency
    
    '3.8.8 Data dos itens de orçamento
    Data As Date
    
    '4.0.2 Identificar quem editou o item do orçamento
    ID As Long
    Alterado As Boolean
    Nome As String
    
    '4.0.6 Acabamento da alteração da identificação do operador que lança a peça, no tocando a deleção de peças
    Deletar As Boolean
    Existente As Boolean
End Type

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
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal LPString As Any, ByVal lpFileName As String) As Long

'2.8.4 Gravar no log o tempo de carregamento da tela de orçamento
Public MomChamOrc As Date

'3.0.5 Transações na gravação das tarefas
Public WK   As Workspace

Public Enum tpRec
    tpAdiantamento = 0
    tpComissao = 1
    tpValeTransp = 2
    tpPagamento = 3
    tpOutros = 4
End Enum
Public gTipo As tpRec

'4.9.1 Ajuste da Impressão em Fita via Usb
Public TamFita As Integer
'4.2.1 Impressão da observação
'Public Const TamFita = 55

'4.8.6 Edição de recibos
Public TpRecs(4) As String

'5.0.4 Correções de bugs
Public gbNrOrc      As Integer

