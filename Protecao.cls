VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProtecao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'2.0.5 Implantação do Log
'2.0.0 Implantação da proteção

Option Explicit

'Private Const READ_CONTROL = &H20000
'Private Const KEY_QUERY_VALUE = &H1
'Private Const KEY_SET_VALUE = &H2
'Private Const KEY_CREATE_SUB_KEY = &H4
'Private Const KEY_ENUMERATE_SUB_KEYS = &H8
'Private Const KEY_NOTIFY = &H10
'Private Const KEY_CREATE_LINK = &H20
'Private Const REG_SZ = 1
'Private Const REG_DWord = 2
'Private Const REG_BINARY = 3
'Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
'Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
'Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
'Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
'Private Const HKEY_LOCAL_MACHINE = &H80000002
'Private Const HKEY_CURRENT_USER = &H80000001
'Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

Public Senha         As String
Public ContraSenha   As String
Private Key          As Long
Private CaminhoChave As String
Private lcCliente    As String

Public Function VeRegistro() As Boolean
Dim i          As Integer
Dim Decomposto As Integer
Dim iSenha     As Integer
Dim rc         As Long
Dim hKey       As Long
Dim KeyValSize As Long
Dim tmpVal     As String
Dim SubChave   As String
Dim sSenha     As String
Dim KeyVal     As String
Dim Aux        As String

Const ERROR_SUCCESS = 0
Const KeyValType = 0

'2.0.5 Implantação do Log
rc = RegOpenKeyEx(INI.Key, CaminhoChave, 0, KEY_ALL_ACCESS, hKey)

If (rc <> ERROR_SUCCESS) Then
   VeRegistro = False
Else
   tmpVal = String$(1024, 0)
   KeyValSize = 1024
   rc = RegQueryValueEx(hKey, "CC", 0, KeyValType, tmpVal, KeyValSize)
   If (rc <> ERROR_SUCCESS) Then
      VeRegistro = False
   Else
      sSenha = Left(tmpVal, KeyValSize - 1)
      tmpVal = String$(1024, 0)
      KeyValSize = 1024
      rc = RegQueryValueEx(hKey, "SS", 0, KeyValType, tmpVal, KeyValSize)
      iSenha = Val(tmpVal)
      Decomposto = CInt(Decompoe(sSenha) / 75)
      If Abs(Decomposto - iSenha) < 2 Then
         VeRegistro = True
         
        tmpVal = String$(1024, 0)
        KeyValSize = 1024
        rc = RegQueryValueEx(hKey, "Cliente", 0, KeyValType, tmpVal, KeyValSize)
        lcCliente = tmpVal
         
      Else
         VeRegistro = False
      End If
   End If
End If
rc = RegCloseKey(hKey)
End Function

Public Function Cliente() As String
Cliente = lcCliente
End Function

Private Function EhPar(Num As Integer) As Boolean
If ((Num / 2) - Int(Num / 2)) = 0.5 Then
   EhPar = False
Else
   EhPar = True
End If
End Function

Private Function Decompoe(dado As String) As Long
Dim C   As Integer
Dim Aux As String

Aux = ""
For C = 1 To 11
   If EhPar(C) Then
      Aux = Aux + Mid(dado, C, 1)
   End If
Next
Decompoe = Val(Aux)
End Function

Public Sub Gera()
Dim C            As Integer
Dim cMoita       As Integer
Dim cSenha       As Integer
Dim iSemente     As Integer
Dim iContraSenha As Integer
Dim iSenha       As Long
Dim sChave       As String
Dim sSenha       As String
Dim sContraSenha As String
Dim sMoita       As String

Randomize Timer
iSemente = Int(Rnd * 440)
iSenha = iSemente * 75
iContraSenha = iSemente * 57
sSenha = Trim(str(iSenha))
sSenha = ComplStr(sSenha, 5, "0", 0)
sSenha = Mistura(sSenha)
sSenha = InclueHD(sSenha)
sContraSenha = Trim(str(iContraSenha))
sContraSenha = ComplStr(sContraSenha, 5, "0", 0)
sContraSenha = Mistura(sContraSenha)
Senha = sSenha
ContraSenha = sContraSenha
End Sub

Private Function Mistura(Num As String) As String
Dim Tam       As Integer
Dim TaMoita   As Integer
Dim cMoita    As Long
Dim cNum      As Integer
Dim C         As Integer
Dim sChave    As String
Dim sMoita    As String
Dim ResulTemp As String
Dim Ret       As String
Dim TaBom     As Boolean
Dim Saco      As Long

Screen.MousePointer = vbHourglass
TaMoita = Len(Num)
Tam = Len(Num) + TaMoita
Saco = 1
TaBom = False
Do While TaBom = False
   cNum = 1
   cMoita = 1
   Ret = ""
   sMoita = Trim(str(Int(Rnd * (10 ^ TaMoita))))
   sMoita = ComplStr(sMoita, TaMoita, "0", 0)
   For C = 1 To 11
      If EhPar(C) = False Then
         Ret = Ret + Mid(sMoita, cMoita, 1)
         cMoita = cMoita + 1
      Else
         Ret = Ret + Mid(Num, cNum, 1)
         cNum = cNum + 1
      End If
   Next
   ResulTemp = ComplStr(Ret, Tam, "0", 0)
   TaBom = True
   If InStr(ResulTemp, "00") > 0 Then
      TaBom = False
   Else
      If (Mid(ResulTemp, 1, 1) = "0") And (Mid(ResulTemp, 3, 1) = "0") And (Mid(ResulTemp, 5, 1) = "0") And (Mid(ResulTemp, 7, 1) = "0") And (Mid(ResulTemp, 9, 1) = "0") And (Mid(ResulTemp, 11, 1) = "0") Then TaBom = False
      If TaBom = False Then Saco = Saco + 1
      If Saco = 30000 Then TaBom = True
   End If
Loop
Mistura = ResulTemp
Screen.MousePointer = vbDefault
End Function

Public Function Confere(eSenha As String, eContraSenha As String)
Dim rSenha       As Single
Dim rContraSenha As Single

eSenha = RetiraHD(eSenha)
rSenha = Decompoe(eSenha) / 75
rContraSenha = Decompoe(eContraSenha) / 57
Confere = (rSenha = rContraSenha)
End Function

Public Sub Implanta(Nome As String)
Dim Ret
Dim iSenha As Integer
Dim rc     As Long
Dim sSenha As String

rc = RegCreateKey(Key, CaminhoChave, Ret)
rc = RegSetValueEx(Ret, "CC", 0, REG_SZ, ByVal Senha, Len(Senha))
Senha = RetiraHD(Senha)
iSenha = Int(Decompoe(Senha) / 75)
sSenha = Trim(str(iSenha))
rc = RegSetValueEx(Ret, "SS", 0, REG_SZ, ByVal sSenha, Len(sSenha))
rc = RegSetValueEx(Ret, "Cliente", 0, REG_SZ, ByVal Nome, Len(Nome))
RegCloseKey Ret
End Sub

Public Function Autentica(sSenha As String) As String
Dim sContraSenha As String
Dim iSemente     As Integer
Dim cSemente     As Currency
Dim iContraSenha As Currency

sSenha = RetiraHD(sSenha)
cSemente = Decompoe(sSenha)
iSemente = CInt(cSemente / 75)
iContraSenha = iSemente * 57
sContraSenha = Trim(str(iContraSenha))
sContraSenha = ComplStr(sContraSenha, 5, "0", 0)
sContraSenha = Mistura(sContraSenha)
Autentica = sContraSenha
End Function

Public Function InclueHD(sSenha As String) As String
Dim Volu   As Currency
Dim Num    As String * 2
Dim Um     As Integer

If InStr(Command, "REG") = 0 Then
    Volu = GetSerialNumber("C:\")
Else
   Volu = Val(Mid(Command, 4))
End If

Volu = Volu - Int(Volu / 100) * 100
Num = Trim(str(Volu))
Mid(sSenha, 7, 1) = Left(Num, 1)
Mid(sSenha, 9, 1) = Right(Num, 1)
Um = Val(Mid(sSenha, 8, 1))
Um = Um + Val(Left(Num, 1))
Um = Um Mod 10
Mid(sSenha, 8, 1) = Trim(str(Um))
Um = Val(Mid(sSenha, 10, 1))
Um = Um + Val(Right(Num, 2))
Um = Um Mod 10
Mid(sSenha, 10, 2) = Trim(str(Um))
InclueHD = sSenha
End Function

Private Function RetiraHD(eSenha As String) As String
Dim Um As String
Dim Num As Integer

Um = Mid(eSenha, 7, 1)
Num = Val(Mid(eSenha, 8, 1))
Num = Num - Um
If Num < 0 Then Num = Num + 10
Mid(eSenha, 8, 1) = Trim(str(Num))

Um = Mid(eSenha, 9, 1)
Num = Val(Mid(eSenha, 10, 1))
Num = Num - Um
If Num < 0 Then Num = Num + 10
Mid(eSenha, 10, 1) = Trim(str(Num))
RetiraHD = eSenha
End Function

