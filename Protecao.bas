Attribute VB_Name = "modProtecao"
'2.0.3 Ajuste no funcionamento das permissões

Option Explicit

'2.0.3 Ajuste no funcionamento das permissões
Public Permissao As Boolean   'Controle de permissao
Public Protecao  As New cProtecao

'2.0.5 Implantação do Log
Public VoluOrig  As Currency

Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public Function PoeNome() As String
Dim CLI   As String
Dim Compl As String
Dim Aux   As String
Dim C     As String * 1

CLI = VeVolume()

Select Case CLI

   Case "BON"
      Compl = "Registrado para Boni"
   Case "TRO"
      Compl = "Registrado para Trojan"
   Case Else
        Compl = "Demonstração"
End Select

'2.0.3 Ajuste no funcionamento das permissões
Permissao = (CLI > "")

If Permissao Then
   Loga "Usuário Registrado"
Else
   Permissao = Protecao.VeRegistro
   If Permissao Then
      Compl = "Registrado para " & Protecao.Cliente
      Loga Compl
   Else
      '2.0.5 Implantação do Log
      Loga "Usuário Não Registrado " & VoluOrig
   End If
End If

If Command$ = "NaoReg" Then
    Permissao = False
    Compl = "Não registrado [para testes]"
End If

PoeNome = Compl
End Function

Function VeVolume() As String
Dim Buf As Integer
Dim Volu As Currency
Dim Aux As String

VoluOrig = GetSerialNumber("C:\")

VoluOrig = 1351610406

If VoluOrig > 10 Then
    Volu = Int(VoluOrig / 10)
Else
    Volu = VoluOrig
End If
Select Case Volu

    Case 135161040 'PC que tem a impressora
        Aux = "BON" 'Boni AutoPeças
    Case 189249639
        Aux = "BON" 'Servidor
    Case Else
        Aux = ""
End Select
If Aux > "" Then
   VeVolume = Aux
End If
End Function

Public Function GetSerialNumber(strDrive As String) As Long
Dim SerialNum As Long
Dim Res As Long
Dim Temp1 As String
Dim Temp2 As String

Temp1 = String$(255, Chr$(0))
Temp2 = String$(255, Chr$(0))
Res = GetVolumeInformation(strDrive, Temp1, _
Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
GetSerialNumber = SerialNum
End Function

