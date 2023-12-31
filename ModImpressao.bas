Attribute VB_Name = "ModImpressao"
'4.9.5 Definir o nr do tamanho da fonte na impressão USB
'4.9.4 Alteração da fonte da impressão USB de 10 para 12
'4.9.1 Ajuste da Impressão em Fita via Usb
'4.9.0 Previsão para impressora em Fita via Usb
'4.1.8 Log para local do acionamento da impressora
'3.6.6 tratamento para pc da impressora desligado
'2.7.2 Logar todas mensagens
'2.2.3 Ajuste no procedimento de forçar a digitação das mãos de obras
'2.0.6 Alteração do funcionamento interno das variáveis de configuração
'2.0.2 Retirar os acentos da impressão
'2.0.2 Não imprimir campos sem conteudo
'2.0.1 Impressão em fita

Option Explicit

Public Sub LPRINT(Aux As String, Optional Logar As Boolean = True)
Dim BufImp As Integer
Dim X      As Long
Dim L      As Integer

Static bVez As Boolean

'3.4.7 Impressão das tarefas
If Logar Then
    Loga Aux
End If

BufImp = FreeFile
Open CurDir$ + "\Impres.prn" For Append As BufImp

Imprime:
Print #BufImp, Aux
Debug.Print Aux
Close #BufImp

'2.0.6 Alteração do funcionamento interno das variáveis de configuração
If INI.UtTempIni = 1 And Not (Command$ = "R") And Len(Aux) > 0 Then
   Espera (INI.TempIni)
End If

Impresso:
L = L + 1
Aux = ""
X = DoEvents
Exit Sub

DeuGuruLP:
If Err = 52 Then
   Close #BufImp
   X = DoEvents
   Espera 0.001
   Open "LPT1" For Output As BufImp
   On Error GoTo DeuGuruLP
   Resume Imprime
   Err = 0
Else
   '2.7.2 Logar todas mensagens
   msgboxL Error$(Err) + " na linha " + STR$(Erl), vbCritical, "Erro"
   Resume Impresso
End If
End Sub

Public Sub ImprBuferizada_Inicializa()
Dim Arq As String
Dim X   As Long

LPRINT Chr$(15)
If Not (Command$ = "R") Then
   Arq = CurDir$ + "\Impres.prn"
   Loga "Arquivo da impressão: " & Arq
   If Dir$(Arq) > "" Then
      Do While Dir$(Arq) > ""
         Kill Arq
         X = DoEvents
      Loop
   End If
End If
End Sub

'4.9.1 Ajuste da Impressão em Fita via Usb
Private Sub ImprUsb(nmArq As String)
Dim Buf       As Integer
Dim Impressao As String

'4.9.5 Definir o nr do tamanho da fonte na impressão USB
Printer.Font.Size = INI.NrFonte
'4.9.4 Alteração da fonte da impressão USB de 10 para 12
'Printer.Font.Size = 12

Printer.Font.Name = "Courier New"
Buf = FreeFile()
Open nmArq For Input As #Buf
Do While (EOF(Buf) = False)
    Line Input #Buf, Impressao
    Printer.Print Impressao
Loop
Printer.EndDoc
Close #Buf
End Sub

Public Function ImprBuferizada_Finaliza() As Boolean
Dim a     As Integer
Dim Buf   As Integer
Dim Buf2  As Integer
Dim Texto As String
Dim Aux   As String
Dim nmArq As String
Dim X     As Long

If INI.LinhasApos > 0 Then
    For a = 0 To INI.LinhasApos
        LPRINT " ", False
    Next
End If

If Not (Command$ = "R") Then

    '4.9.0 Previsão para impressora em Fita via Usb
    nmArq = CurDir$ + "\Impres.prn"
    If INI.TpImpress = 2 Then
        Loga "Impressao na USB", lDBG
    
         '4.9.1 Ajuste da Impressão em Fita via Usb
         ImprUsb nmArq
    Else
        Loga "Impressao na Matricial", lDBG
        Buf = FreeFile
        Open nmArq For Input As Buf
        Do While EOF(Buf) = False
           Line Input #Buf, Aux
           Texto = Texto + Aux + vbCrLf
        Loop
        Close #Buf
        Buf2 = FreeFile
        '2.0.6 Alteração do funcionamento interno das variáveis de configuração
        
        '4.1.8 Log para local do acionamento da impressora
        Loga "Impressao acionada em: " & INI.Impressora
        
        '3.6.6 tratamento para pc da impressora desligado
TentaDenovo:
        On Local Error GoTo PathNotFound
        Open INI.Impressora For Output As #Buf2
        On Local Error GoTo ErroImpBufFin
        Print #Buf2, Texto
        On Local Error GoTo 0
        Close #Buf2
    End If
End If

ImprBuferizada_Finaliza = True
Exit Function

ErroImpBufFin:
'On Local Error Resume Next
'Close #Buf2
'On Local Error GoTo 0
'2.7.2 Logar todas mensagens
msgboxL "Erro ao imprimir em " & INI.Impressora
Exit Function

'3.6.6 tratamento para pc da impressora desligado
PathNotFound:
If MsgBox("Deseja tentar novamente", vbExclamation + vbYesNo, "Computador que ta a impressora esta desligado") = vbYes Then
    GoTo TentaDenovo
Else
    Exit Function
End If
End Function
Public Function CENTRAL(dado As String, Letras As Integer) As String
Dim a    As Integer
Dim Aux  As String
Dim meio As Single

If Len(dado) / 2 < Letras Then
   For a = Len(dado) To 1 Step -1
      If Mid$(dado, a, 1) > " " Then GoTo Tira
   Next
Tira:
   Aux = Left$(dado, a)
   meio = Letras - Int(Len(Aux) / 2)
   CENTRAL = Space$(meio) + Aux + Space$(meio)
Else
   CENTRAL = dado
End If
End Function

Private Sub Espera(Tempo As Single)
Dim Depois As Double
Dim X      As Integer

If Tempo = 0 Then Tempo = 0.1
Depois = Now + ((Tempo / 1000) / 86400)
Do While Now < Depois
   X = DoEvents()
Loop
End Sub

Public Sub LPRINTST(Titulo As String, Conteudo As String, Optional Complemento As String)
'2.2.5 Não imprimir campos sem conteudo [2]
'2.2.3 Ajuste no procedimento de forçar a digitação das mãos de obras
'2.0.2 Não imprimir campos sem conteudo
If Conteudo > "" And Conteudo <> "0" And Trim(Conteudo) <> "0,00" Then
    If IsMissing(Complemento) Then
        LPRINT Titulo & " " & Conteudo
    Else
        LPRINT Titulo & " " & Conteudo & " " & Complemento
    End If
End If
End Sub

Public Function SemAcento(Palavra As String) As String
'2.0.2 Retirar os acentos da impressão
Dim a%
Dim letra%

For a% = 1 To Len(Palavra)
    letra% = TiraAcentos(Asc(Mid$(Palavra, a%, 1)))
    Mid$(Palavra, a%, 1) = Chr$(letra%)
Next
SemAcento = Palavra
End Function

Private Function TiraAcentos(CodLetra As Integer) As Integer
'2.0.2 Retirar os acentos da impressão
Dim lsLetra As String * 1

CodLetra = Asc(Chr$(CodLetra))
Select Case CodLetra
   Case 193, 192, 194, 196, 195
      CodLetra = 65
   Case 199
      CodLetra = 67
   Case 200, 202, 203, 230
      CodLetra = 69
   Case 204, 205, 206, 207
      CodLetra = 73
   Case 211, 210, 212, 214, 213
      CodLetra = 79
   Case 217, 218, 219, 220
      CodLetra = 85
   Case 224, 225, 226, 227, 228
      CodLetra = 97
   Case 231
      CodLetra = 99
   Case 232, 234, 235
      CodLetra = 101
   Case 236, 237, 238, 239
      CodLetra = 105
   Case 243, 242, 244, 245, 246
      CodLetra = 111
   Case 249, 250, 251, 252
      CodLetra = 117
   Case Else
      If CodLetra > 127 Then
         CodLetra = 0
      End If
End Select
TiraAcentos = CodLetra
End Function
