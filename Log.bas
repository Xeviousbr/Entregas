Attribute VB_Name = "Log"
'3.0.2 Log em Rede
'2.9.9 Gravar o Log já na pasta Log
'2.7.2 Logar todas mensagens
'2.0.5 Implantação do Log

Public Logar            As Boolean
Private AplicDirLog     As String
Private ArqLog          As String
Private LogInicializado As Boolean
Public Enum tLog
    lCMD = 0    'Comandos do usuário
    lSQL = 1    'Comandos de consulta SQL
    lUPD = 2    'Comandos de alteração na base de dados
    lDBG = 3    'Informações logadas
    lERR = 4    'Erros
    lMSG = 5    'Mensagens
End Enum

'2.0.5 Implantação do Log
Public Sub PreparaOLog()
Dim DtArq  As Date
Dim DT     As String
Dim NM     As String
Dim NmLog  As String
Dim Copiar As Boolean

'3.0.2 Log em Rede
Dim LogarEmRede As Boolean
Dim snrPC       As String

Logar = True
NmLog = App.EXEName

'3.0.2 Log em Rede
If INI.LogEmRede Then
    If InStr(App.Path, "\\") > 0 Then
        LogarEmRede = True
    End If
End If
If LogarEmRede Then
    AplicDirLog = App.Path & "\Log\"
    snrPC = "-" & Trim(STR(INI.PC))
Else
    '2.9.9 Gravar o Log já na pasta Log
    AplicDirLog = AplicDirNat & "\Log\"
End If
ArqLog = AplicDirLog & App.EXEName & snrPC & ".log"

On Local Error Resume Next
MkDir AplicDirLog
On Local Error GoTo 0

If FileExists(ArqLog) Then
    DtArq = FileDateTime(ArqLog)
    If Int(DtArq) < Int(Now) Then
       Copiar = True
    Else
       If FileLen(ArqLog) > 1000000 Then
          Copiar = True
       End If
    End If
    
    If Copiar Then
        Atraso 0.2
        DT = Format(Now, "dd/mm/yyyy")
        DT = Left(DT, 2) + Mid(DT, 4, 2) + Right(DT, 4)
'        On Local Error Resume Next
'        MkDir AplicDirNat + "\LOG"
'        On Local Error GoTo 0

        NM = NmLog & snrPC & IIf(snrPC > "", "-", "") & DT & ".Log"
        'NM = NmLog & DT & ".Log"
        '
        On Local Error GoTo FazODirLog
        
        '3.0.2 Log em Rede
        FileCopy ArqLog, AplicDirLog & NM
        'FileCopy AplicDirLog & NmLog & ".Log", AplicDirLog & NM
        
        On Local Error GoTo 0
        
        '3.0.2 Log em Rede
        Kill ArqLog
        'Kill AplicDirLog & NmLog & ".Log"
        
'        MkDir AplicDirNat + "\LOG"
'        On Local Error GoTo 0
'        NM = NmLog + DT + ".Log"
'        On Local Error GoTo FazODirLog
'        FileCopy AplicDirNat + "\" + NmLog + ".Log", AplicDirNat + "\Log\" + NM
'        On Local Error GoTo 0
'        Kill AplicDirNat + "\" + NmLog + ".Log"
    End If
End If
Sai:
LogInicializado = True
Exit Sub

FazODirLog:
Select Case Err
    Case 52
        Kill App.EXEName + NmLog
        Resume Sai
    Case 53
        Resume Sai
    Case Else
        MkDir CurDir() + "\Log"
        Resume
End Select
End Sub

Public Sub Loga(ByVal Texto As String, Optional Operacao As tLog)
'2.0.5 Implantação do Log
Dim Buf    As Integer
Dim agora  As String
Dim sCLog  As String
Dim sOper  As String

Static ContLog As Long

Debug.Print Texto

'If Texto = "SELECT DISTINCTROW Orcamento.Data, Orcamento.Total, Orcamento.Kilom, Carros.Modelo, Orcamento.Cliente, Orcamento.Carro, Orcamento.Orcamento, Orcamento.Pagamento, Clientes.Ender, Carros.Placa, Clientes.NrCli ,Orcamento.Obs ,Clientes.Observacao, Clientes.NrCli ,Orcamento.ObsMec ,Carros.Historico ,Clientes.Funcionario from Orcamento, Clientes, Carros Where Clientes.NrCli =" Then
'    Stop
'End If

If LogInicializado Then
    If Logar Then
        Buf = FreeFile
        
        '2.9.9 Gravar o Log já na pasta Log
        'ArqLog = AplicDirLog & App.EXEName & ".log"
        'ArqLog = AplicDirNat + "\" + App.EXEName + ".log"
        
        Select Case Operacao
            Case lCMD:
                sOper = "CMD:"
            Case lSQL:
                sOper = "SQL:"
            Case lUPD:
                sOper = "UPD:"
            Case lDBG:
                sOper = "DBG:"
            Case lERR:
                sOper = "ERR:"
            Case lMSG:
                sOper = "MSG:"
        End Select
        
        On Local Error GoTo SemPath
        Open ArqLog For Append As #Buf
        On Local Error GoTo 0
        ContLog = ContLog + 1
        sCLog = "|" & Trim(STR(ContLog)) & "|"
        Texto = Format(Now, "hh:mm:ss") & sCLog & sOper & " " & Texto
        Print #Buf, Texto
        Close #Buf
    End If
End If
Exit Sub

SemPath:
Logar = False
End Sub

Public Function msgboxL(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = "") As VbMsgBoxResult
'2.7.2 Logar todas mensagens
Loga Title & IIf(Title > "", ":", "") & Prompt, lMSG
msgboxL = MsgBox(Prompt, Buttons, Title)
End Function

