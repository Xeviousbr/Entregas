VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Classe para centralizar as funções de envio de email

'Alternativametne o objeto MAPI poderia ser criado em tempo de execussão já na classe
'assim não precisaria ser recebido por parametro

Private MAPIm      As MAPIMessages
Private MAPIs      As MAPISession

Private l_Cliente      As String
Private l_enderMailCli As String

Public Sub EnviaEmail()
Screen.MousePointer = vbHourglass

On Error GoTo SemLogin
MAPIs.Action = 1 'Logon
MAPIm.SessionID = MAPIs.SessionID
MAPIm.Action = 6 'Compor novo email
MAPIm.MsgSubject = "Orçamento " & Config.Nome & "Orçamento nr. " & Config.NrPed
MAPIm.RecipAddress = l_enderMailCli
MAPIm.MsgNoteText = ComposicaoMail

On Local Error Resume Next
Screen.MousePointer = vbDefault
MAPIm.Action = 2 'vbMessageSendDlg
On Local Error GoTo 0
Exit Sub

SemLogin:
Screen.MousePointer = vbDefault
'MsgBox Error(Err)
If MsgBox("É necessário configurar o OutLook para utilizar esse recurso", vbExclamation + vbYesNo + vbDefaultButton1, "Deseja fazer isso agora?") = vbYes Then
    Shell "C:\Arquivos de programas\Outlook Express\msimn.exe"
End If
End Sub

Private Function ComposicaoMail() As String
Dim PosRecibo  As Integer
Dim Buf        As Integer
Dim AuxJaTinta As String * 1
Dim Texto      As String
Dim TbDecr     As Recordset

'Descrição
SQL = "Select DescricaoREC, Config_Telef2, Config_Ender2 From Config "
'Set TbDecr = SQLog(SQL)
Texto = TbDecr!DescricaoREC & vbNewLine

PosRecibo = InStr(Texto, "RECIBO")
If PosRecibo > 0 Then
    Texto = RTrim(Left(Texto, PosRecibo - 1))
    Texto = Replace(Texto, "|", " ")
End If

'Endereço
'Texto = Texto & Trim(Config.Ender) & vbNewLine

'Orçamento
Texto = Texto & "Orçamento Nr: " & Config.NrPed & vbNewLine

'Nome do cliente
Texto = Texto & "Cliente: " & l_Cliente & vbNewLine & vbNewLine

'Endereço do Cliente
'Texto = Texto & "Endereço: " & Ender.Text & vbNewLine

If FileExists(CurDir$ + "\Impres.prn") Then
    Kill CurDir$ + "\Impres.prn"
End If

AuxJaTinta = Config.JaTinta
Config.JaTinta = "0"
'OrcMail
Config.JaTinta = AuxJaTinta

Buf = FreeFile
Open CurDir$ + "\Impres.prn" For Input As Buf
Do While EOF(Buf) = False
   Line Input #Buf, Aux
   Texto = Texto & Aux & vbNewLine
Loop
Close #Buf
TbDecr.Close

Debug.Print Texto

ComposicaoMail = Texto
End Function

Public Property Let objMAPIMessage(ByVal vNewValue As MAPIMessages)
Set MAPIm = vNewValue
End Property

Public Property Let objMAPISession(ByVal vNewValue As MAPISession)
Set MAPIs = vNewValue
End Property

Public Property Let nmCliente(ByVal vNewValue As String)
l_Cliente = vNewValue
End Property

Public Property Let enderMailCli(ByVal vNewValue As String)
l_enderMailCli = vNewValue
End Property
