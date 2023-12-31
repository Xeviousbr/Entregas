VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmValores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valores"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6000
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btFechar 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   255
      Left            =   7080
      TabIndex        =   7
      Top             =   4020
      Width           =   375
   End
   Begin VB.CommandButton btImprimir 
      Caption         =   "Imprimir"
      Height          =   315
      Left            =   4560
      TabIndex        =   6
      Top             =   4020
      Width           =   1215
   End
   Begin VB.CommandButton btExcluir 
      Caption         =   "Excluir"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3060
      TabIndex        =   5
      Top             =   4020
      Width           =   1215
   End
   Begin VB.CommandButton btPagar 
      Caption         =   "Pagar"
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   4020
      Width           =   1215
   End
   Begin VB.CommandButton btEditar 
      Caption         =   "Editar"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   4020
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3585
      Left            =   60
      TabIndex        =   2
      Top             =   360
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   6324
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Não há movimentação financeira com este cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   5355
   End
   Begin VB.Label lbNome 
      Alignment       =   2  'Center
      Caption         =   "..."
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5715
   End
End
Attribute VB_Name = "frmValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'4.2.1 Impressão da observação
'3.9.5 Exclusão de pagamentos
'3.9.1 Pagamento com parcelas

Option Explicit

Private Sub btEditar_Click()
If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = "SIM" Then
    EditParc.DtPago = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
End If
EditParc.Valor = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
EditParc.Parcela = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)
EditParc.Show 1
Mostra
End Sub

'3.9.5 Exclusão de pagamentos
Private Sub btExcluir_Click()
If MsgBox("Confirma a exclusão deste pagamento", vbYesNo + vbDefaultButton2, "Exclusão de pagamentos") = vbYes Then

    Dim Valor As Currency
    Valor = Consulta("Select Valor From Parcelas Where idParc = " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6))

    ExecSql "Delete From Parcelas Where idParc = " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)
            
    If Consulta("Select Count(1) From Parcelas WHERE Parcelas.Cli = " & clsCLi.NrCli) = 0 Then
        'Se não resta mais parcelas deve ser retirado a informação de pagamento
        ExecSql "Update Orcamento Set Pagamento = null, VlrPago = 0 Where Orcamento = " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)
    Else
        If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = "SIM" Then
            'Caso contrário, caso a parcela tinha sido paga deve abater o valor
            Dim SQL As String
            SQL = "Update Orcamento Set VlrPago = VlrPago - " & VlrSql(STR(Valor))
            SQL = SQL & " Where Orcamento = " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7)
            ExecSql SQL
        End If
    End If
    Mostra
End If
End Sub

Private Sub btFechar_Click()
Unload Me
End Sub

Private Sub btImprimir_Click()
Dim a    As Integer
Dim AuxS As String

'4.2.1 Impressão da observação
'Const TamFita = 55

ImprBuferizada_Inicializa
LPRINT Chr$(15)
LPRINT CENTRAL("RELATORIO DE PARCELAS", TamFita / 2)
LPRINT CENTRAL(GCliente, TamFita / 2)
LPRINT "Data        Valor    Pagamento"
For a = 1 To (MSFlexGrid1.Rows - 1)
    AuxS = MSFlexGrid1.TextMatrix(a, 1) & " "
    AuxS = AuxS & ComplStr(MSFlexGrid1.TextMatrix(a, 2), 8, " ", 2)
    If MSFlexGrid1.TextMatrix(a, 3) = "SIM" Then
        AuxS = AuxS & " " & MSFlexGrid1.TextMatrix(a, 8)
    End If
    LPRINT AuxS
Next
LPRINT " "
LPRINT "Porto Alegre, " & Format$(Now, "dd") & " de " & Format$(Now, "mmmm") + " de " + Format$(Now, "yyyy")
ImprBuferizada_Finaliza
MsgBox "Impressão realizada", vbInformation, "OrCarro"
End Sub

Private Sub btPagar_Click()
Dim Resp       As String
Dim SQL        As String
Dim rsMec      As Recordset

Load frmSenha
frmSenha.Tipo = 1
frmSenha.Show 1
If frmSenha.Resultado = True Then
    Resp = frmSenha.Senha
End If
Unload frmSenha
If Resp = "" Then
    MsgBox "Operador não identificado"
Else
    Set rsMec = BuscaMec(Resp, " and Oper > 0 and Recebe = 1 ")
    If rsMec.EOF Then
        MsgBox "Operador não identificado"
    Else
        SQL = "Update Parcelas Set Pagto = " & DTSqld(Now)
        SQL = SQL & ", BalcRec = " & rsMec!codi
        SQL = SQL & " Where idParc = " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)
        ExecSql SQL
        MSFlexGrid1.Clear
        Mostra
    End If
End If
End Sub

Private Sub Mostra()
Dim a         As Integer
Dim SQL       As String
Dim rsValores As Recordset

SQL = "SELECT Parcelas.Cli, Parcelas.NrParc, Parcelas.Data, Parcelas.Valor, Parcelas.Pagto, Parcelas.BalcFez, Mec.Nome AS Fez, Rece.Nome, Parcelas.idParc "

'3.9.5 Exclusão de pagamentos
SQL = SQL & ",Parcelas.Orc, Parcelas.Pagto "

SQL = SQL & "FROM Mecanicos Rece RIGHT JOIN (Mecanicos AS Mec INNER JOIN Parcelas ON Mec.codi = Parcelas.BalcFez) ON Rece.codi = Parcelas.BalcRec "
SQL = SQL & "WHERE Parcelas.Cli = " & clsCLi.NrCli

AbreTB rsValores, SQL, dbOpenDynaset
On Local Error GoTo NumTem
rsValores.MoveLast
On Local Error GoTo 0
rsValores.MoveFirst
MSFlexGrid1.Rows = rsValores.RecordCount + 1
a = 1
Do While rsValores.EOF = False
    MSFlexGrid1.TextMatrix(a, 0) = rsValores!NrParc
    MSFlexGrid1.TextMatrix(a, 1) = Format(rsValores!Data, "dd/mm/yyyy")
    MSFlexGrid1.TextMatrix(a, 2) = Format(rsValores!Valor, "###,###.00")
    MSFlexGrid1.TextMatrix(a, 4) = rsValores!Fez
    If rsValores!Pagto > 0 Then
        MSFlexGrid1.TextMatrix(a, 3) = "SIM"
        MSFlexGrid1.TextMatrix(a, 5) = rsValores!Nome
    Else
        MSFlexGrid1.TextMatrix(a, 3) = " "
    End If
    If a = 1 Then
        If MSFlexGrid1.TextMatrix(a, 3) = "SIM" Then
            MSFlexGrid1.TextMatrix(a, 8) = rsValores!Pagto
            btPagar.Enabled = False
        End If
    End If
    MSFlexGrid1.TextMatrix(a, 6) = rsValores!idParc
    
    '3.9.5 Exclusão de pagamentos
    MSFlexGrid1.TextMatrix(a, 7) = rsValores!Orc
        
    a = a + 1
    rsValores.MoveNext
Loop

'3.9.5 Exclusão de pagamentos
If INI.ModoOperacao = tpEscritorio Then
    btExcluir.Enabled = True
    btEditar.Enabled = True
End If

Sai_LoadPagto:
MSFlexGrid1.TextMatrix(0, 0) = "Parcela"
MSFlexGrid1.TextMatrix(0, 1) = "Data"
MSFlexGrid1.TextMatrix(0, 2) = "Valor"
MSFlexGrid1.TextMatrix(0, 3) = "Pago"
MSFlexGrid1.TextMatrix(0, 4) = "Atendeu"
MSFlexGrid1.TextMatrix(0, 5) = "Recebeu"
MSFlexGrid1.ColAlignment(3) = vbCenter
Exit Sub

NumTem:
MSFlexGrid1.Visible = False
Resume Sai_LoadPagto
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
InicForm Me
lbNome.Caption = GCliente
Mostra
End Sub

Private Sub MSFlexGrid1_Click()
btPagar.Enabled = Not (MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = "SIM")
If INI.ModoOperacao = tpEscritorio Then
    btEditar.Enabled = True
End If
End Sub

