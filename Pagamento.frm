VERSION 5.00
Object = "{00028C4A-0000-0000-0000-000000000046}#5.0#0"; "TDBG5.OCX"
Begin VB.Form FormaPagamento 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forma de pagamento"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3120
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   3120
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data DataP 
      Caption         =   "DataT"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Z:\Share\Orcarro\OrCarro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      EOFAction       =   2  'Add New
      Exclusive       =   0   'False
      Height          =   435
      Left            =   2220
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select IDPC, Data, Valor From ParcelasTemp"
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
   End
   Begin TrueDBGrid50.TDBGrid TDBGrid1 
      Bindings        =   "Pagamento.frx":0000
      Height          =   1815
      Left            =   60
      OleObjectBlob   =   "Pagamento.frx":0014
      TabIndex        =   7
      Top             =   1140
      Width           =   2955
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1980
      TabIndex        =   6
      Top             =   3420
      Width           =   1035
   End
   Begin VB.CommandButton btAdic 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   5
      Top             =   3420
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1860
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3000
      Width           =   1035
   End
   Begin VB.TextBox txEntrada 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1860
      TabIndex        =   2
      Top             =   360
      Width           =   1035
   End
   Begin VB.Label lbAFaltante 
      Caption         =   "Valor que falta a pagar: R$ "
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   780
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Label Label3 
      Caption         =   "Identificação: "
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3060
      Width           =   1395
   End
   Begin VB.Label Label2 
      Caption         =   "Pagamento a Vista:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   1395
   End
   Begin VB.Label lbTotal 
      Caption         =   "Valor Total:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2475
   End
End
Attribute VB_Name = "FormaPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private lcOK        As Boolean
Private lcOrcamento As Integer
Private Parcela     As Integer
Private lcValor     As Currency
Private Restante    As Currency
Private VlrEntrada  As Double
Private VlrParcelas As Double

'3.9.2 Gravar quem recebeu o pagamento das parcelas
Private lcQuemPagou As Integer

'3.9.3 Permitir pagamentos em pedaços
Private lcVlrPago As Currency

'3.9.2 Gravar quem recebeu o pagamento das parcelas
'Public Property Get Resultado() As Boolean
'Resultado = lcOK
'End Property

'3.9.2 Gravar quem recebeu o pagamento das parcelas
Public Property Get QuemPagou() As Integer
QuemPagou = lcQuemPagou
End Property

'3.9.2 Gravar quem recebeu o pagamento das parcelas
Public Property Let QuemPagou(QuemEh As Integer)
lcQuemPagou = QuemEh
End Property

Public Property Get Valor() As Currency
Valor = lcValor
End Property

Public Property Let Valor(ByVal vNewValue As Currency)
lcValor = vNewValue
Restante = lcValor
End Property

Private Sub InsereParcela(iParc As Integer, Data As Date, Valor As Double, Fez As Integer, Optional Recebeu As Integer = 0)
Dim SQL  As String
Dim Aux1 As String
Dim Aux2 As String

If iParc = 0 Then
    Aux1 = " ,BalcRec, Pagto"
    Aux2 = " ," & Recebeu & ", " & DTSqld(Now)
End If
Insere:
SQL = "Insert Into Parcelas (Orc, Cli, NrParc, Data, Valor, BalcFez" & Aux1 & ") Values ("
SQL = SQL & lcOrcamento & ", "
SQL = SQL & clsCLi.NrCli & ", "
SQL = SQL & iParc & ", "
SQL = SQL & DTSqld(Data) & ", "
SQL = SQL & VlrSql(STR(Valor)) & ", "
SQL = SQL & Fez
SQL = SQL & Aux2 & ")"
ExecSql SQL
End Sub

Private Sub btAdic_Click()
Dim iParc      As Integer

'3.9.2 Gravar quem recebeu o pagamento das parcelas
'Dim Balconista As Integer

Dim VlrParc    As Currency
Dim VlrEntrada As Double
Dim Resp       As String
Dim SQL        As String
Dim rs         As Recordset
Dim rsMec      As Recordset

If Text1.Text = "" Then
    MsgBox "Operador não identificado"
    Exit Sub
End If
Resp = Text1.Text
Set rsMec = BuscaMec(Resp, " and Oper > 0 ")
If rsMec.EOF Then
    MsgBox "Operador não identificado"
    Exit Sub
Else

    '3.9.2 Gravar quem recebeu o pagamento das parcelas
    QuemPagou = rsMec!codi
    'Balconista = rsMec!codi
    'lcOK = True
    
End If
DataP.Enabled = False
VlrParc = Consulta("Select Sum(Valor) From ParcelasTemp Where IDPC = " & INI.PC)
VeValor txEntrada.Text, VlrEntrada, txEntrada, 0

'3.9.4 Ajustar para não calcular errado o somatório das parcelas
Dim X As Long
X = DoEvents()

'3.9.3 Permitir pagamentos em pedaços
Dim VlrPagoTemp As Currency
VlrPagoTemp = VlrEntrada + VlrParc

'3.9.3 Permitir pagamentos em pedaços
If Valor < VlrPagoTemp Then
'If Valor <> (VlrEntrada + VlrParc) Then

    MsgBox "Valor inválido", vbExclamation, "Orcamento"
Else

    '3.9.2 Conserto do erro do pagamento a vista
    If VlrEntrada > 0 Then
    
        '3.9.2 Gravar quem recebeu o pagamento das parcelas
        InsereParcela 0, Int(Now), VlrEntrada, QuemPagou, QuemPagou
        'InsereParcela 0, Int(Now), VlrEntrada, Balconista
        
    End If
    If VlrParc > 0 Then
    
        SQL = "SELECT Data, Valor, Auto "
        SQL = SQL & "From ParcelasTemp "
        SQL = SQL & "WHERE IDPC = " & INI.PC
        SQL = SQL & " Order by Auto "
        AbreTB rs, SQL, dbOpenSnapshot
        rs.MoveFirst
        
        '3.9.2 Conserto do erro do pagamento a vista
    '    If VlrEntrada > 0 Then
    '        InsereParcela 0, Int(Now), VlrEntrada, Balconista
    '    End If
        
        If rs.EOF = False Then
            iParc = 1
            Do While rs.EOF = False
            
                '3.9.2 Gravar quem recebeu o pagamento das parcelas
                InsereParcela iParc, rs!Data, rs!Valor, QuemPagou
                'InsereParcela iParc, rs!Data, rs!Valor, Balconista
                
                iParc = iParc + 1
                rs.MoveNext
            Loop
        End If
    End If
    
    '4.0.0 Conserto do pagamento parcial
    VlrPago = VlrEntrada
    '3.9.3 Permitir pagamentos em pedaços
    'VlrPago = VlrPagoTemp
    
    Unload Me
End If
End Sub

Private Sub Command1_Click()
'3.9.4 Ao cancelar o pagamento não deve gravar
QuemPagou = 0
VlrPago = 0
Unload Me
End Sub

Private Sub DataP_Validate(Action As Integer, Save As Integer)
Static Quant As Integer
Static aData(12) As Date

Dim a         As Integer
Dim Col       As Integer
Dim Row       As Integer
Dim Data      As Date
Dim sData     As String
Dim sValor    As String
Dim MensErro  As String

If Action = 6 Then
    'Crítica da data
    TDBGrid1.Col = 0
    sData = TDBGrid1.Text
    
    ', Data
    If CriticaData(sData, Data) = 0 Then
    
        'Isso deveria funcionar
        Save = 0
        
        'Isso me parece irrelevante
        Exit Sub
    Else
        If Data < Now Then
            TDBGrid1.Text = ""
            MsgBox "Data Inválida"
            Exit Sub
        Else
            Row = TDBGrid1.Row
            If Row > Quant Then
                Quant = Row
            End If
            For a = Quant To 0 Step -1
                If Data < aData(a) Then
                    TDBGrid1.Text = ""
                    MsgBox "Data Anterior a outra já registrada"
                    Exit Sub
                End If
            Next
            aData(Row) = Data
            
        End If
    End If
    
    'Crítica do valor
    TDBGrid1.Col = 1
    sValor = TDBGrid1.Text
    If sValor = "" Then
        MensErro = "Valor inválido"
    Else
        Valor = Valo(dado:=sValor)
        If Valor <= 0 Then
            MensErro = "Valor inválido"
        Else
            If Valor > Restante Then
                MensErro = "Valor maior do que o valor a ser cobrado"
            End If
        End If
    End If
    If MensErro > "" Then
        TDBGrid1.Col = 0
        TDBGrid1.Text = ""
        TDBGrid1.Col = 1
        TDBGrid1.Text = ""
        MsgBox MensErro, vbExclamation, "OrCarro"
        Exit Sub
    End If
End If
End Sub

Private Sub Form_Load()
Dim ErrCria As Long

VlrEntrada = 0
VlrParcelas = 0

'3.9.2 Gravar quem recebeu o pagamento das parcelas
lcQuemPagou = 0
'lcOK = False

InicForm Me
If ExecSql("Delete From ParcelasTemp Where IDPC = " & INI.PC) = 3078 Then
    ErrCria = ExecSql("CREATE TABLE ParcelasTemp ( Auto autoincrement, IDPC LONG, Data DATETIME, Valor Money, PRIMARY KEY (Auto, IDPC) )")
    If ErrCria = 0 Then
        ErrCria = ExecSql("CREATE TABLE Parcelas ( idParc autoincrement, Orc LONG, Cli Integer, NrParc Integer, Data DATETIME, Valor Money, Pagto DATETIME, BalcFez Integer, BalcRec Integer, PRIMARY KEY (idParc, Orc, Cli, NrParc) )")
    End If
End If
                
If ErrCria > 0 Then
    MsgBox Error(ErrCria)
End If

TDBGrid1.Columns(0).Width = 0
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
TDBGrid1.Col = 0
TDBGrid1.Text = INI.PC
End Sub

Private Sub TDBGrid1_AfterInsert()
Dim SQL As String

SQL = "Update ParcelasTemp Set IDPC = " & INI.PC
SQL = SQL & " Where IDPC Is Null or IDPC = 0 "
ExecSql SQL

VlrParcelas = Consulta("Select Sum(Valor) From ParcelasTemp Where IDPC = " & INI.PC)
MostraVlrAPagar
End Sub

Private Sub TDBGrid1_KeyPress(KeyAscii As Integer)
'3.9.4 Substituir a virgula por ponto na digitação das parcelas
If KeyAscii = 46 Then
    KeyAscii = 44
End If
End Sub

Private Sub txEntrada_GotFocus()
Seleciona
End Sub

Private Sub txEntrada_KeyUp(KeyCode As Integer, Shift As Integer)
MostraVlrAPagar
End Sub

Private Sub MostraVlrAPagar()
VeValor txEntrada.Text, VlrEntrada, txEntrada, 0
Restante = lcValor - VlrEntrada - VlrParcelas
If Restante <= 0 Then
    Restante = 0
    lbAFaltante.Visible = False
Else
    lbAFaltante.Caption = "Valor que falta a pagar: R$ " & Format(Restante, "###,###.00")
    lbAFaltante.Visible = True
End If
End Sub

Public Property Get Orcamento() As Integer
Orcamento = lcOrcamento
End Property

Public Property Let Orcamento(ByVal vNewValue As Integer)
lcOrcamento = vNewValue
DataP.DatabaseName = Base
ContinuaDaCriacao:
DataP.RecordSource = "Select Auto, IDPC, Data, Valor From ParcelasTemp Where IDPC = " & INI.PC
DataP.Enabled = True
DataP.Refresh
End Property

'3.9.3 Permitir pagamentos em pedaços
Public Property Get VlrPago() As Currency
VlrPago = lcVlrPago
End Property

'3.9.3 Permitir pagamentos em pedaços
Public Property Let VlrPago(ByVal vNewValue As Currency)
lcVlrPago = vNewValue
End Property

