VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmRelPagamentos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Pagamentos"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5340
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btConsulta 
      Caption         =   "Consulta"
      Height          =   375
      Left            =   2700
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      Height          =   615
      Left            =   60
      TabIndex        =   7
      Top             =   420
      Width           =   5235
      Begin VB.CheckBox Check 
         Caption         =   "Avulso"
         Height          =   375
         Index           =   4
         Left            =   4380
         TabIndex        =   14
         Top             =   180
         Width           =   795
      End
      Begin VB.CheckBox Check 
         Caption         =   "Pagamento"
         Height          =   375
         Index           =   3
         Left            =   3240
         TabIndex        =   11
         Top             =   180
         Width           =   1155
      End
      Begin VB.CheckBox Check 
         Caption         =   "Vale Transporte"
         Height          =   375
         Index           =   2
         Left            =   1800
         TabIndex        =   10
         Top             =   180
         Width           =   1455
      End
      Begin VB.CheckBox Check 
         Caption         =   "Comissão"
         Height          =   375
         Index           =   1
         Left            =   780
         TabIndex        =   9
         Top             =   180
         Width           =   975
      End
      Begin VB.CheckBox Check 
         Caption         =   "Vale"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   180
         Width           =   675
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Relatório"
      Height          =   375
      Left            =   1395
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txDtFIM 
      Height          =   285
      Left            =   3360
      MaxLength       =   20
      TabIndex        =   2
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txDtINI 
      Height          =   285
      Left            =   1980
      MaxLength       =   20
      TabIndex        =   1
      Top             =   60
      Width           =   975
   End
   Begin VB.ComboBox cbVend 
      Height          =   315
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1140
      Width           =   2355
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2865
      Left            =   60
      TabIndex        =   13
      ToolTipText     =   "Clique duplo chama a impressão"
      Top             =   1980
      Visible         =   0   'False
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   5054
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "até"
      Height          =   195
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Apartir de"
      Height          =   195
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Vendedores"
      Height          =   195
      Left            =   540
      TabIndex        =   3
      Top             =   1200
      Width           =   1020
   End
End
Attribute VB_Name = "frmRelPagamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'4.3.1 Inclusão do recibo avulso nos relatórios de recibos
'4.2.1 Impressão da observação
'4.0.5 ReImpressão individual dos vales
'3.8.3 Relatório de pagamentos

Option Explicit

'4.0.5 ReImpressão individual dos vales
Private Expandido As Boolean
Private Cont      As Integer
Private Tpo       As String

'4.3.1 Inclusão do recibo avulso nos relatórios de recibos
Private sTipo(4)  As String
'Private sTipo(3)  As String

Private rsVale    As Recordset
Private cRecibo   As clsRecibo

Private Sub btConsulta_Click()
'4.0.5 ReImpressão individual dos vales
Dim Linha As Integer

Consulta
If Expandido = False Then
    Me.Height = Me.Height * 2.125
    MSFlexGrid1.ColWidth(0) = 0
    MSFlexGrid1.ColWidth(1) = 600
    MSFlexGrid1.ColWidth(2) = 1950
    MSFlexGrid1.ColWidth(3) = 850
    MSFlexGrid1.ColWidth(4) = 800
    Expandido = True
    MSFlexGrid1.Visible = True
End If
MSFlexGrid1.Clear
MSFlexGrid1.TextMatrix(0, 1) = "Tipo"
MSFlexGrid1.TextMatrix(0, 2) = "Nome"
MSFlexGrid1.TextMatrix(0, 3) = "Data"
MSFlexGrid1.TextMatrix(0, 4) = "Valor"
rsVale.MoveLast
MSFlexGrid1.Rows = rsVale.RecordCount + 1
rsVale.MoveFirst
Do While rsVale.EOF = False
    Linha = Linha + 1
    MSFlexGrid1.TextMatrix(Linha, 0) = rsVale!ID
    MSFlexGrid1.TextMatrix(Linha, 1) = sTipo(rsVale!Tipo)
    MSFlexGrid1.TextMatrix(Linha, 2) = rsVale!Nome
    MSFlexGrid1.TextMatrix(Linha, 3) = Format(rsVale!Data, "dd/MM/YY") & " "
    MSFlexGrid1.TextMatrix(Linha, 4) = Format(rsVale!Valor, "##,###,###,##0.00")
    rsVale.MoveNext
Loop
End Sub

'4.0.5 ReImpressão individual dos vales
Private Sub btConsulta_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub

Private Sub Command1_Click()
If IsDate(txDtINI.Text) = False Or IsDate(txDtFIM.Text) = False Then
    msgboxL "Data Inválida"
Else

    '4.3.1 Inclusão do recibo avulso nos relatórios de recibos
    If (Check(0).Value + Check(1).Value + Check(2).Value + Check(3).Value + Check(4).Value) = 0 Then
    'If Check(0).Value = 0 And Check(1).Value = 0 And Check(2).Value = 0 Then
    
        msgboxL "Escolha pelo menos um tipo"
    Else
        'fazer o relatório
        GravaOpgRelPag
        Impressao
        Unload Me
    End If
End If
End Sub

Private Sub Form_Load()
Dim TbMec As Recordset

txDtFIM.Text = Format(Now, "DD/MM/YYYY")
txDtINI.Text = "01" + Right(txDtFIM.Text, 8)
AbreTB TbMec, "Select Nome From Mecanicos Where Ativo = True and Nome > '' Order by Nome ", dbOpenDynaset
cbVend.AddItem "Todos"
Do While TbMec.EOF = False
    cbVend.AddItem TbMec.Fields("Nome")
    TbMec.MoveNext
Loop
TbMec.Close
cbVend.ListIndex = 0
MostraOpgRelPag

'4.0.5 ReImpressão individual dos vales
Expandido = False
sTipo(0) = "Adiantamento"
sTipo(1) = "Comissao"
sTipo(2) = "Vale Transp"
sTipo(3) = "Pagamento"

'4.3.1 Inclusão do recibo avulso nos relatórios de recibos
sTipo(4) = "Pagamento"

Set cRecibo = New clsRecibo
End Sub

Private Sub MostraOpgRelPag()
Dim Nr  As Integer
Dim Ind As Integer
Dim Mul As Integer

Nr = INI.OpcRelPag

'4.3.1 Inclusão do recibo avulso nos relatórios de recibos
For Ind = 4 To 0 Step -1
'For Ind = 3 To 0 Step -1

    Mul = (2 ^ Ind)
    If (Nr >= Mul) Then
        Check(Ind).Value = 1
        Nr = Nr - Mul
    End If
Next
End Sub

Private Sub GravaOpgRelPag()
Dim Nr  As Integer
Dim Ind As Integer

'4.3.1 Inclusão do recibo avulso nos relatórios de recibos
For Ind = 4 To 0 Step -1
'For Ind = 3 To 0 Step -1

    If Check(Ind).Value Then
        Nr = Nr + (2 ^ Ind)
    End If
Next
INI.OpcRelPag = Nr
End Sub

Private Sub Impressao()
Dim ObsGra   As Boolean
Dim SpacoTot As Integer
Dim sValor   As Single
Dim E        As String
Dim Total    As Currency
Dim Aux      As String

'4.0.5 ReImpressão individual dos vales
'Dim rsVale  As Recordset
'Dim SQL     As String
'Dim SelTp   As String
'Dim a       As Integer
'Dim Cont    As Integer
'Dim Tpo      As String
'Dim sTipo(3) As String

'4.2.1 Impressão da observação
'Const TamFita = 55

ImprBuferizada_Inicializa

LPRINT "RELATÓRIO DE PAGAMENTOS "

'4.0.5 ReImpressão individual dos vales
'Aux = Aux & "Tipo"
'For a = 0 To 3
'    If Check(a).Value Then
'        Cont = Cont + 1
'        Tpo = Tpo & UCase(Check(a).Caption) & " - "
'        SelTp = SelTp & Trim(STR(a)) & ","
'    End If
'Next
'SelTp = Left(SelTp, Len(SelTp) - 1)
'SQL = "SELECT Data, Valor, Pago, Tipo, obs, Nome"
'SQL = SQL & " from Vales"
'SQL = SQL & " INNER JOIN Mecanicos ON Vales.IdOperador = Mecanicos.codi"
'SQL = SQL & " Where Data Between " & DTSqls(txDtINI.Text) & " And " & DTSqls(txDtFIM.Text, True)
'SQL = SQL & " and Tipo in (" & SelTp & ")"
'SQL = SQL & " order by Vales.ID "
'AbreTB rsVale, SQL, dbOpenDynaset
Consulta

If Cont > 1 Then
    Aux = Aux & "s:"
Else
    Aux = Aux & ":"
End If
Aux = Aux & Left(Tpo, Len(Tpo) - 3)
LPRINT Aux
LPRINT "De " & txDtINI.Text & " até " & txDtFIM.Text
If cbVend.ListIndex = 0 Then
    LPRINT "Todos "
Else
    LPRINT "Nome: " & cbVend.Text
End If
LPRINT String(TamFita, "-")

If cbVend.ListIndex = 0 Then
    Aux = "Nome                |"
    ObsGra = True
    SpacoTot = 39
Else
    Aux = ""
    ObsGra = False
    SpacoTot = 17
End If
Aux = Aux & "Tipo        |  Data  | Valor  |Obs"
LPRINT Aux

'4.3.1 Inclusão do recibo avulso nos relatórios de recibos
LPRINT String(TamFita, "-")

Do While rsVale.EOF = False
    If cbVend.ListIndex = 0 Then
    
        '4.3.1 Inclusão do recibo avulso nos relatórios de recibos
        If rsVale!Tipo = 4 Then
            Aux = ComplStr(rsVale!NomeAvulso, 22, " ", 0)
            
        Else
            Aux = ComplStr(rsVale!Nome, 22, " ", 0)
        End If
    Else
        Aux = ""
    End If
    Aux = Aux & ComplStr(sTipo(rsVale!Tipo), 13, " ", 0)
    Aux = Aux & Format(rsVale!Data, "dd/MM/YY") & " "
    Aux = Aux & ComplStr(Format(rsVale!Valor, "##,###,###,##0.00"), 9, " ", 2)
    If Trim(rsVale!Obs) > "" Then
        If ObsGra Then
            Aux = Aux & " *"
        Else
            Aux = Aux & " " & rsVale!Obs
        End If
    End If
    LPRINT Aux
    Total = Total + rsVale!Valor
    rsVale.MoveNext
Loop
LPRINT String(TamFita, "-")
If Total > 0 Then
    LPRINT Space(SpacoTot) & "Soma: " & Format(Total, "##,###,###,##0.00")
    sValor = Total
End If
If ImprBuferizada_Finaliza = False Then
    Exit Sub
End If
End Sub

'4.0.5 ReImpressão individual dos vales
Private Sub Consulta()
Dim a       As Integer
Dim SelTp   As String
Dim Aux     As String
Dim SQL     As String

Aux = Aux & "Tipo"

'4.3.1 Inclusão do recibo avulso nos relatórios de recibos
For a = 0 To 4
'For a = 0 To 3

    If Check(a).Value Then
        Cont = Cont + 1
        Tpo = Tpo & UCase(Check(a).Caption) & " - "
        SelTp = SelTp & Trim(STR(a)) & ","
    End If
Next
SelTp = Left(SelTp, Len(SelTp) - 1)

'4.3.1 Inclusão do recibo avulso nos relatórios de recibos
SQL = "SELECT ID, Data, Valor, Pago, Tipo, obs, Nome, NomeAvulso"
'SQL = "SELECT ID, Data, Valor, Pago, Tipo, obs, Nome"

SQL = SQL & " from Vales"
SQL = SQL & " INNER JOIN Mecanicos ON Vales.IdOperador = Mecanicos.codi"
SQL = SQL & " Where Data Between " & DTSqls(txDtINI.Text) & " And " & DTSqls(txDtFIM.Text, True)
SQL = SQL & " and Tipo in (" & SelTp & ")"
SQL = SQL & " order by Vales.ID "
AbreTB rsVale, SQL, dbOpenDynaset
End Sub

'4.0.5 ReImpressão individual dos vales
Private Sub MSFlexGrid1_DblClick()
If msgboxL("Deseja realmente realizar esta impressão?", vbYesNo + vbQuestion, "Impressão de pagamentos") = vbYes Then
    ImprPagIndiv
End If
End Sub

'4.0.5 ReImpressão individual dos vales
Private Sub ImprPagIndiv()
Dim ID         As Integer
Dim SQL        As String
Dim rsEsseVale As Recordset

ID = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0))

SQL = "SELECT Mecanicos.Nome, Vales.Tipo, Mecanicos.Ende, Vales.obs, Vales.Valor, Mecanicos.codi "
SQL = SQL & "FROM Mecanicos "
SQL = SQL & "INNER JOIN Vales ON Mecanicos.codi = Vales.IdOperador "
SQL = SQL & "WHERE Vales.ID = " & ID
AbreTB rsVale, SQL
gTipo = rsVale!Tipo
Select Case gTipo
    Case 0, 1 'Adiantamento, Comissão
        cRecibo.ReciboFita rsVale!Nome, rsVale!Ende, rsVale!Obs, rsVale!Valor, "", False, "", rsVale!codi, False
        MsgBox "Impressão realizada"
    Case 2 'Vale Transporte
        cRecibo.ReciboVT rsVale!Nome, rsVale!Ende, rsVale!Valor, ""
        MsgBox "Impressão realizada"
    Case 3  'Pagamento Mensal
        'cRecibo.ReciboPagamento rsVale!Mecanicos, rsVale!Ende, nrMec, "", Text1(1).Text, rsVale!Valor, Vale, rsVale!Obs, Text1(2).Text, Text1(2).Visible
        MsgBox "Impossível realizar esta impressão"
End Select
End Sub

'4.0.5 ReImpressão individual dos vales
Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub
