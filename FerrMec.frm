VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FerrMec 
   Caption         =   "Ferramentas do Mecânico: "
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4635
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btImpr 
      Caption         =   "Imprimir"
      Enabled         =   0   'False
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Top             =   1260
      Width           =   1035
   End
   Begin VB.CommandButton btDev 
      Caption         =   "Devolver"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1260
      Width           =   1035
   End
   Begin VB.CommandButton btAdic 
      Caption         =   "Obter"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   840
      Width           =   1035
   End
   Begin VB.ComboBox cbFerr 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   900
      Width           =   2355
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5985
      Left            =   60
      TabIndex        =   5
      Top             =   1680
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   10557
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label lbQuant 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   540
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade: "
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   8
      Top             =   600
      Width           =   915
   End
   Begin VB.Label lbNome 
      Caption         =   "Nome: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   660
      TabIndex        =   7
      Top             =   120
      Width           =   3795
   End
   Begin VB.Label Label2 
      Caption         =   "Nome: "
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Ferramenta: "
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "FerrMec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'3.7.7 Aumento do limite de ferramentas cadastradas de 50 para 200
'3.7.4 Ferramentas de mecânicos permite conjunto de ferramentas para cada mecânico
'3.7.1 Mostrar a quantidade das ferramentas
'3.7.0 Ferramentas

Option Explicit

Private Mudando   As Boolean
Private lcnrFunc  As Integer
Private tAdic     As Integer
Private tRemov    As Integer
Private lcNome    As String
Private sAdic     As String

'3.7.7 Aumento do limite de ferramentas cadastradas de 50 para 200
Private Const MaxFerr = 200
Private Adic(MaxFerr)  As String
Private Remov(MaxFerr) As String
'Private Adic(50)  As String
'Private Remov(50) As String

'3.7.1 Mostrar a quantidade das ferramentas
Private lcQuantidade As Integer

Public Property Get nrFunc() As Integer
nrFunc = lcnrFunc
End Property

Public Property Let nrFunc(ByVal vNewValue As Integer)
Dim Linha  As Integer
Dim SQL    As String
Dim tbGrid As Recordset

lcnrFunc = vNewValue
MSFlexGrid1.ColWidth(0) = 1
MSFlexGrid1.ColWidth(1) = 1000
MSFlexGrid1.ColWidth(2) = 2100
MSFlexGrid1.ColWidth(3) = 1000
MSFlexGrid1.TextMatrix(0, 1) = "Código"
MSFlexGrid1.TextMatrix(0, 2) = "Descrição"
MSFlexGrid1.TextMatrix(0, 3) = "Data"
SQL = "SELECT FerrMec.ID, FerrMec.codigo, Ferramentas.Descricao, FerrMec.Data "
SQL = SQL & "FROM FerrMec "
SQL = SQL & "INNER JOIN Ferramentas ON FerrMec.codigo = Ferramentas.Codigo "
SQL = SQL & "WHERE FerrMec.idMec = " & nrFunc
AbreTB tbGrid, SQL, dbOpenDynaset
If tbGrid.EOF = False Then
    tbGrid.MoveLast
    MSFlexGrid1.Rows = tbGrid.RecordCount + 1
    tbGrid.MoveFirst
    Do While tbGrid.EOF = False
        Linha = Linha + 1
        MSFlexGrid1.TextMatrix(Linha, 0) = tbGrid!ID
        MSFlexGrid1.TextMatrix(Linha, 1) = tbGrid!Codigo
        MSFlexGrid1.TextMatrix(Linha, 2) = tbGrid!Descricao
        MSFlexGrid1.TextMatrix(Linha, 3) = Format(tbGrid!Data, "DD/MM/YYYY")
        tbGrid.MoveNext
    Loop
End If

'3.7.4 Ferramentas de mecânicos permite conjunto de ferramentas para cada mecânico
AtualizaFerr

'3.7.1 Mostrar a quantidade das ferramentas
Quantidade = Linha
End Property

Public Property Let Nome(ByVal vNewValue As String)
lcNome = vNewValue
lbNome.Caption = lcNome
End Property

Public Property Get Nome() As String
Nome = lcNome
End Property

Private Sub AtualizaFerr()
Dim SQL    As String
Dim TbFerr As Recordset

Mudando = True
'3.7.4 Ferramentas de mecânicos permite conjunto de ferramentas para cada mecânico
SQL = "SELECT Ferramentas.Descricao "
SQL = SQL & "from Ferramentas "
SQL = SQL & "WHERE Descricao Is Not Null "
SQL = SQL & "AND Codigo Not In "
SQL = SQL & "(SELECT Codigo "
SQL = SQL & "from FerrMec "
SQL = SQL & "Where idMec = " & nrFunc
SQL = SQL & ") "
If sAdic > "" Then
    SQL = SQL & " and Codigo not in ( " & Left(sAdic, Len(sAdic) - 1) & ")"
End If
SQL = SQL & "ORDER BY Descricao "
'SQL = "Select Descricao "
'SQL = SQL & "From Ferramentas "
'SQL = SQL & "Where Func is null "
'SQL = SQL & " and Descricao Is not Null "
'If sAdic > "" Then
'    SQL = SQL & " and Codigo not in ( " & Left(sAdic, Len(sAdic) - 1) & ")"
'End If

AbreTB TbFerr, SQL, dbOpenDynaset
cbFerr.Clear
Do While TbFerr.EOF = False
    cbFerr.AddItem TbFerr.Fields("Descricao")
    TbFerr.MoveNext
Loop
TbFerr.Close
Mudando = False
End Sub

Private Sub btAdic_Click()
Dim Linha As Integer
Dim Codigo As String

Mudando = True
Linha = MSFlexGrid1.Rows - 1
If MSFlexGrid1.TextMatrix(Linha, 1) > "" Then
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    Linha = Linha + 1
End If
Codigo = Consulta("Select Codigo From Ferramentas Where Descricao = '" & cbFerr.Text & "'")
MSFlexGrid1.TextMatrix(Linha, 1) = Codigo
MSFlexGrid1.TextMatrix(Linha, 2) = cbFerr.Text
MSFlexGrid1.TextMatrix(Linha, 3) = Format(Now, "DD/MM/YYYY")
tAdic = tAdic + 1
Adic(tAdic) = cbFerr.Text
sAdic = sAdic & "'" & Codigo & "',"
AtualizaFerr
btImpr.Enabled = True
Mudando = False

'3.7.1 Mostrar a quantidade das ferramentas
Quantidade = Quantidade + 1

'3.7.7 Aumento do limite de ferramentas cadastradas de 50 para 200
If tAdic > 199 Then
    msgboxL "Limite de ferramentas alcançado" & Chr(13) & "Contacte o programador"
    btAdic.Enabled = False
End If
End Sub

Private Function Codigo_Descr(Codigo As String)
Codigo_Descr = Consulta("Select Codigo From Ferramentas Where Descricao = " & FA(Codigo))
End Function

Private Sub btDev_Click()
Dim Linha As Integer
Dim Codigo As String

Mudando = True
Linha = MSFlexGrid1.Row

tRemov = tRemov + 1
Remov(tRemov) = MSFlexGrid1.TextMatrix(Linha, 2)

'3.7.4 Ferramentas de mecânicos permite conjunto de ferramentas para cada mecânico
cbFerr.AddItem MSFlexGrid1.TextMatrix(Linha, 2)

MSFlexGrid1.RemoveItem Linha
btImpr.Enabled = True
Mudando = False

'3.7.1 Mostrar a quantidade das ferramentas
Quantidade = Quantidade - 1
End Sub

Private Sub btImpr_Click()
Dim a      As Integer
Dim SQL    As String
Dim Codigo As String

ImprBuferizada_Inicializa
LPRINT "O mecanico: " & Nome
If tAdic Then
    LPRINT "Assume que esta sob sua guarda as ferramentas abaixo"
    LPRINT ""
    LPRINT ComplStr("Codigo", 10, " ", 0) & "Descricao"
    For a = 1 To tAdic
        LPRINT ComplStr(Codigo_Descr(Adic(a)), 10, " ", 0) & Adic(a)
    Next
    LPRINT ""
End If
If tRemov Then
    LPRINT "Informa que devolveu as ferramentas abaixo"
    LPRINT ""
    LPRINT ComplStr("Codigo", 10, " ", 0) & "Descricao"
    For a = 1 To tRemov
        LPRINT ComplStr(Codigo_Descr(Remov(a)), 10, " ", 0) & Remov(a)
    Next
    LPRINT ""
End If
LPRINT String(TamFita, "-")
If ImprBuferizada_Finaliza = False Then
    Exit Sub
End If

If tAdic Then
    For a = 1 To tAdic
        Codigo = FA(Codigo_Descr(Adic(a)))
        SQL = "Insert Into FerrMec (idMec, codigo, Data) Values ( " & nrFunc
        SQL = SQL & "," & Codigo
        SQL = SQL & "," & DTSqld(Int(Now)) & ")"
        ExecSql SQL
        
        SQL = "Update Ferramentas Set Func = " & FA(Nome)
        SQL = SQL & " Where codigo = " & Codigo
        ExecSql SQL
    Next
End If
If tRemov Then
    For a = 1 To tRemov
        Codigo = FA(Codigo_Descr(Remov(a)))
        SQL = "Delete From FerrMec Where idMec = " & nrFunc
        SQL = SQL & " and codigo = " & Codigo
        ExecSql SQL
        
        SQL = "Update Ferramentas Set Func = null "
        SQL = SQL & " Where codigo = " & Codigo
        ExecSql SQL
    Next
End If

tRemov = 0
tAdic = 0
Unload Me
End Sub

Private Sub cbFerr_Click()
btAdic.Enabled = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
'3.7.4 Ferramentas de mecânicos permite conjunto de ferramentas para cada mecânico
'AtualizaFerr
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If (tAdic + tRemov) > 0 Then
    If MsgBox("Atenção não foi realizada a impressão" + vbCrLf + "As alterações não serão salvas", vbDefaultButton2 + vbYesNo + vbQuestion, "Tem certeza que quer sair agora?") = vbNo Then
        Cancel = 1
        Exit Sub
    End If
    
    '3.7.1 Mostrar a quantidade das ferramentas
    tAdic = 0
    tRemov = 0
    
End If
End Sub

Private Sub MSFlexGrid1_Click()
If Mudando = False Then
    btDev.Enabled = True
End If
End Sub

'3.7.1 Mostrar a quantidade das ferramentas
Private Property Get Quantidade() As Integer
Quantidade = lcQuantidade
End Property

'3.7.1 Mostrar a quantidade das ferramentas
Private Property Let Quantidade(ByVal vNewValue As Integer)
lcQuantidade = vNewValue
lbQuant.Caption = STR(lcQuantidade)
End Property
