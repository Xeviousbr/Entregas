VERSION 5.00
Object = "{00028C4A-0000-0000-0000-000000000046}#5.0#0"; "TDBG5.OCX"
Begin VB.Form PagtoForn 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Boleto"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3120
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   3120
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ckEmail 
      Caption         =   "Enviar email"
      Height          =   255
      Left            =   60
      TabIndex        =   10
      ToolTipText     =   "Envia email ao fornecedor com o conteúdo do boleto gerado"
      Top             =   4260
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Observação"
      Height          =   1095
      Left            =   60
      TabIndex        =   8
      Top             =   3120
      Width           =   2955
      Begin VB.TextBox txObs 
         Alignment       =   1  'Right Justify
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   2715
      End
   End
   Begin VB.TextBox txDoc 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   600
      TabIndex        =   7
      Top             =   2760
      Width           =   2415
   End
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
      Top             =   1740
      Visible         =   0   'False
      Width           =   2775
   End
   Begin TrueDBGrid50.TDBGrid TDBGrid1 
      Bindings        =   "PagtoForn.frx":0000
      Height          =   1815
      Left            =   60
      OleObjectBlob   =   "PagtoForn.frx":0014
      TabIndex        =   4
      Top             =   840
      Width           =   2955
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1980
      TabIndex        =   3
      Top             =   4560
      Width           =   1035
   End
   Begin VB.CommandButton btAdic 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   4560
      Width           =   1035
   End
   Begin VB.TextBox txEntrada 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Top             =   60
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Doc:"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   2820
      Width           =   495
   End
   Begin VB.Label lbAFaltante 
      Caption         =   "Valor que falta a pagar: R$ "
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Label Label2 
      Caption         =   "Valor total a pagar:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1395
   End
End
Attribute VB_Name = "PagtoForn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'4.9.6 Previsão para data em branco no boleto
'4.6.3 Conserto do boleto em branco
'4.5.1 Campos DOC, Obs e CPF no Boleto

Option Explicit

Private l_OK          As Boolean
Private Parcela       As Integer
Private l_Fornec      As Integer
Private l_idPagtoForn As Long
Private l_Valor       As Currency
Private Restante      As Currency
Private ValorAPagar   As Double
Private VlrParcelas   As Double
Private l_DOC         As String
Private l_Obs         As String

Public Property Get Valor() As Currency
Valor = l_Valor
End Property

Public Property Let Valor(ByVal vNewValue As Currency)
l_Valor = vNewValue
Restante = l_Valor
End Property

Private Sub btAdic_Click()
Dim a     As Integer
Dim Row   As Integer
Dim Quant As Integer

'4.9.6 Previsão para data em branco no boleto
Row = TDBGrid1.Row
If Row > Quant Then
    Quant = Row
End If
For a = 0 To (Quant - 1)
    TDBGrid1.Row = a
    TDBGrid1.Col = 1
    If TDBGrid1.Text = "" Then
        MsgBox "Data esta em branco é necessário informar"
        Exit Sub
    End If
Next

'4.5.1 Campos DOC, Obs e CPF no Boleto
DOC = txDoc.Text
Obs = txObs.Text
Valor = ValorAPagar
OK = True
Me.Hide
End Sub

Private Sub Command1_Click()
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
    
    '4.6.3 Conserto do boleto em branco
    TDBGrid1.Col = 1
    'TDBGrid1.Col = 0
    
    sData = TDBGrid1.Text
    
    ', Data
    If CriticaData(sData, Data) = 0 Then
    
        'Isso deveria funcionar
        Save = 0
        
        'Isso me parece irrelevante
        Exit Sub
    Else
    
        '4.6.3 Conserto do boleto em branco
        If Data < Int(Now) Then
        'If Data < Now Then
        
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
VlrParcelas = 0
InicForm Me
ExecSql "Delete From ParcelasTemp Where IDPC = " & INI.PC
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
VeValor txEntrada.Text, ValorAPagar, txEntrada, 0
Restante = ValorAPagar - VlrParcelas
If Restante <= 0 Then
    Restante = 0
    lbAFaltante.Visible = False
Else
    lbAFaltante.Caption = "Valor que falta a pagar: R$ " & Format(Restante, "###,###.00")
    lbAFaltante.Visible = True
End If
btAdic.Enabled = Not (lbAFaltante.Visible)
End Sub

Public Property Get Fornec() As Integer
Fornec = l_Fornec
End Property

Public Property Let Fornec(ByVal vNewValue As Integer)
l_Fornec = vNewValue

idPagtoForn = Consulta("Select Max(idPagtoForn) From PagtoForn") + 1

DataP.DatabaseName = Base
DataP.RecordSource = "Select Auto, IDPC, Data, Valor From ParcelasTemp Where IDPC = " & INI.PC
DataP.Enabled = True
DataP.Refresh
End Property

Public Property Get OK() As Boolean
OK = l_OK
End Property

Public Property Let OK(ByVal vNewValue As Boolean)
l_OK = vNewValue
End Property

Public Property Get idPagtoForn() As Long
idPagtoForn = l_idPagtoForn
End Property

Public Property Let idPagtoForn(ByVal vNewValue As Long)
l_idPagtoForn = vNewValue
End Property

'4.5.1 Campos DOC, Obs e CPF no Boleto
Public Property Get DOC() As String
DOC = l_DOC
End Property

Public Property Let DOC(ByVal vNewValue As String)
l_DOC = vNewValue
End Property

Public Property Get Obs() As String
Obs = l_Obs
End Property

Public Property Let Obs(ByVal vNewValue As String)
l_Obs = vNewValue
End Property
