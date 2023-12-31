VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAjustes 
   Caption         =   "Ajustes nos Orçamentos"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5295
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5295
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btExclui 
      Caption         =   "&Excluir"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2820
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Botao 
      Caption         =   "&Transferir"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Ajustes.frx":0000
      Left            =   1620
      List            =   "Ajustes.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   420
      Width           =   3615
   End
   Begin VB.ComboBox cbCliente 
      Height          =   315
      ItemData        =   "Ajustes.frx":0004
      Left            =   1620
      List            =   "Ajustes.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   60
      Width           =   3615
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1785
      Left            =   60
      TabIndex        =   4
      Top             =   780
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3149
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cliente de Destino: "
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cliente de Origem: "
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   1335
   End
End
Attribute VB_Name = "frmAjustes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'3.8.6 Operação de adaptação de orçamentos

Option Explicit

Private strcbCliente As String
Private CliMostrado As String

Private Sub Botao_Click()
Dim SQL As String

Screen.MousePointer = vbHourglass
SQL = "Update Orcamento "
SQL = SQL & " Set cliente = " & FA(Trim(Combo1.Text))
SQL = SQL & " Where Orcamento = " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
ExecSql SQL
Screen.MousePointer = vbDefault
CliMostrado = ""
Unload Me
End Sub

Private Sub btExclui_Click()
Dim SQL As String

Screen.MousePointer = vbHourglass
SQL = "Delete From Orcamento "
SQL = SQL & " Where Orcamento = " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
ExecSql SQL
Screen.MousePointer = vbDefault
CliMostrado = ""
Unload Me
End Sub

Private Sub cbCliente_Click()
If cbCliente.Text <> CliMostrado Then
    MostraOrcDoCLi
End If
End Sub

Private Sub cbCliente_LostFocus()
If cbCliente.Text > "" Then
    strcbCliente = ""
    If cbCliente.Text <> CliMostrado Then
        MostraOrcDoCLi
    End If
End If
End Sub

Private Sub cbCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    strcbCliente = ""
ElseIf KeyAscii = 27 Then
    Unload Me
Else
    TrataCombo strcbCliente, cbCliente, KeyAscii
End If
End Sub

Private Sub MostraOrcDoCLi()
Dim a       As Integer
Dim b       As Integer
Dim SQL     As String
Dim rsOrcs  As Recordset
Dim ConfATU As Variant
Dim ConfANT As Variant

Screen.MousePointer = vbHourglass
SQL = "SELECT Orcamento.Cliente, Orcamento.Data, Orcamento.Total, Orcamento.Pagamento, Orcamento.Carro, Carros.Modelo, Orcamento.Orcamento "
SQL = SQL & " FROM Orcamento INNER JOIN Carros ON Orcamento.Carro = Carros.Placa "
SQL = SQL & "WHERE Orcamento.Cliente = '" & Trim(cbCliente.Text) & "'"
AbreTB rsOrcs, SQL, dbOpenSnapshot
MSFlexGrid1.Clear
MSFlexGrid1.TextMatrix(0, 0) = "Data"
MSFlexGrid1.TextMatrix(0, 1) = "Valor"
MSFlexGrid1.TextMatrix(0, 2) = "Pago"
MSFlexGrid1.TextMatrix(0, 3) = "Carro"
MSFlexGrid1.TextMatrix(0, 4) = "Placa"
MSFlexGrid1.ColWidth(5) = 0
a = 1
On Local Error GoTo NaoTem
rsOrcs.MoveLast
On Local Error GoTo 0
rsOrcs.MoveFirst
MSFlexGrid1.Rows = rsOrcs.RecordCount + 1
Do While rsOrcs.EOF = False
    ConfATU = Empty
    For b = 0 To 3
        ConfATU = ConfATU & rsOrcs(b)
    Next
    If ConfATU <> ConfANT Then
        ConfANT = ConfATU
        MSFlexGrid1.TextMatrix(a, 0) = Format(rsOrcs!Data, "dd/mm/yyyy")
        MSFlexGrid1.TextMatrix(a, 1) = Format(rsOrcs!Total, "###,###.00")
        If rsOrcs!Pagamento > 0 Then
            MSFlexGrid1.TextMatrix(a, 2) = "PAGO"
        Else
            MSFlexGrid1.TextMatrix(a, 2) = " "
        End If
        MSFlexGrid1.TextMatrix(a, 3) = rsOrcs!Modelo
        MSFlexGrid1.TextMatrix(a, 4) = rsOrcs!Carro
        MSFlexGrid1.TextMatrix(a, 5) = rsOrcs!Orcamento
    End If
    rsOrcs.MoveNext
    a = a + 1
Loop
CliMostrado = cbCliente.Text
Combo1.Clear
clsCLi.CarCliCB Combo1, Trim(cbCliente.Text)
btExclui.Enabled = True

NaoTem:
Screen.MousePointer = vbDefault
End Sub

Private Sub Combo1_Click()
Botao.Enabled = (Combo1.Text > "")
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    strcbCliente = ""
ElseIf KeyAscii = 27 Then
    Unload Me
Else
    TrataCombo strcbCliente, Combo1, KeyAscii
End If
End Sub

Private Sub Form_Load()
clsCLi.CarCliCB cbCliente
End Sub
