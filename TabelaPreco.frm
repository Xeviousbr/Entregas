VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form TabelaPreco 
   Caption         =   "Tabela Preço"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   11865
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   255
      Left            =   12000
      TabIndex        =   8
      Top             =   180
      Width           =   375
   End
   Begin VB.TextBox txTitulo 
      Height          =   315
      Index           =   0
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txTitulo 
      Height          =   315
      Index           =   1
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txTitulo 
      Height          =   315
      Index           =   2
      Left            =   5940
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txTitulo 
      Height          =   315
      Index           =   3
      Left            =   8880
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txBusca 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   2115
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   7425
      Index           =   0
      Left            =   60
      TabIndex        =   6
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   13097
      _Version        =   393216
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   7425
      Index           =   1
      Left            =   3000
      TabIndex        =   7
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   13097
      _Version        =   393216
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   7425
      Index           =   2
      Left            =   5940
      TabIndex        =   9
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   13097
      _Version        =   393216
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid 
      Height          =   7425
      Index           =   3
      Left            =   8880
      TabIndex        =   10
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   13097
      _Version        =   393216
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Pesquisa"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "TabelaPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'3.2.9 Nova tela de tabela de preços

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim rsTitulos As Recordset

AbreTB rsTitulos, "Select TitModelo1, TitModelo2, TitModelo3, TitModelo4 From Config "
If IsNull(rsTitulos!TitModelo1) = False Then
    txTitulo(0).Text = rsTitulos!TitModelo1
End If
If IsNull(rsTitulos!TitModelo2) = False Then
    txTitulo(1).Text = rsTitulos!TitModelo2
End If
If IsNull(rsTitulos!TitModelo3) = False Then
    txTitulo(2).Text = rsTitulos!TitModelo3
End If
If IsNull(rsTitulos!TitModelo4) = False Then
    txTitulo(3).Text = rsTitulos!TitModelo4
End If
Carrega
End Sub

Private Sub Carrega()
Dim a         As Integer
Dim Linha     As Integer
Dim Pesq      As String
Dim SQL       As String
Dim tbGrid(3) As Recordset

For a = 0 To 3
    SQL = "Select Conteudo, Valor From ConfigModelo Where Coluna = " & (a + 1)
    If txBusca.Text > "" Then
        Pesq = "*" & Trim$(txBusca.Text) & "*"
        SQL = SQL & " and Conteudo Like '" & Pesq & "'"
    End If
    SQL = SQL & " Order By Conteudo"
    AbreTB tbGrid(a), SQL, dbOpenDynaset
    If tbGrid(a).EOF = False Then
        tbGrid(a).MoveLast
        MSFlexGrid(a).Rows = tbGrid(a).RecordCount
        MSFlexGrid(a).Cols = 2
        MSFlexGrid(a).ColWidth(0) = 1800
        MSFlexGrid(a).ColWidth(1) = 700
        tbGrid(a).MoveFirst
        Linha = 0
        Do While tbGrid(a).EOF = False
            MSFlexGrid(a).TextMatrix(Linha, 0) = tbGrid(a).Fields(0).Value
            Valor = SN(tbGrid(a).Fields(1).Value, vbCurrency)
            If Valor Then
                MSFlexGrid(a).TextMatrix(Linha, 1) = Format(Valor, "####.00")
            End If
            tbGrid(a).MoveNext
            Linha = Linha + 1
        Loop
    End If
Next
End Sub

Private Sub txBusca_KeyUp(KeyCode As Integer, Shift As Integer)
Dim a As Integer

For a = 0 To 3
    MSFlexGrid(a).Clear
Next
Carrega
End Sub
