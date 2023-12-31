VERSION 5.00
Begin VB.Form frmVlrTotais 
   Caption         =   "Valores Totais"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4350
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Height          =   255
      Left            =   4380
      TabIndex        =   14
      Top             =   180
      Width           =   315
   End
   Begin VB.Frame Frame2 
      Height          =   1395
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   2235
      Begin VB.TextBox txDtINI 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   11
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Consultar"
         Height          =   315
         Left            =   600
         TabIndex        =   10
         Top             =   960
         Width           =   1035
      End
      Begin VB.TextBox txDtFIM 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   9
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Data Inicial:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Data Final:"
         Height          =   195
         Index           =   8
         Left            =   300
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   2340
      TabIndex        =   0
      Top             =   60
      Width           =   1995
      Begin VB.TextBox txVlrPg 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Tarefas concluídas e pagas no perídio"
         Top             =   420
         Width           =   915
      End
      Begin VB.TextBox txVlraPg 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Tarefas assumidas no período mas não concluídas"
         Top             =   1020
         Width           =   915
      End
      Begin VB.TextBox txVlrNpgNoInt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Tarefas concluídas no período mas ainda não pagas"
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Aberto"
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   7
         Top             =   420
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "A pagar"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1020
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "Valores"
         Height          =   195
         Index           =   3
         Left            =   960
         TabIndex        =   5
         Top             =   180
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Pago"
         Height          =   195
         Index           =   10
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmVlrTotais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'5.2.1 Ajuste do valor A Pagar na consulta de totais
'5.2.1 Fechar a tela dos valores totais com esc
'5.2.0 Ajuste no campo Pago do relatório de totais
'5.1.9 Ajustes para a tela de consulta de valores
'5.1.8 Tela para consulta de valores totais de orçamentos

Option Explicit

Private Sub Command1_Click()
Dim SQL     As String
Dim Aberto  As Currency
Dim APAgar  As Currency
Dim ParcelasPagas As Currency
Dim VlrPagParc As Currency
Dim ParcelasPagas2 As Currency
Dim a       As Integer
Dim sData   As String
Dim rsParcial   As Recordset
 
If IsDate(txDtINI.Text) = False Or IsDate(txDtFIM.Text) = False Then
    msgboxL "Data Inválida"
Else
    txVlrPg.Text = ""
    txVlrNpgNoInt.Text = ""
    txVlraPg.Text = ""
    sData = DTSqls(txDtINI.Text) & " and " & DTSqls(txDtFIM.Text, True)
        
    'Aberto
    SQL = "SELECT Sum(Total) AS Soma  "
    SQL = SQL & "from Orcamento "
    SQL = SQL & "WHERE VlrPago <> Total "
    SQL = SQL & "and Orcamento.Data BetWeen " & sData
    Aberto = Consulta(SQL)
    MostraValor txVlrPg, Aberto
    
    'Pago
    SQL = "SELECT Sum(Valor) AS SomaDeValor "
    SQL = SQL & "FROM Parcelas "
    
    '5.2.0 Ajuste no campo Pago do relatório de totais
    SQL = SQL & "Where Data BetWeen " & sData
    
    ParcelasPagas = Consulta(SQL)
    MostraValor txVlrNpgNoInt, ParcelasPagas
    
    '5.1.9 Ajustes para a tela de consulta de valores
    SQL = "SELECT Sum(Parcelas.Valor) AS SomaDeValor "
    SQL = SQL & "FROM Parcelas "
    SQL = SQL & "Where Orc in ("
    SQL = SQL & "Select Orcamento from Orcamento Where VlrPago <> Total "
        
    '5.2.1 Ajuste do valor A Pagar na consulta de totais
    SQL = SQL & "and Data BetWeen " & sData
    
    SQL = SQL & ") "
    ParcelasPagas2 = Consulta(SQL)
    
    'A pagar
    SQL = "SELECT (Total -  VlrPago) as x "
    SQL = SQL & "from Orcamento "
    SQL = SQL & "WHERE Data BetWeen " & sData
    SQL = SQL & "and Total <> VlrPago "
    VlrPagParc = Consulta(SQL)
    APAgar = Aberto - ParcelasPagas2
    MostraValor txVlraPg, APAgar
End If
End Sub

Private Sub Command2_Click()
'5.2.1 Fechar a tela dos valores totais com esc
Unload Me
End Sub

Private Sub Form_Load()
Dim DT As Date

InicForm Me
DT = DtINIRel()
txDtINI.Text = Format(DT, "dd/mm/yyyy")
txDtFIM.Text = Format(Now, "dd/mm/yyyy")
End Sub
