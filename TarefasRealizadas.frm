VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form TarefasRealizadas 
   Caption         =   "Taréfas realizadas"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5385
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   5385
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5235
      Begin VB.TextBox txPlaca 
         Height          =   315
         Left            =   2700
         TabIndex        =   14
         Top             =   900
         Width           =   915
      End
      Begin VB.ComboBox cbTipo 
         Height          =   315
         ItemData        =   "TarefasRealizadas.frx":0000
         Left            =   960
         List            =   "TarefasRealizadas.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   900
         Width           =   1155
      End
      Begin VB.CommandButton btPesquisar 
         Caption         =   "Pesquisar"
         Height          =   315
         Left            =   3720
         TabIndex        =   10
         Top             =   900
         Width           =   1455
      End
      Begin VB.CheckBox ckTarefas 
         Caption         =   "Não Concluidas"
         Height          =   195
         Index           =   1
         Left            =   3720
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox ckTarefas 
         Caption         =   "Concluídas"
         Height          =   195
         Index           =   0
         Left            =   3720
         TabIndex        =   7
         Top             =   120
         Width           =   1155
      End
      Begin VB.ComboBox cbMecanico 
         Height          =   315
         ItemData        =   "TarefasRealizadas.frx":0004
         Left            =   960
         List            =   "TarefasRealizadas.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   540
         Width           =   2715
      End
      Begin VB.TextBox txDtFim 
         Height          =   285
         Left            =   2700
         TabIndex        =   4
         Text            =   "31/12/2013"
         ToolTipText     =   "O critério da data corresponde a data do orçamento"
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox txDtIni 
         Height          =   285
         Left            =   1260
         TabIndex        =   2
         Text            =   "31/12/2013"
         ToolTipText     =   "O critério da data corresponde a data do orçamento"
         Top             =   180
         Width           =   975
      End
      Begin VB.CheckBox ckData 
         Caption         =   "Data Inicial"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "O critério da data corresponde a data do orçamento"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Placa"
         Height          =   195
         Index           =   3
         Left            =   2220
         TabIndex        =   13
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Mecânico"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Final"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   3
         ToolTipText     =   "O critério da data corresponde a data do orçamento"
         Top             =   240
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5985
      Left            =   60
      TabIndex        =   9
      Top             =   1380
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   10557
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
End
Attribute VB_Name = "TarefasRealizadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'3.5.7 Impedir que mecânicos vejam orçamentos pela pesquisa de tarefas
'3.2.2 Mostrar o orçamento ao dar duplo clique na pesquisa de tarefas
'3.2.1 Pesquisa de tarefas realizadas

Option Explicit

Private Mudando As Boolean

Private Sub btPesquisar_Click()
Dim Linha  As Integer
Dim SQL    As String
Dim tbGrid As Recordset

Screen.MousePointer = vbHourglass
SQL = "SELECT Mecanicos.Nome, Orcamento.Carro, Tarefas.DtConclusao, Tarefas.Pago "

'3.2.2 Mostrar o orçamento ao dar duplo clique na pesquisa de tarefas
SQL = SQL & ", Orcamento.Orcamento "

SQL = SQL & "FROM Mecanicos, Orcamento, Tarefas "
SQL = SQL & "Where Tarefas.Mec = Mecanicos.codi "
SQL = SQL & "and Orcamento.Orcamento = Tarefas.Orc "
SQL = SQL & "and Tarefas.id > " & INI.Orc1
SQL = SQL & " and Tarefas.Mec > 0 "
If ckData.Value = 1 Then
    SQL = SQL & "and Orcamento.Data Between " & DTSqls(txDtINI.Text) & " And " & DTSqls(txDtFIM.Text, True)
End If
If ckTarefas(0).Value = 0 Or ckTarefas(1).Value = 0 Then
    If ckTarefas(0).Value = 1 Then
        SQL = SQL & " and Tarefas.Situacao = 3 "
    Else
        SQL = SQL & " and Tarefas.Situacao = 2 "
    End If
End If
If cbMecanico.ListIndex > 0 Then
    SQL = SQL & " and Tarefas.Mec = " & cbMecanico.ItemData(cbMecanico.ListIndex)
End If
If cbTipo.ListIndex > 0 Then
    SQL = SQL & " and Tarefas.concerto = " & cbTipo.ItemData(cbTipo.ListIndex)
End If
If txPlaca.Text > "" Then
    SQL = SQL & " and Orcamento.Carro = '" & txPlaca.Text & "' "
End If
SQL = SQL & " Order By Tarefas.id "
AbreTB tbGrid, SQL, dbOpenDynaset
If tbGrid.EOF = False Then
    tbGrid.MoveLast
    MSFlexGrid1.Rows = tbGrid.RecordCount + 1
    Caption = "Tarefas Realizadas " & tbGrid.RecordCount & " registros"
    tbGrid.MoveFirst
    Do While tbGrid.EOF = False
        Linha = Linha + 1
        MSFlexGrid1.TextMatrix(Linha, 0) = tbGrid!Orcamento
        MSFlexGrid1.TextMatrix(Linha, 1) = tbGrid!Nome
        MSFlexGrid1.TextMatrix(Linha, 2) = tbGrid!Carro
        If tbGrid!DtConclusao > 0 Then
            MSFlexGrid1.TextMatrix(Linha, 3) = Format(tbGrid!DtConclusao, "DD/MM/YYYY")
            If tbGrid!DtConclusao > 0 Then
                MSFlexGrid1.TextMatrix(Linha, 4) = Format(tbGrid!PAGO, "DD/MM/YYYY")
            End If
        End If
        tbGrid.MoveNext
    Loop
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub ckData_Click()
If Mudando = False Then
    INI.DtPesqTarefas = ckData.Value
End If
txDtINI.Enabled = ckData.Value
txDtFIM.Enabled = ckData.Value
End Sub

Private Sub ckTarefas_Click(Index As Integer)
If Mudando = False Then
    If Index Then
        INI.TarefasNaoConclu = ckTarefas(1).Value
    Else
        INI.TarefasConclu = ckTarefas(0).Value
    End If
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Dim tbMecs  As Recordset
Dim tbTipos As Recordset

Mudando = True
txDtFIM.Text = Format(Now, "DD/MM/YYYY")
txDtINI.Text = Format(Now - 30, "DD/MM/YYYY")
ckData.Value = INI.DtPesqTarefas
txDtINI.Enabled = ckData.Value
txDtFIM.Enabled = ckData.Value
If INI.TarefasConclu = 0 And INI.TarefasNaoConclu = 0 Then
    ckTarefas(0).Value = 1
    ckTarefas(1).Value = 1
Else
    ckTarefas(0).Value = INI.TarefasConclu
    ckTarefas(1).Value = INI.TarefasNaoConclu
End If
AbreTB tbMecs, "Select codi, Nome From Mecanicos Where Oper = 0 and Nome > '' Order By Nome "
cbMecanico.AddItem "TODOS"
Do While tbMecs.EOF = False
    cbMecanico.AddItem tbMecs!Nome
    cbMecanico.ItemData(cbMecanico.ListCount - 1) = tbMecs!codi
    tbMecs.MoveNext
Loop
cbMecanico.ListIndex = 0
AbreTB tbTipos, "Select tipo, concerto From tpConcertos Order By concerto "
cbTipo.AddItem "TODOS"
Do While tbTipos.EOF = False
    cbTipo.AddItem tbTipos!concerto
    cbTipo.ItemData(cbTipo.ListCount - 1) = tbTipos!Tipo
    tbTipos.MoveNext
Loop
cbTipo.ListIndex = 0

'3.2.2 Mostrar o orçamento ao dar duplo clique na pesquisa de tarefas
MSFlexGrid1.ColWidth(0) = 0

MSFlexGrid1.ColWidth(1) = 2050
MSFlexGrid1.ColWidth(2) = 900
MSFlexGrid1.ColWidth(3) = 1000
MSFlexGrid1.ColWidth(4) = 1000
MSFlexGrid1.TextMatrix(0, 1) = "Mecânico"
MSFlexGrid1.TextMatrix(0, 2) = "Placa"
MSFlexGrid1.TextMatrix(0, 3) = "Conclusão"
MSFlexGrid1.TextMatrix(0, 4) = "Pagamento"
Mudando = False
End Sub

'3.2.2 Mostrar o orçamento ao dar duplo clique na pesquisa de tarefas
Private Sub MSFlexGrid1_DblClick()
Dim Orc As Long

'3.5.7 Impedir que mecânicos vejam orçamentos pela pesquisa de tarefas
If INI.ModoOperacao <> tpMecanico Then
    MSFlexGrid1.Col = 0
    Orc = Val(MSFlexGrid1.Text)
    CarregaTarefas Orc
    Load frmOrc
    frmOrc.NrOrcamento = Orc
    frmOrc.Show
End If
End Sub
