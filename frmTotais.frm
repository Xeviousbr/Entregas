VERSION 5.00
Begin VB.Form frmRelTotais 
   Caption         =   "Relação de Totais"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3540
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Relatório"
      Height          =   375
      Left            =   1140
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txDtFim 
      Height          =   285
      Left            =   2400
      MaxLength       =   20
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txDtINI 
      Height          =   285
      Left            =   1020
      MaxLength       =   20
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lbFuncao 
      Alignment       =   2  'Center
      Caption         =   "Detalhado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "até"
      Height          =   195
      Left            =   2040
      TabIndex        =   3
      Top             =   540
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Apartir de"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   840
   End
End
Attribute VB_Name = "frmRelTotais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2.9.7 Relatório de Totais Resumido
'2.9.0 Adaptação do relatório de totais as tarefas dinamicas
'2.7.2 Logar todas mensagens
'2.4.5 Ajuste do relatório de totais em que não mostrava pagamentos se pedir um só dia de relatório
'2.4.4 Ajuste no relatório de totais para o caso de datas do mesmo dia[2]
'2.4.2 Prevenção para nulos no orçamento
'2.4.2 Ajuste no relatório de totais para o caso de datas do mesmo dia
'2.2.8 Ajuste na seleção da data final no relatório de totais
'2.2.8 Acréscimo do campo da Data da criação do Orçamento, na relação de totais
'2.2.6 Relatório de totais

Private Sub RelatorioDetalhado()
Dim SQL As String

SQL = "Select Count(*) From Orcamento WHERE Pagamento BetWeen " & DTSqls(txDtIni.Text) & " And " & DTSqls(txDtFim.Text, True)
If Consulta(SQL) = 0 Then
    msgboxL "Não há pagamentos neste intervalo de datas", vbInformation, "OrCarro"
Else
    SQL = "SELECT  Orcamento.Orcamento, Orcamento.Data, Orcamento.Total, TotItens.TotalItens, Orcamento.Cliente, Orcamento.Pagamento "
    SQL = SQL & "FROM Orcamento, [SELECT Sum(Itens_Orc.Valor) AS TotalItens, Orçamento FROM Itens_Orc GROUP BY Orçamento, Orçamento]. as TotItens "
    SQL = SQL & "WHERE Orcamento.Pagamento BetWeen " & DTSqls(txDtIni.Text) & " And " & DTSqls(txDtFim.Text, True)
    SQL = SQL & " and TotItens.Orçamento = Orcamento.Orcamento "
    SQL = SQL & "Order By Orcamento.Orcamento"
    
    Load relTotais
    relTotais.lbTitulo.Caption = "Relação de totais entre " & txDtIni.Text & " e " & txDtFim.Text
    relTotais.Linhas.RecordSource = SQL
    relTotais.Linhas.DatabaseName = App.Path + "\OrCarro.mdb"
    relTotais.Show
    Unload Me
End If
End Sub

Private Sub RelatorioResumo()
'2.9.7 Relatório de Totais Resumido
Form1.Processa
Unload Me
End Sub

Private Sub Command1_Click()
'2.9.7 Relatório de Totais Resumido
If Command1.Caption = "Relatório" Then
    RelatorioDetalhado
Else
    RelatorioResumo
End If
End Sub

Public Sub Funcao(Tipo As Integer)
'2.9.7 Relatório de Totais Resumido
If Tipo Then
    lbFuncao.Caption = "Resumido"
    Command1.Caption = "Gráfico"
    txDtIni.Text = Format(Consulta("Select Data From Orcamento Order by Data"), "DD/MM/YYYY")
Else
    Command1.Caption = "Relatório"
    lbFuncao.Caption = "Detalhado"
    txDtIni.Text = Format(Now - 7, "DD/MM/YYYY")
End If
txDtFim.Text = Format(Now, "DD/MM/YYYY")
End Sub

Private Sub txDtFim_GotFocus()
Seleciona
End Sub

Private Sub txDtINI_GotFocus()
Seleciona
End Sub

Private Sub txDtINI_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Unload Me
End If
End Sub
