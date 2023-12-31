VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Orcarro.Chart Chart1 
      Height          =   5775
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   10186
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'3.0.5 No gráfico de totais passar a considerar como período curto intervalo de menos de 60 dias
'2.9.8 Opção para visualização mensal do gráfico
'2.9.7 Relatório de Totais Resumido

Option Explicit

Private Itens          As Integer
Private Max            As Single
Private Min            As Single
Private T(100)         As Single
Private P(100)         As Single
Private M(100)         As Single
Private DTFim          As String
Private MyCaption(100) As String

Public Sub RefreshGraph()
Dim myArray(2)  As String
Dim MyColor(2)  As Long
Dim MyLegend(2) As String
Dim a           As Integer

MyColor(0) = vbGreen
MyColor(1) = vbRed
MyColor(2) = vbBlue

For a = 1 To Itens
    myArray(0) = myArray(0) & Trim(STR(T(a))) & ","
    myArray(1) = myArray(1) & Trim(STR(P(a))) & ","
    myArray(2) = myArray(2) & Trim(STR(M(a))) & ","
Next
myArray(0) = LetraMenosUm(myArray(0))
myArray(1) = LetraMenosUm(myArray(1))
myArray(2) = LetraMenosUm(myArray(2))
        
Chart1.MaxValue = Max * 1.5
Chart1.MinValue = Min
Chart1.Rows = Itens
Chart1.Cols = 3
Chart1.DrawGraph myArray, MyColor, MyCaption

MyLegend(0) = "Totais"
MyLegend(1) = "Peças"
MyLegend(2) = "Mecânica"

Chart1.DrawLegend MyLegend, MyColor, "Valores em Milhares de Reais"
End Sub

Private Sub ProcPorMes()
Dim MenorAgora As Long
Dim SQL        As String
Dim MesAgora   As Date
Dim EsseMes    As Date
Dim rsGrafico  As Recordset

Const Div = 1000

SQL = "SELECT Orcamento.Pagamento, First(Orcamento.Total) AS Total, Sum(Itens_Orc.Valor) AS Pecas "
SQL = SQL & "from Orcamento, Itens_Orc "
SQL = SQL & "Where Itens_Orc.Orçamento = Orcamento.Orcamento "
SQL = SQL & "and Orcamento.Pagamento > 0 "
SQL = SQL & "and Orcamento.Pagamento BetWeen " & DTSqls(frmRelTotais.txDtIni.Text) & " and " & DTFim
SQL = SQL & " GROUP BY Orcamento.Pagamento "
SQL = SQL & "ORDER BY Orcamento.Pagamento "
AbreTB rsGrafico, SQL, dbOpenForwardOnly
EsseMes = -1
Do While rsGrafico.EOF = False
    MesAgora = DateSerial(Year(rsGrafico!Pagamento), Month(rsGrafico!Pagamento), 1)
    If EsseMes < MesAgora Then
        EsseMes = MesAgora
        If Itens > 1 Then
            If Max < T(Itens) Then
                Max = T(Itens)
            End If
            MenorAgora = IIf(P(Itens) < M(Itens), P(Itens), M(Itens))
            If Min < MenorAgora Then
                Min = MenorAgora
            End If
        Else
            Min = IIf(P(Itens) < M(Itens), P(Itens), M(Itens))
        End If
        Itens = Itens + 1
        T(Itens) = rsGrafico!Total / Div
        P(Itens) = rsGrafico!Pecas / Div
        M(Itens) = T(Itens) - P(Itens)
        MyCaption(Itens) = Format(rsGrafico!Pagamento, "MMM/YY")
    Else
        T(Itens) = T(Itens) + rsGrafico!Total / Div
        P(Itens) = P(Itens) + rsGrafico!Pecas / Div
        M(Itens) = T(Itens) - P(Itens)
    End If
    rsGrafico.MoveNext
Loop
End Sub

Private Sub ProcPorDia()
'2.9.8 Opção para visualização mensal do gráfico
Dim MenorAgora As Long
Dim SQL        As String
Dim DiaAgora   As Date
Dim EsseDia    As Date
Dim rsGrafico  As Recordset

SQL = "SELECT Orcamento.Pagamento, First(Orcamento.Total) AS Total, Sum(Itens_Orc.Valor) AS Pecas "
SQL = SQL & "from Orcamento, Itens_Orc "
SQL = SQL & "Where Itens_Orc.Orçamento = Orcamento.Orcamento "
SQL = SQL & "and Orcamento.Pagamento > 0 "
SQL = SQL & "and Orcamento.Pagamento BetWeen " & DTSqls(frmRelTotais.txDtIni.Text) & " and " & DTFim
SQL = SQL & " GROUP BY Orcamento.Pagamento "
SQL = SQL & "ORDER BY Orcamento.Pagamento "
AbreTB rsGrafico, SQL, dbOpenForwardOnly
EsseDia = -1
Do While rsGrafico.EOF = False
    DiaAgora = Int(rsGrafico!Pagamento)
    If EsseDia < DiaAgora Then
        EsseDia = DiaAgora
        If Itens > 1 Then
            If Max < T(Itens) Then
                Max = T(Itens)
            End If
            MenorAgora = IIf(P(Itens) < M(Itens), P(Itens), M(Itens))
            If Min < MenorAgora Then
                Min = MenorAgora
            End If
        Else
            Min = IIf(P(Itens) < M(Itens), P(Itens), M(Itens))
        End If
        Itens = Itens + 1
        T(Itens) = rsGrafico!Total
        P(Itens) = rsGrafico!Pecas
        M(Itens) = T(Itens) - P(Itens)
        MyCaption(Itens) = Format(rsGrafico!Pagamento, "DD/MMM")
    Else
        T(Itens) = T(Itens) + rsGrafico!Total
        P(Itens) = P(Itens) + rsGrafico!Pecas
        M(Itens) = T(Itens) - P(Itens)
    End If
    rsGrafico.MoveNext
Loop
Max = Max * 1.7
End Sub

Public Sub Processa()
'2.9.8 Opção para visualização mensal do gráfico
DTFim = DTSqls(Format(DateValue(frmRelTotais.txDtFim.Text) + 1, "DD/MM/YYYY"))

'3.0.5 No gráfico de totais passar a considerar como período curto intervalo de menos de 60 dias
If DateValue(frmRelTotais.txDtFim.Text) - DateValue(frmRelTotais.txDtIni.Text) < 60 Then
'If DateValue(frmRelTotais.txDtFIM.Text) - DateValue(frmRelTotais.txDtINI.Text) < 33 Then

    ProcPorDia
Else
    ProcPorMes
End If
Show
Chart1.Width = Me.Width
Chart1.Height = Me.Height * 0.95
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Resize()
Dim Margem As Long

Margem = Me.Height * 0.05
If Margem < 500 Then Margem = 500
Chart1.Width = Me.Width
Chart1.Height = Me.Height - Margem
End Sub

Private Sub Form_Unload(Cancel As Integer)
Itens = 0
Max = 0
Erase T
Erase P
Erase M
Erase MyCaption
End Sub
