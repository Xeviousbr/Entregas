VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RelFin 
   Caption         =   "Relatório Financeiro"
   ClientHeight    =   10290
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12135
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   21405
   _ExtentY        =   18150
   SectionData     =   "RelFin.dsx":0000
End
Attribute VB_Name = "RelFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'5.0.5 Relatório Financeiro

Private Sub ActiveReport_KeyUp(KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub ActiveReport_FetchData(EOF As Boolean)
Static smValor    As Currency
Static smPago     As Currency
Static smValorTot As Currency
Static smFalta    As Currency
Dim vValor        As Currency
Dim vTotal        As Currency
Dim vPago         As Currency
Dim Dias          As Integer
Dim Data          As Date
Dim SQL           As String

If EOF Then
    lbValorTot.Caption = Format(smValorTot, "#.00")
    lbPago.Caption = Format(smPago, "#.00")
    lbNaoPago.Caption = Format(smValorTot - smPago, "#.00")
Else
    SQL = "SELECT Last(Parcelas.Data) AS Data "
    SQL = SQL & "from Parcelas "
    SQL = SQL & "WHERE Parcelas.Orc = " & Linhas.Recordset!Orcamento
    Data = Consulta(SQL)
        
    If Data = 0 Then
        Dias = Now - Linhas.Recordset!Data
        lbDias.Caption = Dias & " dias de atraso"
    Else
        Dias = Now - Data
        lbDias.Caption = Dias & " dias de atraso"
    End If
    
    If Data > Linhas.Recordset!Pagamento Then
        fldPago.Text = Format(Data, "dd/mm/yyyy")
    Else
        If Linhas.Recordset!Pagamento = 0 Then
            fldPago.Text = ""
        Else
            fldPago.Text = Format(Linhas.Recordset!Pagamento, "dd/mm/yyyy")
        End If
    End If
    vTotal = SN(Linhas.Recordset!Total, vbCurrency)
    vPago = SN(Linhas.Recordset!VlrPago, vbCurrency)
    vValor = vTotal - vPago
    smValorTot = smValorTot + vTotal
    smValor = smValor + vValor
    smPago = smPago + vPago
    smFalta = smFalta + vValor
End If
End Sub
