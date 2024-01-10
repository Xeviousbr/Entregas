VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} Recibo 
   Caption         =   "Estoque - Recibo (ActiveReport)"
   ClientHeight    =   12990
   ClientLeft      =   60
   ClientTop       =   1395
   ClientWidth     =   13260
   WindowState     =   2  'Maximized
   _ExtentX        =   23389
   _ExtentY        =   22913
   SectionData     =   "Recibo.dsx":0000
End
Attribute VB_Name = "Recibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"ActiveReport"
'4.8.4 ReImpressão dos Vales
'4.8.2 Mostrar o termo vale transporte no recibo em folha cheia
'4.8.1 Ajuste na informação do período do relatório do vale transporte
'4.7.5 Alteração da posição da observação no recibo de pagamento
'4.5.9 Impressão da observação no Recibo em folha inteira
'4.3.7 Tela para recibo avulso
'4.3.5 Campos de Identificação e Endereço no recibo avulso
'4.3.4 Recibo gráfico para recibo avulso
'2.7.1 Permitir informar só o vale no recibo
'2.7.0 Conserto da impressão do recibo (erro originado na 2.6.9)
'2.6.9 Passa a não ser necessário informar valor para o recibo
'2.6.8 Log para verificar o endereço na impressão do orçamento
'2.6.6 Campos Folga e Vale no recibo do mecânico

Option Explicit

Private Sub ActiveReport_DataInitialize()
Dim SQL      As String
Dim TBRecibo As Recordset

SQL = "Select Empresa, Fones, Endereco From Config"
AbreTB TBRecibo, SQL, dbOpenDynaset

Empresa.Text = TBRecibo!Empresa
txTelefone.Text = "Telefone: " & TBRecibo!Fones
txEnder.Text = "Endereço: " & TBRecibo!Endereco

'4.8.4 ReImpressão dos Vales
'txDia.Text = "Porto Alegre, " & Day(Now) & " de " & MesExtenso & " de " & Year(Now)
End Sub

Private Sub ActiveReport_KeyUp(KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

'4.3.7 Tela para recibo avulso
''4.3.5 Campos de Identificação e Endereço no recibo avulso
'Public Sub ReciboAvulso(Det As String, Valor As Double, Destinatario As String, Ident As String, Endereco As String)
''4.3.4 Recibo gráfico para recibo avulso
''Public Sub ReciboAvulso(Det As String, Valor As Double, Destinatario As String)
'txNome.Text = "Nome: " & Destinatario
'
''4.3.5 Campos de Identificação e Endereço no recibo avulso
'txEnderMec.Top = txEnderMec.Top + 600
'txEnderMec.Text = "Endereço: " & Endereco
'txSemana.Top = txSemana.Top - 2100
'txSemana.Text = Ident
'txValor.Top = txValor.Top + 950
''txEnderMec.Visible = False
''txSemana.Visible = False
'
'ColocaValor Valor
'End Sub

'4.3.4 Recibo gráfico para recibo avulso
Private Sub ColocaValor(Valor As Double)
txValor.Text = "Valor: " & Format(Valor, "##,###,###,##0.00")
txExtenso.Text = "(" & Extenso(CSng(Valor)) & ")"
End Sub

'4.8.2 Mostrar o termo vale transporte no recibo em folha cheia
Public Sub RecebeDados(Mecanico As String, Encereco As String, Valor As Double, Semana As String, Vale As Double, Folga As String, Obs As String, Tipo As Integer, Data As Date)
'4.5.9 Impressão da observação no Recibo em folha inteira
'Public Sub RecebeDados(Mecanico As String, Encereco As String, Valor As Double, Semana As String, Vale As Double, Folga As String, Obs As String)
'Public Sub RecebeDados(Mecanico As String, Encereco As String, Valor As Double, Semana As String, Vale As Double, Folga As String)
'2.6.6 Campos Folga e Vale no recibo do mecânico
Dim sValor As Single

'4.8.2 Mostrar o termo vale transporte no recibo em folha cheia
If Tipo = 2 Then
    txTitulo.Text = "Recibo de Vale Transporte"
End If

txNome.Text = "Nome: " & Mecanico

txEnderMec.Text = "Endereço: " & Encereco

'2.6.8 Log para verificar o endereço na impressão do orçamento
Loga "Endereço sendo impresso no recibo: " & txEnderMec.Text

'2.7.0 Conserto da impressão do recibo (erro originado na 2.6.9)
If Valor > 0 Then

    '4.3.4 Recibo gráfico para recibo avulso
    ColocaValor Valor
    If Vale > 0 Then
        ColocaVale Vale
    Else
        sValor = Valor
        txExtenso.Top = txVale.Top
    End If
'2.6.9 Passa a não ser necessário informar valor para o recibo
'If Valor < 0 Then

'    txValor.Text = "Valor: " & Format(Valor, "##,###,###,##0.00")
'    If Vale > 0 Then
'        ColocaVale Vale
'        '2.7.1 Permitir informar só o vale no recibo
''        txVale.Text = " - Vale: " & Format(Vale, "##,###,###,##0.00")
''        sValor = Valor - Vale
''        txVale.Visible = True
''        txNovoValor.Text = "Novo Valor: " & Format(sValor, "##,###,###,##0.00")
''        txNovoValor.Visible = True
'    Else
'        sValor = Valor
'        txExtenso.Top = txVale.Top
'    End If
'    txExtenso.Text = "(" & Extenso(sValor) & ")"

ElseIf Vale > 0 Then
    ColocaVale Vale
Else
    txValor.Visible = False
    txExtenso.Visible = False
End If

'4.8.1 Ajuste na informação do período do relatório do vale transporte
txSemana.Text = Semana
'txSemana.Text = "Referente a semana de " & Semana

If Folga > "" Then
    txFolga.Text = Folga
    txFolga.Visible = True
End If

'4.5.9 Impressão da observação no Recibo em folha inteira
If Obs > "" Then
    txObs.Text = "Observação: " & Obs
    txObs.Visible = True
End If

'4.8.4 ReImpressão dos Vales
txDia.Text = "Porto Alegre, " & Day(Data) & " de " & MesExtenso(Data) & " de " & Year(Data)

End Sub
Private Sub ColocaVale(Vale As Double)
'2.7.1 Permitir informar só o vale no recibo
Dim sVale As Single

txValor.Visible = False
txVale.Visible = True
txVale.Text = "Vale: " & Format(Vale, "##,###,###,##0.00")
sVale = Vale
txExtenso.Text = "(" & Extenso(sVale) & ")"
End Sub
