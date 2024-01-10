VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ReciboAvulso 
   Caption         =   "Estoque - Recibo (ActiveReport)"
   ClientHeight    =   12990
   ClientLeft      =   60
   ClientTop       =   1395
   ClientWidth     =   19080
   WindowState     =   2  'Maximized
   _ExtentX        =   36671
   _ExtentY        =   22913
   SectionData     =   "ReciboAvulso.dsx":0000
End
Attribute VB_Name = "ReciboAvulso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"ActiveReport"
'4.8.4 ReImpressão dos Vales
'4.7.0 Retirada da observação dupla no recibo avulso
'4.6.4 Impressão da observação no Recibo Avulso em folha inteira
'4.6.0 Impressão da observação do recibo avulso
'4.3.7 Tela para recibo avulso

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
txDia.Text = "Porto Alegre, " & Day(Now) & " de " & MesExtenso(Now) & " de " & Year(Now)
'txDia.Text = "Porto Alegre, " & Day(Now) & " de " & MesExtenso & " de " & Year(Now)

End Sub

Private Sub ActiveReport_KeyUp(KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

'4.7.0 Retirada da observação dupla no recibo avulso
Public Sub ReciboAvulso(Det As String, Valor As Double, Destinatario As String, Ident As String, Endereco As String, QuemPaga As String)
'Public Sub ReciboAvulso(Det As String, Valor As Double, Destinatario As String, Ident As String, Endereco As String, QuemPaga As String, Obs As String)

txQuemPaga.Text = txQuemPaga & " " & QuemPaga
txQuemRecebe.Text = txQuemRecebe.Text & " " & Destinatario
txIdent.Text = txIdent.Text & Ident
txEnderMec.Text = txEnderMec.Text & " " & Endereco

'4.6.0 Impressão da observação do recibo avulso
If Det > "" Then
    txReferente.Visible = True
    txReferente.Text = txReferente.Text & " " & Det
End If

'4.7.0 Retirada da observação dupla no recibo avulso
'4.6.4 Impressão da observação no Recibo Avulso em folha inteira
'If Obs > "" Then
'    txObs.Text = "Observação: " & Obs
'    txObs.Visible = True
'End If

ColocaValor Valor
End Sub

Private Sub ColocaValor(Valor As Double)
txValor.Text = "Valor: " & Format(Valor, "##,###,###,##0.00")
txExtenso.Text = "(" & Extenso(CSng(Valor)) & ")"
End Sub
