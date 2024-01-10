VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ReciboX 
   Caption         =   "Estoque - Recibo (ActiveReport)"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   1395
   ClientWidth     =   15240
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "ReciboX.dsx":0000
End
Attribute VB_Name = "ReciboX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"ActiveReport"
Private TBRecibo As Recordset
   
Private Sub ActiveReport_FetchData(EOF As Boolean)
'If TBRecibo.EOF = True Then
'   EOF = True
'   Exit Sub
'End If
'EOF = False
'Fields("Linha1") = "Recebi de " & Trim(Config.Nome) & " o valor de R$ " & vlimpr(STR(TBRecibo.Fields("SomaDeValor")), 0, 2) & ", referente a pagamento de comissões do período de " & Comissoes.Text & " até " & Comissoes.Text1 & "."
'Fields("Empresa") = Consulta("Select Empresa From Config")
Fields("Empresa") = "teste"
'Fields("Empresa") = Config.Nome
EOF = True
End Sub

Private Sub ActiveReport_KeyUp(KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub
