VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} relTotaisResumido 
   Caption         =   "Relação de Totais"
   ClientHeight    =   10290
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   20108
   _ExtentY        =   18150
   SectionData     =   "RelTotaisResumido.dsx":0000
End
Attribute VB_Name = "relTotaisResumido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2.9.7 Relatório de Totais Resumido

Private Sub ActiveReport_KeyUp(KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

'Private Sub ActiveReport_FetchData(EOF As Boolean)
'Static smMecanica, smEletrica, smItens, smTotal As Currency
'
''2.9.0 Adaptação do relatório de totais as tarefas dinamicas
'Dim cMecanica As Currency
'Dim cEletrica As Currency
'Dim SQL       As String
'
'If EOF Then
'    lbMecanica.Caption = Format(smMecanica, "#.00")
'    lbEletrica.Caption = Format(smEletrica, "#.00")
'    lbItens.Caption = Format(smItens, "#.00")
'    lbTotal.Caption = Format(smTotal, "#.00")
'Else
'
'    '2.9.0 Adaptação do relatório de totais as tarefas dinamicas
'    SQL = "SELECT Sum(Tarefas.Vlr) from Tarefas"
'    SQL = SQL & " Where Tarefas.Orc = " & Linhas.Recordset!Orcamento
'    SQL = SQL & " and Tarefas.Concerto = 0 "
'    SQL = SQL & " GROUP BY Tarefas.Mec "
'    cMecanica = Consulta(SQL)
'    fldMecanica.Text = Format(cMecanica, "##,###,###,##0.00")
'    smMecanica = smMecanica + cMecanica
'
'    SQL = "SELECT Sum(Tarefas.Vlr) from Tarefas"
'    SQL = SQL & " Where Tarefas.Orc = " & Linhas.Recordset!Orcamento
'    SQL = SQL & " and Tarefas.Concerto = 4"
'    SQL = SQL & " GROUP BY Tarefas.Mec "
'    cEletrica = Consulta(SQL)
'    fldEletrica.Text = Format(cEletrica, "##,###,###,##0.00")
'    smEletrica = smEletrica + cEletrica
'    smTotal = smTotal + cMecanica + cEletrica + Linhas.Recordset!TotalItens
'    'smMecanica = smMecanica + Linhas.Recordset!Mecanica
'    'smEletrica = smEletrica + Linhas.Recordset!Eletricidade
'    'smTotal = smTotal + Linhas.Recordset!Mecanica + Linhas.Recordset!Eletricidade + Linhas.Recordset!TotalItens
'
'    smItens = smItens + Linhas.Recordset!TotalItens
'End If
'End Sub
