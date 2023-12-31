VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'4.8.8 Mostrar as observação da impressão dos vales
'4.8.4 ReImpressão dos Vales
'4.8.0 Compensação dos vales na operação de pagamento
'4.6.1 Correção da situação em que a observação não era impressa no pagamento
'4.3.7 Tela para recibo avulso
'4.3.5 Campos de Identificação e Endereço no recibo avulso
'4.3.0 Recibo avulso
'4.2.1 Impressão da observação
'4.0.8 Correção da impressão da comissão, depois da versão 4.0.5
'4.0.5 ReImpressão individual dos vales

Option Explicit

Public Sub ReciboFita(Mecanico As String, Endereco As String, Det As String, Valor As Double, Folga As String, VisFolga As Boolean, Semana As String, nrMec As Integer, Gravar As Boolean)
Dim sValor   As Single

'4.2.1 Impressão da observação
'Const TamFita = 55

ImprBuferizada_Inicializa

'Recibo
LPRINT CENTRAL("RECIBO", TamFita / 2)

'Empresa
LPRINT CENTRAL(INI.Empresa, TamFita / 2)

'Telefone
LPRINT CENTRAL(INI.Fones, TamFita / 2)

'Endereço
LPRINT CENTRAL(INI.Endereco, TamFita / 2)
If SN(INI.CGC > "", vbString) Then
    LPRINT CENTRAL("CNPj: " & INI.CGC, TamFita / 2)
End If

'--------
LPRINT String(TamFita, "-")

'Nome
LPRINT "Nome: " & Mecanico
LPRINT "Nr do RG: " & RG(Mecanico)

'Endereço do Mecânico
LPRINT "Endereço: " & Endereco
If gTipo = tpAdiantamento Then
    If Det > "" Then
        LPRINT " "
        
        '4.0.8 Correção da impressão da comissão, depois da versão 4.0.5
        Dim Aux As String
        Aux = Det
        LPRINT Aux
        'LPRINT Det
        
        If Det > "" Then
            If Asc(Right(Det, 1)) <> 10 Then
                LPRINT " "
            End If
        End If
    End If
End If

'Valor
If Valor > 0 Then
    LPRINT "Valor: " & Format(Valor, "##,###,###,##0.00")
    sValor = Valor
End If

'Extenso
LPRINT "(" & Extenso(sValor) & ")"

'Folga
If Folga > "" And VisFolga = True Then
    LPRINT Folga
End If
LPRINT " "
LPRINT "Concordo com o valor acima citado"

'Semana
If gTipo = tpComissao Then
    LPRINT "Semana de " & Semana
    If Det > "" Then
        LPRINT " "
        LPRINT "Referente aos serviços nos carros: "
        LPRINT " "
        
        '4.0.8 Correção da impressão da comissão, depois da versão 4.0.5
        Dim Aux1 As String
        Aux1 = Det
        LPRINT Aux1
        'LPRINT Det
        
        If Asc(Right(Det, 1)) <> 10 Then
            LPRINT " "
        End If
    End If
End If
LPRINT " "
'----Porto Alegre, 21 de Outubro de 2012

'4.8.4 ReImpressão dos Vales
LPRINT "Porto Alegre, " & Day(Now) & " de " & MesExtenso(Now) & " de " & Year(Now)
'LPRINT "Porto Alegre, " & Day(Now) & " de " & MesExtenso & " de " & Year(Now)

LPRINT " "
LPRINT " "
LPRINT String(TamFita, "-")
If ImprBuferizada_Finaliza = False Then
    Exit Sub
End If
If gTipo = tpComissao Then
    If Gravar Then
        ExecSql "Update Tarefas Set Pago = Int(Now) Where Mec = " & nrMec & " and Pago is null and Situacao = 3 "
    End If
End If
End Sub

Public Sub ReciboVT(Mecanico As String, Endereco As String, Valor, Semana As String)
Dim sValor   As Single
Dim SQL      As String

'4.2.1 Impressão da observação
'Const TamFita = 55

ImprBuferizada_Inicializa
LPRINT CENTRAL("RECIBO DE VALE TRANSPORTE", TamFita / 2)
LPRINT CENTRAL(INI.Empresa, TamFita / 2)
LPRINT CENTRAL(INI.Fones, TamFita / 2)
LPRINT CENTRAL(INI.Endereco, TamFita / 2)
If SN(INI.CGC > "", vbString) Then
    LPRINT CENTRAL("CNPj: " & INI.CGC, TamFita / 2)
End If
LPRINT String(TamFita, "-")
LPRINT CENTRAL("RECIBO DE VALE TRANSPORTE", TamFita / 2)

'--------
LPRINT String(TamFita, "-")

LPRINT "Nome: " & Mecanico
LPRINT "Nr do RG: " & RG(Mecanico)
LPRINT "Endereço: " & Endereco
If Valor > 0 Then
    LPRINT "Valor: " & Format(Valor, "##,###,###,##0.00")
    sValor = Valor
End If
LPRINT "(" & Extenso(sValor) & ")"
If Semana > "" Then
    LPRINT " "
    LPRINT Semana
    LPRINT " "
End If
LPRINT "Concordo com o valor acima citado"
LPRINT " "
'----Porto Alegre, 21 de Outubro de 2012

'4.8.4 ReImpressão dos Vales
LPRINT "Porto Alegre, " & Day(Now) & " de " & MesExtenso(Now) & " de " & Year(Now)
'LPRINT "Porto Alegre, " & Day(Now) & " de " & MesExtenso & " de " & Year(Now)

LPRINT " "
LPRINT " "
LPRINT String(TamFita, "-")
If ImprBuferizada_Finaliza = False Then
    Exit Sub
End If
End Sub

Public Sub ReciboPagamento(Mecanico As String, Endereco As String, nrMec As Integer, lbSemana As String, Semana As String, Valor As Double, Vale As Double, Det As String, Folga As String, VisFolga As Boolean)
Dim sValor!, SQL$, Aux$
Dim tbVales  As Recordset
Dim Data As String

'4.2.1 Impressão da observação
'Const TamFita = 55

ImprBuferizada_Inicializa
'Recibo
LPRINT CENTRAL("RECIBO de Pagamento", TamFita / 2)

'Empresa
LPRINT CENTRAL(INI.Empresa, TamFita / 2)

'Telefone
LPRINT CENTRAL(INI.Fones, TamFita / 2)

'Endereço
LPRINT CENTRAL(INI.Endereco, TamFita / 2)

If SN(INI.CGC > "", vbString) Then
    LPRINT CENTRAL("CNPj: " & INI.CGC, TamFita / 2)
End If

'--------
LPRINT String(TamFita, "-")

'FUNC: MARCELO
'CARGO: Mecanico
'MÊS: JANEIRO 2013
'         salario 1000#
'            vale 10.00  10-01-13
'            vale   50.00  15-01-13
'            vale   30.00   20-01-13
'
'valor a receber  910.00

LPRINT "Nome: " & Mecanico
LPRINT "Nr do RG: " & RG(Mecanico)
LPRINT "Endereço: " & Endereco
If Consulta("Select Oper From Mecanicos Where codi = " & nrMec & " and Ativo = True") = 0 Then
    LPRINT "Cargo: Mecanico "
Else
    LPRINT "Cargo: Balconista "
End If
If lbSemana = "Semana" Then
    LPRINT "Semana de " & Semana
Else
    LPRINT Semana
End If

'Valor
LPRINT "Salario: " & Format(Valor, "##,##0.00")
sValor! = Valor
If Vale Then

    '4.8.8 Mostrar as observação da impressão dos vales
    SQL$ = "Select * From Vales Where idOperador = " & nrMec & " and Pago = 0 and tipo = 0 Order By Data "
    'SQL$ = "Select * From Vales Where idOperador = " & nrMec & " and Pago = 0 and tipo = 2 Order By Data "
    
    LPRINT "Vales:"
    AbreTB tbVales, SQL, dbOpenDynaset
    Do While tbVales.EOF = False
        Aux$ = Format(tbVales!Valor, "##,##0.00")
        Aux$ = "Vale: " & ComplStr(Aux$, 8, " ", 2) & " " & Format(tbVales!Data, "DD/MM/YYYY") & tbVales!Obs
        LPRINT Left(Aux$, 54)
        tbVales.MoveNext
    Loop
    LPRINT "             ---------"
    LPRINT "Soma dos vales:    " & Format(Vale, "##,##0.00")
    sValor! = sValor! - Vale
End If
If Det > "" Then
    LPRINT " "
    
    '4.0.8 Correção da impressão da comissão, depois da versão 4.0.5
    Dim Aux2 As String
    Aux = Det
    LPRINT Aux2
    
    '4.6.1 Correção da situação em que a observação não era impressa no pagamento
    'A correção consistiu em descomentar a linha abaixo
    LPRINT Det
    
    '4.6.1 Correção da situação em que a observação não era impressa no pagamento
    'e mais esse IF
    If Len(Det) > 0 Then
    
        If Asc(Right(Det, 1)) <> 10 Then
            LPRINT " "
        End If
    End If
End If
LPRINT "Valor a receber: " & Format(sValor!, "##,##0.00")

'Extenso
LPRINT "(" & Extenso(sValor!) & ")"

'Folga
If Folga > "" And VisFolga = True Then
    LPRINT Folga
End If
LPRINT "Concordo com o valor acima citado"

LPRINT " "
'----Porto Alegre, 21 de Outubro de 2012

'4.8.4 ReImpressão dos Vales
LPRINT "Porto Alegre, " & Day(Now) & " de " & MesExtenso(Now) & " de " & Year(Now)
'LPRINT "Porto Alegre, " & Day(Now) & " de " & MesExtenso & " de " & Year(Now)

LPRINT " "
LPRINT " "
LPRINT String(TamFita, "-")
If ImprBuferizada_Finaliza = False Then
    Exit Sub
End If

'4.8.0 Compensação dos vales na operação de pagamento
'SQL$ = "Update Vales Set Pago = "
'SQL$ = SQL$ & DTSqld(Now) & " Where idOperador = "
'SQL$ = SQL$ & nrMec & " and Pago = 0 and Tipo = 0 "
'ExecSql SQL$
End Sub


Private Function RG(Mecanico As String) As String
RG = SN(Consulta("Select RG from Mecanicos Where Nome = '" & Mecanico & "' and Ativo = True "), vbString)
End Function

'4.3.7 Tela para recibo avulso
Public Sub ReciboOutros(Det As String, Valor As Double, Destinatario As String, Ident As String, Endereco As String, QuemPaga As String)
'4.3.5 Campos de Identificação e Endereço no recibo avulso
'Public Sub ReciboOutros(Det As String, Valor As Double, Destinatario As String, Ident As String, Endereco As String)
'txQuemPaga.Text
'4.3.0 Recibo avulso'
'Public Sub ReciboOutros(Det As String, Valor As Double, Destinatario As String)

Dim sValor   As Single

ImprBuferizada_Inicializa

'Recibo
LPRINT CENTRAL("RECIBO", TamFita / 2)

'Empresa
LPRINT CENTRAL(INI.Empresa, TamFita / 2)

'Telefone
LPRINT CENTRAL(INI.Fones, TamFita / 2)

'Endereço
LPRINT CENTRAL(INI.Endereco, TamFita / 2)
If SN(INI.CGC > "", vbString) Then
    LPRINT CENTRAL("CNPj: " & INI.CGC, TamFita / 2)
End If

'--------
LPRINT String(TamFita, "-")

LPRINT "QuemPaga: " & QuemPaga
LPRINT " "

'4.3.7 Tela para recibo avulso
LPRINT "QuemRecebe: " & Destinatario
LPRINT " "
'LPRINT "Nome: " & Destinatario

'4.3.5 Campos de Identificação e Endereço no recibo avulso
If Ident > "" Then
    LPRINT "Identificacao: " & Ident
    LPRINT " "
End If
If Endereco > "" Then
    LPRINT "Endereco: " & Endereco
    LPRINT " "
End If

'Valor
LPRINT "Valor: " & Format(Valor, "##,###,###,##0.00")
sValor = Valor

'Extenso
LPRINT "(" & Extenso(sValor) & ")"

'Detalhe
LPRINT " "
Dim Aux1 As String
Aux1 = Det
LPRINT "Referente a " & Aux1
If Asc(Right(Det, 1)) <> 10 Then
    LPRINT " "
End If
        
LPRINT " "
'----Porto Alegre, 21 de Outubro de 2012

'4.8.4 ReImpressão dos Vales
'LPRINT "Porto Alegre, " & Day(Now) & " de " & MesExtenso & " de " & Year(Now)
LPRINT "Porto Alegre, " & Day(Now) & " de " & MesExtenso(Now) & " de " & Year(Now)

LPRINT " "
LPRINT " "
LPRINT String(TamFita, "-")
ImprBuferizada_Finaliza
End Sub
