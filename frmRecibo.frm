VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRecibo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibo do Mec�nico"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4260
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox txDet 
      Height          =   1095
      Left            =   60
      TabIndex        =   11
      Top             =   2160
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1931
      _Version        =   393217
      TextRTF         =   $"frmRecibo.frx":0000
   End
   Begin VB.TextBox txVale 
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Top             =   540
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   1140
      TabIndex        =   3
      Top             =   900
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1140
      TabIndex        =   4
      Top             =   1260
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txValor 
      Height          =   285
      Left            =   1140
      TabIndex        =   1
      Top             =   540
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   435
      Left            =   1620
      TabIndex        =   5
      Top             =   1620
      Width           =   1215
   End
   Begin VB.ComboBox cbMecanico 
      Height          =   315
      ItemData        =   "frmRecibo.frx":0082
      Left            =   1140
      List            =   "frmRecibo.frx":0084
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3075
   End
   Begin VB.Label lbVale 
      AutoSize        =   -1  'True
      Caption         =   "Vale:"
      Height          =   195
      Left            =   2820
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Vlr a receber:"
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Semana:"
      Height          =   195
      Index           =   2
      Left            =   405
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor:"
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   7
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Mec�nico:"
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   6
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "frmRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'4.8.5 ReImpress�o do recibo com todos os campos
'4.8.2 Mostrar o termo vale transporte no recibo em folha cheia
'4.8.0 Compensa��o dos vales na opera��o de pagamento
'4.7.8 Troca do componente da Observa��o do recibo de Text para RichTextBox
'4.5.9 Impress�o da observa��o no Recibo em folha inteira
'4.5.9 Gravar o recibo de adiantamento tamb�m na impress�o em folha inteira
'4.3.7 Tela para recibo avulso
'4.3.5 Campos de Identifica��o e Endere�o no recibo avulso
'4.3.4 Recibo gr�fico para recibo avulso
'4.3.1 Grava��o da observa��o so recibo avulso
'4.3.0 Recibo avulso
'4.0.5 ReImpress�o individual dos vales
'3.8.2 Diminui��o da margem esquerda da impress�o do vale possibilitando maior espa�o para as observa��es
'3.8.0 Impress�o da observa��o na impress�o dos vales no pagamento
'3.7.6 Impedir nome de mec�nico em branco no recibo
'3.6.7 Deixar de imprimir duas vezes o recibo
'3.6.6 tratamento para pc da impressora desligado
'3.6.5 Mais log para a opera��o de criar recibo
'3.6.0 Conserto da impress�o da comiss�o, em que estava imprimindo duas vezes
'3.6.0 Otimiza��o da programa��o relativa a impress�o dos carros no recibo
'3.6.0 Impress�o da observa��o no recibo mensal
'3.5.8 Observa��o para o recibo de adiantamento
'3.5.6 Conserto da impress�o dos vales no pagamento mensal
'3.5.4 Conserto da informa��o do mec�nico no recibo
'3.5.3 Conserto do SQL das comiss�es (3.5.1)
'3.5.2 Corre��o da sele��o dos operadores no recibo (3.5.1)
'3.5.1 N�o excluir fisicamente Mec�nico
'3.5.0 Mostrar o valor a receber no recibo mensal
'3.4.9 Gravar as observa��es do vale
'3.4.8 Adiantamento n�o deve zerar as comiss�es
'3.4.7 Grava��o dos recibos
'3.4.7 Recibo de pagamento mensal
'3.4.8 RG em todos os recibos
'3.4.6 Informa��o do mes do vale transporte
'3.4.4 Vale Transporte
'3.3.7 Impress�o do Vale n�o deve abater imprimir carros nem abater tarefas a receber
'3.3.2 Impedir a possibilidade de continuar valores de comiss�o ap�s a impress�o do recibo
'3.3.0 Crit�rio de quantidade de carros para liberar as comiss�es
'2.9.0 Mostrar os carros das tarefas dos mec�nicos, previamente no recibo
'2.8.8 Mudan�a da cr�tica da libera��o da comiss�o
'2.7.5 Tar�fas Din�micas
'2.7.4-5 Linha para impress�o no recibo em fita
'2.7.4 Impress�o do Recibo em Matricial
'2.7.4 Calculo de Comiss�es no Recibo
'2.7.2 Logar todas mensagens
'2.6.9 Passa a n�o ser necess�rio informar valor para o recibo
'2.6.8 Log para verificar o endere�o na impress�o do or�amento
'2.6.6 Campos Folga e Vale no recibo do mec�nico
'2.6.5 Conserto do Recibo [campo endere�o]

Option Explicit

Private PercComiss   As Single
Private lcValor      As Double
Private lcVale       As Double
Private lcMecanico   As String
Private lcEndereco   As String
Private strcbCliente As String

'2.9.0 Mostrar os carros das tarefas dos mec�nicos, previamente no recibo
Private rsComiss    As Recordset

'3.3.2 Impedir a possibilidade de continuar valores de comiss�o ap�s a impress�o do recibo
Private nrMec As Integer

'4.0.5 ReImpress�o individual dos vales
Private cRecibo As clsRecibo
'3.4.4 Vale Transporte
'Private Enum tpRec
'    tpAdiantamento = 0
'    tpComissao = 1
'    tpValeTransp = 2
'    tpPagamento = 3
'End Enum
'Private gTipo As tpRec

'4.3.0 Recibo avulso
Private JaAtivou As Boolean

Private Sub cbMecanico_Click()
'3.5.4 Conserto da informa��o do mec�nico no recibo
'3.5.1 N�o excluir fisicamente Mec�nico
nrMec = Consulta("Select codi From Mecanicos Where Nome = '" & cbMecanico.Text & "' and Ativo = True ")
'nrMec = Consulta("Select codi From Mecanicos Where Nome = '" & cbMecanico.Text & "'")

Select Case gTipo
    Case tpComissao
        If INI.UtComissoes Then
            If Vale = 0 Then
                VeValores
            End If
        End If
        
    '3.7.6 Mec�nico com recebimento semanal
    Case tpAdiantamento
        VeVales
    Case tpPagamento
        VeVales
        If Consulta("Select tpRec From Mecanicos Where Nome = '" & cbMecanico.Text & "' and Ativo = True ") = 0 Then
            Label1(2).Caption = "Mes"
            Text1(1).Text = Mes()
        Else
            Label1(2).Caption = "Semana"
            Text1(1).Text = CompoemSemana()
        End If
'    Case tpAdiantamento, tpPagamento
'        VeVales

End Select
End Sub

Private Sub VeValores()
'2.7.4 Calculo de Comiss�es no Recibo
Dim a           As Integer
Dim Comiss      As Currency
Dim SQL         As String
Dim SomaTarefas As Double

'2.9.0 Mostrar os carros das tarefas dos mec�nicos, previamente no recibo
'Dim rsComiss    As Recordset

'2.9.0 Mostrar os carros das tarefas dos mec�nicos, previamente no recibo
SQL = "SELECT Tarefas.Vlr, (Tarefas.Vlr * Mecanicos.PercComiss/100) as Comiss "
SQL = SQL & ", Orcamento.Carro, Carros.Modelo, Carros.Cor, Tarefas.ID "

'3.3.2 Impedir a possibilidade de continuar valores de comiss�o ap�s a impress�o do recibo
SQL = SQL & ", Mecanicos.codi "

SQL = SQL & "from Tarefas, Mecanicos, Orcamento, Carros "
SQL = SQL & "WHERE Mecanicos.Nome='" & cbMecanico.Text & "' "
SQL = SQL & "AND Tarefas.Mec=Mecanicos.codi "
SQL = SQL & "AND Tarefas.Situacao=3 "
SQL = SQL & "AND Tarefas.Pago Is Null "
SQL = SQL & "And Orcamento.Orcamento = Tarefas.Orc "
SQL = SQL & "and Carros.Placa = Orcamento.Carro "
'SQL = "SELECT Sum(Tarefas.Vlr) as SomaTarefas, First(Mecanicos.PercComiss) as PercComiss "
'SQL = SQL & "from Tarefas, Mecanicos "
'SQL = SQL & "Where Mecanicos.Nome = '" & cbMecanico.Text & "' "
'SQL = SQL & "And Tarefas.Mec=Mecanicos.codi "
'SQL = SQL & "AND Tarefas.Situacao=3 "
'SQL = SQL & "AND Tarefas.Pago Is Null"
AbreTB rsComiss, SQL, dbOpenDynaset
txDet.Text = ""

If rsComiss.EOF = False Then

    '3.3.2 Impedir a possibilidade de continuar valores de comiss�o ap�s a impress�o do recibo
    nrMec = rsComiss!codi
    
    Do While rsComiss.EOF = False
    
        '3.1.7 Indicar quantos carros foram consertados pelo mec�nico
        a = a + 1
        
        SomaTarefas = SomaTarefas + rsComiss!Vlr
        Comiss = Comiss + rsComiss!Comiss
        txDet.Text = txDet.Text & rsComiss!Modelo & " " & rsComiss!Cor & " " & rsComiss!Carro & vbCrLf
        rsComiss.MoveNext
    Loop
    rsComiss.MoveFirst
End If

'2.9.0 Mostrar os carros das tarefas dos mec�nicos, previamente no recibo
''2.7.5 Tar�fas Din�micas
'On Local Error GoTo TahNulo
'SomaTarefas = SN(rsComiss!SomaTarefas, vbSingle)
'PercComiss = SN(rsComiss!PercComiss, vbSingle)
'Comiss = SomaTarefas * (PercComiss / 100)

'2.8.8 Mudan�a da cr�tica da libera��o da comiss�o
If (SomaTarefas + 0.01) < INI.VlrGatComiss Then
'If (Comiss + 0.01) < INI.VlrGatComiss Then

    Comiss = 0
    
'3.3.0 Crit�rio de quantidade de carros para liberar as comiss�es
ElseIf a < INI.QtdCarrComiss Then
    Comiss = 0
End If
Valor = Comiss
MostraValor txValor, Comiss

TahNulo:
On Local Error GoTo 0
End Sub

Private Sub cbMecanico_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8
        strcbCliente = ""
    Case 27
        Unload Me
    Case Else
        TrataCombo strcbCliente, cbMecanico, KeyAscii
        '2.6.9 Passa a n�o ser necess�rio informar valor para o recibo
'        VeSePodeHabilitar
End Select
End Sub

Private Sub cbMecanico_LostFocus()
strcbCliente = ""
End Sub

Private Sub Command1_Click()
'3.6.5 Mais log para a opera��o de criar recibo
Loga "Vai consultar o mec�nico", lDBG

'4.3.7 Tela para recibo avulso
'4.3.0 Recibo avulso
'If Tipo = tpOutros Then
'    If txDet.Text = "" Then
'        msgboxL "� necess�rio explicar o motivo do recibo"
'        Exit Sub
'    End If
'    If Text1(1).Text = "" Then
'        msgboxL "� necess�rio informar o destinat�rio do recibo"
'        Exit Sub
'    End If
'Else
    If Consulta("Select count(*) From Mecanicos Where Nome = '" & cbMecanico.Text & "' and Ativo = True ") = 0 Then
        msgboxL "Mec�nico Inexistente"
        Exit Sub
    End If
'End If
''3.5.1 N�o excluir fisicamente Mec�nico
'If Consulta("Select count(*) From Mecanicos Where Nome = '" & cbMecanico.Text & "' and Ativo = True ") = 0 Then
''If Consulta("Select count(*) From Mecanicos Where Nome = '" & cbMecanico.Text & "'") = 0 Then
'
'    '2.7.2 Logar todas mensagens
'    msgboxL "Mec�nico Inexistente"
'Else

    If txValor.Text = "" Then
        msgboxL "� necess�rio informar o valor"
        txValor.SetFocus
    Else

        '2.6.8 Log para verificar o endere�o na impress�o do or�amento
        Loga "Acionado a impress�o do recibo com o mec�nico: " & cbMecanico.Text & " endere�o: " & Endereco
        
        '4.5.9 Gravar o recibo de adiantamento tamb�m na impress�o em folha inteira
        GravaRecibos
        
        '4.8.7 Mostrar os vales pagos na impress�o do pagamento
        '4.8.0 Compensa��o dos vales na opera��o de pagamento
'        If gTipo = tpPagamento Then
'            If Vale > 0 Then
'                PagaVales
'            End If
'        End If
        
        If INI.TpImpress = 0 Then
        
            '4.3.7 Tela para recibo avulso
'            '4.3.4 Recibo gr�fico para recibo avulso
'            If gTipo = tpOutros Then
'
'                '4.3.5 Campos de Identifica��o e Endere�o no recibo avulso
'                Recibo.ReciboAvulso txDet.Text, Valor, Text1(1).Text, txVale.Text, Text1(2).Text
'                'Recibo.ReciboAvulso txDet.Text, Valor, Text1(1).Text
'
'            Else
                                                
                Recibo.RecebeDados cbMecanico.Text, Endereco, Valor, Text1(1).Text, 0, Text1(2).Text, txDet.Text, Tipo, Now
                '4.8.2 Mostrar o termo vale transporte no recibo em folha cheia
                'Recibo.RecebeDados cbMecanico.Text, Endereco, Valor, Text1(1).Text, 0, Text1(2).Text, txDet.Text, Tipo
                '4.5.9 Impress�o da observa��o no Recibo em folha inteira
                'Recibo.RecebeDados cbMecanico.Text, Endereco, Valor, Text1(1).Text, Vale, Text1(2).Text, txDet.Text
                '2.6.6 Campos Folga e Vale no recibo do mec�nico
                'Recibo.RecebeDados cbMecanico.Text, Endereco, Valor, Text1(1).Text, Vale, Text1(2).Text
                
'            End If
            
            Recibo.Show
        Else
        
            '4.5.9 Gravar o recibo de adiantamento tamb�m na impress�o em folha inteira
            'GravaRecibos
                        
            '4.0.5 ReImpress�o individual dos vales
            Select Case gTipo
                Case tpAdiantamento, tpComissao 'Adiantamento, Comiss�o
                    cRecibo.ReciboFita cbMecanico.Text, Endereco, txDet.Text, Valor, Text1(2).Text, Text1(2).Visible, Text1(1).Text, nrMec, True
                Case tpValeTransp 'Vale Transporte
                    cRecibo.ReciboVT cbMecanico.Text, Endereco, Valor, Text1(1).Text
                Case tpPagamento  'Pagamento Mensal
                    cRecibo.ReciboPagamento cbMecanico.Text, Endereco, nrMec, Label1(2).Caption, Text1(1).Text, Valor, Vale, txDet.Text, Text1(2).Text, Text1(2).Visible
                    
                '4.3.7 Tela para recibo avulso
'                '4.3.0 Recibo avulso
'                Case tpOutros
'
'                    '4.3.5 Campos de Identifica��o e Endere�o no recibo avulso
'                    cRecibo.ReciboOutros txDet.Text, Valor, Text1(1).Text, txVale.Text, Text1(2).Text
                    'cRecibo.ReciboOutros txDet.Text, Valor, Text1(1).Text
                    
            End Select

'                Case 0 'Adiantamento
'                    ReciboFita
'                Case 1 'Comiss�o
'                    ReciboFita
'                Case 2 'Vale Transporte
'                    '3.4.4 Vale Transporte
'                    ReciboVT
'                Case 3  'Pagamento Mensal
'                    '3.4.7 Recibo de pagamento mensal
'                    ReciboPagamento
'            End Select
            
            'Unload Me
        End If
        
        '4.8.7 Mostrar os vales pagos na impress�o do pagamento
        If gTipo = tpPagamento Then
            If Vale > 0 Then
                PagaVales
            End If
        End If
        Unload Me
        
    End If
'End If
End Sub

'4.8.0 Compensa��o dos vales na opera��o de pagamento
Private Sub PagaVales()
Dim dVale As Double
Dim SQL   As String
Dim rsVales As Recordset

dVale = Vale
SQL = "SELECT ID, Valor "
SQL = SQL & " From Vales "
SQL = SQL & " WHERE IdOperador=" & nrMec
SQL = SQL & " and Pago=0 and Tipo=0 "
AbreTB rsVales, SQL, dbOpenDynaset
Do While dVale > 0
    If rsVales!Valor Then
        ExecSql "Update Vales Set Pago = " & DTSqld(Now) & " Where ID = " & rsVales!ID
        If rsVales.EOF Then
            dVale = 0
        Else
            dVale = dVale - rsVales!Valor
            rsVales.MoveNext
        End If
    End If
Loop
Vale = dVale
End Sub

'3.4.7 Grava��o dos recibos
Private Sub GravaRecibos()
Dim SQL$

'4.3.1 Grava��o da observa��o so recibo avulso
SQL$ = "Insert Into Vales (IdOperador, Data, Valor, Pago, Tipo, Obs"

'4.3.7 Tela para recibo avulso
'If Tipo = 4 Then
'    SQL$ = SQL$ & ", NomeAvulso"
'End If

'4.8.5 ReImpress�o do recibo com todos os campos
SQL$ = SQL$ & ", Periodo, txValor"

SQL$ = SQL$ & ") values ("
'3.4.9 Gravar as observa��es do vale
'SQL$ = "Insert Into Vales (IdOperador, Data, Valor, Pago, Tipo, Obs) values ("

SQL$ = SQL$ & nrMec & ","
SQL$ = SQL$ & DTSqls(Format(Now, "DD/MM/YYYY HH:MM:SS")) & ","
SQL$ = SQL$ & VlrSql(STR(Valor))
SQL$ = SQL$ & ",0, " & Tipo

'4.8.5 ReImpress�o do recibo com todos os campos
If Tipo = 0 Or Tipo = 3 Then
'4.3.7 Tela para recibo avulso
'If Tipo = 0 Then
'4.3.1 Grava��o da observa��o so recibo avulso
'If Tipo = 0 Or Tipo = 4 Then
'3.5.8 Observa��o para o recibo de adiantamento
'If Tipo = 0 Then
    'SQL$ = SQL$ & ",'" & txDet.Text & "')"
        
    '4.3.1 Grava��o da observa��o so recibo avulso
    SQL$ = SQL$ & "," & FA(txDet.Text)
    
    '4.3.7 Tela para recibo avulso
'    If Tipo = 4 Then
'        SQL$ = SQL$ & "," & FA(Text1(1).Text)
'    End If

    '4.8.5 ReImpress�o do recibo com todos os campos
    'SQL$ = SQL$ & ")"
    
Else

    '3.4.9 Gravar as observa��es do vale
    SQL$ = SQL$ & "," & FA(Text1(2).Text)
End If

'4.8.5 ReImpress�o do recibo com todos os campos
SQL$ = SQL$ & "," & FA(Text1(1).Text) & ","
SQL$ = SQL$ & FA(Text1(2).Text) & ")"

ExecSql SQL$
End Sub

'Private Sub ReciboPagamento()
''3.4.7 Recibo de pagamento mensal
'Dim sValor!, SQL$, Aux$
'Dim tbVales  As Recordset
'Dim Data As String
'
'Const TamFita = 55
'
'ImprBuferizada_Inicializa
''Recibo
'LPRINT CENTRAL("RECIBO de Pagamento", TamFita / 2)
'
''Empresa
'LPRINT CENTRAL(INI.Empresa, TamFita / 2)
'
''Telefone
'LPRINT CENTRAL(INI.Fones, TamFita / 2)
'
''Endere�o
'LPRINT CENTRAL(INI.Endereco, TamFita / 2)
'
''2.7.4-5 Campo de CGC na configuracao
'If SN(INI.CGC > "", vbString) Then
'    LPRINT CENTRAL("CNPj: " & INI.CGC, TamFita / 2)
'End If
'
''--------
'LPRINT String(TamFita, "-")
'
''FUNC: MARCELO
''CARGO: Mecanico
''M�S: JANEIRO 2013
''         salario 1000#
''            vale 10.00  10-01-13
''            vale   50.00  15-01-13
''            vale   30.00   20-01-13
''
''valor a receber  910.00
'
'LPRINT "Nome: " & cbMecanico.Text
'LPRINT "Nr do RG: " & RG()
'LPRINT "Endere�o: " & Endereco
'
''3.5.1 N�o excluir fisicamente Mec�nico
'If Consulta("Select Oper From Mecanicos Where codi = " & nrMec & " and Ativo = True") = 0 Then
''If Consulta("Select Oper From Mecanicos Where codi = " & nrMec) = 0 Then
'
'    LPRINT "Cargo: Mecanico "
'Else
'    LPRINT "Cargo: Balconista "
'End If
'
''3.7.6 Mec�nico com recebimento semanal
'If Label1(2) = "Semana" Then
'    LPRINT "Semana de " & Text1(1).Text
'Else
'
'    LPRINT Text1(1).Text
'End If
'
''Valor
'
'LPRINT "Salario: " & Format(Valor, "##,##0.00")
'sValor! = Valor
'
'If Vale Then
'
'    '3.5.6 Conserto da impress�o dos vales no pagamento mensal
'    SQL$ = "Select * From Vales Where idOperador = " & nrMec & " and Pago = 0 and tipo = 0 Order By Data "
'
'    'SQL$ = "Select * From Vales Where idOperador = " & nrMec & " and Pago = 0 Order By Data "
'
'    LPRINT "Vales:"
'
'    AbreTB tbVales, SQL, dbOpenDynaset
'    Do While tbVales.EOF = False
'        Aux$ = Format(tbVales!Valor, "##,##0.00")
'
'        '3.8.2 Diminui��o da margem esquerda da impress�o do vale possibilitando maior espa�o para as observa��es
'        Aux$ = "Vale: " & ComplStr(Aux$, 8, " ", 2) & " " & Format(tbVales!Data, "DD/MM/YYYY") & tbVales!Obs
'        '3.8.0 Impress�o da observa��o na impress�o dos vales no pagamento
'        'Aux$ = Space(11) & "Vale: " & ComplStr(Aux$, 8, " ", 2) & " " & Format(tbVales!Data, "DD/MM/YYYY") & tbVales!Obs
'
'        LPRINT Left(Aux$, 54)
'        'LPRINT Space(11) & "Vale: " & ComplStr(Aux$, 8, " ", 2) & " " & Format(tbVales!Data, "DD/MM/YYYY")
'
'        tbVales.MoveNext
'    Loop
'    LPRINT "             ---------"
'    LPRINT "Soma dos vales:    " & Format(Vale, "##,##0.00")
'    sValor! = sValor! - Vale
'
'    '3.6.0 Impress�o da observa��o no recibo mensal
'    'LPRINT " "
'
'End If
'
''3.6.0 Impress�o da observa��o no recibo mensal
'If txDet.Text > "" Then
'    LPRINT " "
'    LPRINT txDet.Text
'    If Asc(Right(txDet.Text, 1)) <> 10 Then
'        LPRINT " "
'    End If
'End If
'
'LPRINT "Valor a receber: " & Format(sValor!, "##,##0.00")
'
''Extenso
'LPRINT "(" & Extenso(sValor!) & ")"
'
''Folga
''
'If Text1(2).Text > "" And Text1(2).Visible = True Then
'    LPRINT Text1(2).Text
'End If
'
'LPRINT "Concordo com o valor acima citado"
'
'LPRINT " "
''----Porto Alegre, 21 de Outubro de 2012
'LPRINT "Porto Alegre, " & Day(Now) & " de " & MesExtenso & " de " & Year(Now)
'
''2.7.4-5 Linha para impress�o no recibo em fita
'LPRINT " "
'LPRINT " "
'LPRINT String(TamFita, "-")
'
''3.6.6 tratamento para pc da impressora desligado
''SQL$ = "Update Vales Set Pago = "
''SQL$ = SQL$ & DTSqld(Now) & " Where idOperador = "
''SQL$ = SQL$ & nrMec & " and Pago = 0 and Tipo = 0 "
''ExecSql SQL$
'
''3.6.6 tratamento para pc da impressora desligado
'If ImprBuferizada_Finaliza = False Then
'    Exit Sub
'End If
''ImprBuferizada_Finaliza
'
''3.6.6 tratamento para pc da impressora desligado
'SQL$ = "Update Vales Set Pago = "
'SQL$ = SQL$ & DTSqld(Now) & " Where idOperador = "
'SQL$ = SQL$ & nrMec & " and Pago = 0 and Tipo = 0 "
'ExecSql SQL$
'End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

'4.3.0 Recibo avulso
Private Sub Form_Activate()
If JaAtivou = False Then
    JaAtivou = True
    
    '4.3.7 Tela para recibo avulso
'    If Tipo = tpOutros And Text1(1).Text = "" Then
'        Text1(1).SetFocus
'    Else
        If Mecanico > "" Then
            txValor.SetFocus
        End If
'    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Function CompoemSemana() As String
Dim dMenos   As Integer
Dim dSegunda As Date
Dim Hoje     As Date

Select Case Weekday(Now)
    Case 1
        dMenos = 6
    Case 3  'Ter�a
        dMenos = 1
    Case 4  'Quarta
        dMenos = 2
    Case 5  'Quinta
        dMenos = 3
    Case 6  'Sexta
        dMenos = 4
    Case 7  'Sabado
        dMenos = 5
End Select
dSegunda = Now - dMenos
CompoemSemana = Format(dSegunda, "dd/mm/yyyy") & " a " & Format(dSegunda + 5, "dd/mm/yyyy")
End Function

Public Property Let Endereco(ByVal vNewValue As String)
lcEndereco = vNewValue
End Property

Public Property Get Mecanico() As String
Mecanico = lcMecanico
End Property

Public Property Let Mecanico(ByVal vNewValue As String)
lcMecanico = vNewValue
cbMecanico.Text = vNewValue
End Property

Public Property Get Endereco() As String
If lcEndereco = "" Then
    '3.5.1 N�o excluir fisicamente Mec�nico
    Endereco = Consulta("Select Ende From Mecanicos Where Nome = '" & cbMecanico.Text & "' and Ativo = True ")
    'Endereco = Consulta("Select Ende From Mecanicos Where Nome = '" & cbMecanico.Text & "'")
Else
    Endereco = lcEndereco
End If
End Property

Private Sub Form_Load()
'4.3.0 Recibo avulso
JaAtivou = False

'4.0.5 ReImpress�o individual dos vales
Set cRecibo = New clsRecibo
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Public Property Get Valor() As Double
Valor = lcValor
End Property

Public Property Let Valor(ByVal vNewValue As Double)
lcValor = vNewValue
End Property

Public Property Get Vale() As Double
'2.6.6 Campos Folga e Vale no recibo do mec�nico
Vale = lcVale
End Property

Public Property Let Vale(ByVal vNewValue As Double)
'2.6.6 Campos Folga e Vale no recibo do mec�nico
lcVale = vNewValue
End Property

'Private Sub ReciboFita()
''2.7.4 Impress�o do Recibo em Matricial
'Dim sValor   As Single
''Dim SQL      As String
'
''2.9.0 Mostrar os carros das tarefas dos mec�nicos, previamente no recibo
''Dim rsComiss As Recordset '
'
'Const TamFita = 55
'
'ImprBuferizada_Inicializa
'
''Recibo
'LPRINT CENTRAL("RECIBO", TamFita / 2)
'
''Empresa
'LPRINT CENTRAL(INI.Empresa, TamFita / 2)
'
''Telefone
'LPRINT CENTRAL(INI.Fones, TamFita / 2)
'
''Endere�o
'LPRINT CENTRAL(INI.Endereco, TamFita / 2)
'
''2.7.4-5 Campo de CGC na configuracao
'If SN(INI.CGC > "", vbString) Then
'    LPRINT CENTRAL("CNPj: " & INI.CGC, TamFita / 2)
'End If
'
''--------
'LPRINT String(TamFita, "-")
'
''Nome
'LPRINT "Nome: " & cbMecanico.Text
'
''3.4.7 RG em todos os recibos
'LPRINT "Nr do RG: " & RG()
'
''Endere�o do Mec�nico
'LPRINT "Endere�o: " & Endereco
'
''3.6.0 Conserto da impress�o da comiss�o, em que estava imprimindo duas vezes
'If gTipo = tpAdiantamento Then
'    '3.5.8 Observa��o para o recibo de adiantamento
'
'    If txDet.Text > "" Then
'        LPRINT " "
'        LPRINT txDet.Text
'        If Asc(Right(txDet.Text, 1)) <> 10 Then
'            LPRINT " "
'        End If
'    End If
'End If
'
''Valor
'If Valor > 0 Then
'    LPRINT "Valor: " & Format(Valor, "##,###,###,##0.00")
'    sValor = Valor
'End If
'
'''Vale
''If Vale > 0 Then
''    LPRINT "Vale: " & Format(Vale, "##,###,###,##0.00")
''    sValor = Vale
''End If
'
''Extenso
'LPRINT "(" & Extenso(sValor) & ")"
'
''Folga
''
'If Text1(2).Text > "" And Text1(2).Visible = True Then
'    LPRINT Text1(2).Text
'End If
'
'LPRINT " "
'LPRINT "Concordo com o valor acima citado"
''LPRINT " "
'
''Semana
'If gTipo = tpComissao Then
'
'    LPRINT "Semana de " & Text1(1).Text
'    'LPRINT "Referente a semana de " & Text1(1).Text
'
'    '3.6.0 Otimiza��o da programa��o relativa a impress�o dos carros no recibo
'    If txDet.Text > "" Then
'        LPRINT " "
'        LPRINT "Referente aos servi�os nos carros: "
'        LPRINT " "
'        LPRINT txDet.Text
'        If Asc(Right(txDet.Text, 1)) <> 10 Then
'            LPRINT " "
'        End If
'    End If
''    If rsComiss.EOF = False Then
''        'Carros
''        LPRINT "Referente aos servi�os nos carros: "
''
''        Do While rsComiss.EOF = False
''            LPRINT rsComiss!Modelo & " " & rsComiss!Cor & " " & rsComiss!Carro & " "
''
''            '3.3.2 Impedir a possibilidade de continuar valores de comiss�o ap�s a impress�o do recibo
''            'ExecSql "Update Tarefas Set Pago = Int(Now) Where ID = " & rsComiss!ID
''
''            rsComiss.MoveNext
''        Loop
''
''    End If
'
'End If
'
''3.6.6 tratamento para pc da impressora desligado
''3.4.8 Adiantamento n�o deve zerar as comiss�es
''If gTipo = tpComissao Then
''    ExecSql "Update Tarefas Set Pago = Int(Now) Where Mec = " & nrMec & " and Pago is null and Situacao = 3 "
''End If
'
'LPRINT " "
''----Porto Alegre, 21 de Outubro de 2012
'LPRINT "Porto Alegre, " & Day(Now) & " de " & MesExtenso & " de " & Year(Now)
'
''2.7.4-5 Linha para impress�o no recibo em fita
'LPRINT " "
'LPRINT " "
'LPRINT String(TamFita, "-")
'
''3.6.7 Deixar de imprimir duas vezes o recibo
''3.6.6 tratamento para pc da impressora desligado
''If ImprBuferizada_Finaliza = False Then
''    Exit Sub
''End If
'
''3.6.6 tratamento para pc da impressora desligado
'If ImprBuferizada_Finaliza = False Then
'    Exit Sub
'End If
''ImprBuferizada_Finaliza
'
'If gTipo = tpComissao Then
'    ExecSql "Update Tarefas Set Pago = Int(Now) Where Mec = " & nrMec & " and Pago is null and Situacao = 3 "
'End If
'End Sub

Public Property Get Tipo() As Integer
'3.4.4 Vale Transporte
Tipo = gTipo
End Property

'3.4.4 Vale Transporte
Public Property Let Tipo(ByVal vNewValue As Integer)
Dim SQL   As String
Dim TbMec As Recordset
'Dim Grande As Boolean

gTipo = vNewValue
Select Case gTipo
    Case 0 'Adiantamento
        Me.Height = 2610
        Caption = "Recibo de adiantamento"
        
        '3.5.8 Observa��o para o recibo de adiantamento
        txDet.Top = 1200
        txDet.Visible = True
        txDet.Locked = False
        Label1(3).Caption = "Observa��o"
        Command1.Top = 2500
        frmRecibo.Height = 3500
        'frmRecibo.Height = 2400
        'Text1(2).Visible = True
        
        Label1(3).Visible = True
        Text1(2).Locked = False
        'Command1.Top = Command1.Top - 300
        lbVale.Visible = True
        txVale.Visible = True
        txVale.Locked = True
                
    Case 1 'Comiss�o
        Caption = "Recibo de Comiss�o"
        Text1(1).Text = CompoemSemana
        '2.9.0 Mostrar os carros das tarefas dos mec�nicos, previamente no recibo
        Me.Height = 3825
        txDet.Visible = True
        
        txDet.Visible = True
        Label1(2).Visible = True
        Label1(3).Visible = True
        Text1(1).Visible = True
        Text1(2).Visible = True
        
        '3.5.3 Conserto do SQL das comiss�es (3.5.1)
        SQL = " and Oper = 0 "
        'SQL$ = " Where Oper = 0 "
        
    Case 2 'Vale Transporte
        Caption = "Recibo de Vale Transporte"
        Text1(1).Top = Text1(2).Top
        Label1(2).Top = Label1(3).Top
        
        '3.4.6 Informa��o do mes do vale transporte
        Label1(2).Visible = True
        Text1(1).Visible = True
        Label1(2).Caption = "Mes"
        Text1(1).Text = "Referente ao mes de " & Mes()
        
        frmRecibo.Height = 2400
        Command1.Top = Command1.Top - 300
    Case 3  'Pagamento Mensal
        '3.4.7 Recibo de pagamento mensal
        Caption = "Recibo de Mensal"
    
        '3.5.0 Mostrar o valor a receber no recibo mensal
        Text1(2).Visible = True
        Label1(3).Caption = "Vlr a Receber"
        Label1(3).Visible = True
        Text1(2).Width = txValor.Width
        Text1(2).Locked = True
        'Text1(1).Top = Text1(2).Top
        'Label1(2).Top = Label1(3).Top
        
        Label1(2).Visible = True
        Text1(1).Visible = True
        Label1(2).Caption = "Mes"
        Text1(1).Text = "Referente ao mes de " & Mes()
        lbVale.Visible = True
        txVale.Visible = True
        txVale.Locked = True
        
        '3.6.0 Impress�o da observa��o no recibo mensal
        txDet.Top = 2100
        txDet.Visible = True
        txDet.Locked = False
        frmRecibo.Height = 3800
        'frmRecibo.Height = 2800
        
    '4.3.7 Tela para recibo avulso
    '4.3.0 Recibo avulso
'    Case 4
'        Caption = "Recibo para outros"
'        Text1(1).Top = cbMecanico.Top
'        cbMecanico.Visible = False
'        Text1(1).Visible = True
'        Label1(0).Caption = "Nome"
'        txDet.Visible = True
'        txDet.Locked = False
'        Command1.Top = Text1(2).Top + 80
'        Text1(1).TabIndex = 1
'
'        '4.3.5 Campos de Identifica��o e Endere�o no recibo avulso
''        txDet.Top = Command1.Top
''        frmRecibo.Height = 3500
'        txVale.Top = Text1(2).Top - 50
'        txVale.Left = Text1(2).Left
'        txVale.Width = txVale.Width * 2
'        txVale.Visible = True
'        Text1(2).Top = Text1(2).Top + 300
'        Text1(2).Visible = True
'        Label1(2).Top = Label1(2).Top - 50
'        Label1(2).Caption = "Endere�o:"
'        Label1(2).Visible = True
'        lbVale.Caption = "Identifica��o:"
'        lbVale.Top = txVale.Top + 50
'        lbVale.Left = Label1(2).Left - 200
'        lbVale.Visible = True
'        txDet.Top = 2100
'        Command1.Top = Command1.Top + 600
        
End Select

'3.7.6 Impedir nome de mec�nico em branco no recibo
SQL = "Select Nome From Mecanicos Where Ativo = True and Nome > '' " & SQL & " Order by Nome "
'3.5.3 Conserto do SQL das comiss�es (3.5.1)
'SQL$ = "Select Nome From Mecanicos Where Ativo = True " & SQL$ & " Order by Nome "

AbreTB TbMec, SQL, dbOpenDynaset
Do While TbMec.EOF = False
    cbMecanico.AddItem TbMec.Fields("Nome")
    TbMec.MoveNext
Loop
TbMec.Close
End Property

'3.4.6 Informa��o do mes do vale transporte
Private Function Mes() As String
Dim sMes$

sMes$ = Mid(Format(Now, "DD/MMMM"), 4)
sMes$ = Left(Chr(Asc(sMes$) - 32), 1) & Mid$(sMes$, 2)
sMes$ = sMes$ & " de " & Year(Now)
Mes = sMes$
End Function

'3.4.4 Vale Transporte
'Private Sub ReciboVT()
'Dim sValor   As Single
'Dim SQL      As String
'
'Const TamFita = 55
'
'ImprBuferizada_Inicializa
'
''3.5.8 Troca de posi��o do T�tulo do recibo do vale transporte pelo nome da empresa
'LPRINT CENTRAL("RECIBO DE VALE TRANSPORTE", TamFita / 2)
'
'LPRINT CENTRAL(INI.Empresa, TamFita / 2)
'LPRINT CENTRAL(INI.Fones, TamFita / 2)
'LPRINT CENTRAL(INI.Endereco, TamFita / 2)
'If SN(INI.CGC > "", vbString) Then
'    LPRINT CENTRAL("CNPj: " & INI.CGC, TamFita / 2)
'End If
'
''3.5.8 Troca de posi��o do T�tulo do recibo do vale transporte pelo nome da empresa
'LPRINT String(TamFita, "-")
'LPRINT CENTRAL("RECIBO DE VALE TRANSPORTE", TamFita / 2)
'
''--------
'LPRINT String(TamFita, "-")
'
'LPRINT "Nome: " & cbMecanico.Text
'LPRINT "Nr do RG: " & RG()
'LPRINT "Endere�o: " & Endereco
'If Valor > 0 Then
'    LPRINT "Valor: " & Format(Valor, "##,###,###,##0.00")
'    sValor = Valor
'End If
'LPRINT "(" & Extenso(sValor) & ")"
'
''3.4.6 Informa��o do mes do vale transporte
'If Text1(1).Text > "" Then
'    LPRINT " "
'    LPRINT Text1(1).Text
'    LPRINT " "
'End If
'
'LPRINT "Concordo com o valor acima citado"
'
'LPRINT " "
''----Porto Alegre, 21 de Outubro de 2012
'LPRINT "Porto Alegre, " & Day(Now) & " de " & MesExtenso & " de " & Year(Now)
'
''2.7.4-5 Linha para impress�o no recibo em fita
'LPRINT " "
'LPRINT " "
'LPRINT String(TamFita, "-")
'
''3.6.6 tratamento para pc da impressora desligado
'If ImprBuferizada_Finaliza = False Then
'    Exit Sub
'End If
''ImprBuferizada_Finaliza
'
'End Sub

'4.0.5 ReImpress�o individual dos vales
'Private Function RG() As String
''3.5.1 N�o excluir fisicamente Mec�nico
'RG = SN(Consulta("Select RG from Mecanicos Where Nome = '" & cbMecanico.Text & "' and Ativo = True "), vbString)
''RG = SN(Consulta("Select RG from Mecanicos Where Nome = '" & cbMecanico.Text & "'"), vbString)
'End Function

Private Function VeVales()
Dim SQL$

SQL$ = "SELECT Sum(Valor) AS Soma from Vales WHERE Vales.IdOperador="
SQL$ = SQL$ & nrMec & " and Pago=0 and Tipo=0 "
Vale = Consulta(SQL$)
txVale.Text = Format(Vale, "##,##0.00")
End Function

Private Sub txDet_KeyUp(KeyCode As Integer, Shift As Integer)
'3.5.8 Observa��o para o recibo de adiantamento
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub txValor_Change()
Dim xValor#, APAgar#

VeValor txValor.Text, xValor#, txValor, 0
Valor = xValor#

'3.5.0 Mostrar o valor a receber no recibo mensal
If gTipo = tpPagamento Then
    APAgar = xValor# - Vale
    If APAgar > 0 Then
        Text1(2).Text = Format(APAgar, "##,##0.00")
    Else
        Text1(2).Text = ""
    End If
End If
End Sub
