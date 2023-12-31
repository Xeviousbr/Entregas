VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmReciboAvulso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibo do Mecânico"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4350
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox txDet 
      Height          =   1095
      Left            =   60
      TabIndex        =   11
      Top             =   2520
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1931
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmReciboAvulso.frx":0000
   End
   Begin VB.TextBox txEndereco 
      Height          =   285
      Left            =   1245
      TabIndex        =   3
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox txQuemPaga 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1260
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txRecebe 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1260
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox txIdent 
      Height          =   285
      Left            =   1260
      TabIndex        =   2
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txValor 
      Height          =   285
      Left            =   1260
      TabIndex        =   4
      Top             =   1560
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   435
      Left            =   1620
      TabIndex        =   5
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Endereço:"
      Height          =   195
      Index           =   4
      Left            =   405
      TabIndex        =   10
      Top             =   1260
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Quem Recebe:"
      Height          =   195
      Index           =   3
      Left            =   75
      TabIndex        =   9
      Top             =   540
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Identificação:"
      Height          =   195
      Index           =   2
      Left            =   195
      TabIndex        =   8
      Top             =   900
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor:"
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   7
      Top             =   1620
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Quem Paga"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "frmReciboAvulso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'5.0.2 Utilizar a escolha da impressão somente se a impressora for definida como USB
'4.9.9 Ajustes na escolha da impressora
'4.9.8 Selecionar a impressora ao imprimir
'4.7.7 Troca do componente da Observação do recibo Avulso de Text para RichTextBox
'4.7.6 Pesquisar por cliente ou funcionário no recibo avulso
'4.7.0 Retirada da observação dupla no recibo avulso
'4.6.4 Impressão da observação no Recibo Avulso em folha inteira
'4.5.0 Relatório Boleto apartir dos fornecedores
'4.4.9 Relatório Boleto
'4.3.7 Tela para recibo avulso

Option Explicit

Private EhFornec     As Boolean
Private PercComiss   As Single
Private nrMec        As Integer
Private lcValor      As Double
Private lcVale       As Double
Private l_Fornecedor As String
Private rsComiss     As Recordset
Private cRecibo      As clsRecibo

'4.7.6 Pesquisar por cliente ou funcionário no recibo avulso
'Private Sub btPesquisar_Click()
'Load frmBusca
'frmBusca.Tipo = 3
'frmBusca.Show
'Unload Me
'End Sub

Private Sub Command1_Click()
If txDet.Text = "" Then
    msgboxL "É necessário explicar o motivo do recibo"
    Exit Sub
End If
If txRecebe.Text = "" Then
    msgboxL "É necessário informar o destinatário do recibo"
    Exit Sub
End If

If txValor.Text = "" Then
    msgboxL "É necessário informar o valor"
    txValor.SetFocus
    Exit Sub
Else

    '5.0.2 Utilizar a escolha da impressão somente se a impressora for definida como USB
    If INI.TpImpress = 2 Then

        '4.9.8 Selecionar a impressora ao imprimir
        Load frmConfigImpr
        
        '4.9.9 Ajustes na escolha da impressora
        frmConfigImpr.Caption = "Escolha a impressora"
        frmConfigImpr.Label2.Caption = ""
        
        frmConfigImpr.Show 1
        If frmConfigImpr.OK Then
            If frmConfigImpr.ImprFita Then
                Loga "Impressão em Fita", lDBG
                cRecibo.ReciboOutros txDet.Text, Valor, txRecebe.Text, txIdent.Text, txEndereco.Text, txQuemPaga.Text
            Else
                Loga "Impressão em Folha", lDBG
                ReciboAvulso.ReciboAvulso txDet.Text, Valor, txRecebe.Text, txIdent.Text, txEndereco.Text, txQuemPaga.Text
                ReciboAvulso.Show 1
            End If
        End If
        Unload frmConfigImpr
    Else
        If INI.TpImpress = 0 Then
            ReciboAvulso.ReciboAvulso txDet.Text, Valor, txRecebe.Text, txIdent.Text, txEndereco.Text, txQuemPaga.Text
            ReciboAvulso.Show 1
        Else
            cRecibo.ReciboOutros txDet.Text, Valor, txRecebe.Text, txIdent.Text, txEndereco.Text, txQuemPaga.Text
        End If
    End If
End If
Unload Me
End Sub

Private Sub GravaRecibos()
Dim SQL$

SQL$ = "Insert Into Vales (IdOperador, Data, Valor, Pago, Tipo, Obs"
SQL$ = SQL$ & ", NomeAvulso"
SQL$ = SQL$ & ") values ("
SQL$ = SQL$ & nrMec & ","
SQL$ = SQL$ & DTSqls(Format(Now, "DD/MM/YYYY HH:MM:SS")) & ","
SQL$ = SQL$ & VlrSql(STR(Valor))
SQL$ = SQL$ & ",0, " & 4
SQL$ = SQL$ & "," & FA(txDet.Text)
SQL$ = SQL$ & "," & FA(txRecebe.Text)
SQL$ = SQL$ & ")"
  
ExecSql SQL$
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
EhFornec = False
Set cRecibo = New clsRecibo
End Sub

'Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
'If KeyAscii = 27 Then
'    Unload Me
'End If
'End Sub

Public Property Get Valor() As Double
Valor = lcValor
End Property

Public Property Let Valor(ByVal vNewValue As Double)
lcValor = vNewValue
End Property

Public Property Get Vale() As Double
Vale = lcVale
End Property

Public Property Let Vale(ByVal vNewValue As Double)
lcVale = vNewValue
End Property

Private Sub txDet_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub txValor_Change()
Dim xValor#, APAgar#

VeValor txValor.Text, xValor#, txValor, 0
Valor = xValor#
If gTipo = tpPagamento Then
    APAgar = xValor# - Vale
    If APAgar > 0 Then
        txEndereco.Text = Format(APAgar, "##,##0.00")
    Else
        txEndereco.Text = ""
    End If
End If
End Sub

'Public Property Get Fornecedor() As String
'l_Fornecedor = l_Fornecedor
'End Property
'
'Public Property Let Fornecedor(ByVal vNewValue As String)
'l_Fornecedor = vNewValue
'
'End Property

'4.4.9 Relatório Boleto
Public Sub Fornec()
EhFornec = True
End Sub
