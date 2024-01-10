VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Begin VB.Form Forneced 
   Appearance      =   0  'Flat
   BackColor       =   &H0000FF00&
   Caption         =   "Fornecedores"
   ClientHeight    =   5925
   ClientLeft      =   1350
   ClientTop       =   1350
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   24.688
   ScaleMode       =   4  'Character
   ScaleWidth      =   56.125
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btBoleto 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&Boleto"
      Default         =   -1  'True
      Height          =   360
      Left            =   4071
      TabIndex        =   39
      Top             =   5160
      Width           =   1245
   End
   Begin VB.TextBox txEmail 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1620
      TabIndex        =   37
      Top             =   4320
      Width           =   5040
   End
   Begin VB.TextBox txForn 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   7
      Left            =   5280
      MaxLength       =   9
      TabIndex        =   9
      ToolTipText     =   "Digite aqui o CEP"
      Top             =   960
      Width           =   1380
   End
   Begin VB.TextBox txForn 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   6
      Left            =   1620
      MaxLength       =   9
      TabIndex        =   7
      ToolTipText     =   "Digite aqui o CEP"
      Top             =   960
      Width           =   2940
   End
   Begin VB.TextBox txNroBco 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5280
      MaxLength       =   9
      TabIndex        =   5
      ToolTipText     =   "Digite aqui o CEP"
      Top             =   540
      Width           =   1380
   End
   Begin VB.ComboBox cbBanco 
      Height          =   360
      Left            =   1620
      TabIndex        =   3
      Top             =   540
      Width           =   2955
   End
   Begin VB.CommandButton btPedido 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   360
      Left            =   5400
      TabIndex        =   36
      Top             =   5160
      Width           =   1245
   End
   Begin VB.TextBox txForn 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   4
      Left            =   5340
      MaxLength       =   9
      TabIndex        =   26
      ToolTipText     =   "Digite aqui o CEP"
      Top             =   3480
      Width           =   1320
   End
   Begin VB.TextBox txForn 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   1
      Left            =   1620
      MaxLength       =   30
      TabIndex        =   13
      ToolTipText     =   "Digite a Cidade da Empresa"
      Top             =   1800
      Width           =   5040
   End
   Begin VB.TextBox txForn 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   2
      Left            =   1620
      MaxLength       =   30
      TabIndex        =   15
      ToolTipText     =   "Digite o Bairro da Empresa"
      Top             =   2250
      Width           =   5040
   End
   Begin VB.TextBox txForn 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   3
      Left            =   1620
      MaxLength       =   16
      TabIndex        =   19
      ToolTipText     =   "Digite o CGC da Empresa"
      Top             =   3090
      Width           =   2175
   End
   Begin VB.TextBox txForn 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   0
      Left            =   1620
      MaxLength       =   30
      TabIndex        =   11
      ToolTipText     =   "Digite a Razão Social da empresa"
      Top             =   1380
      Width           =   5040
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&Recibo"
      Height          =   360
      Left            =   2744
      TabIndex        =   35
      Top             =   5160
      Width           =   1245
   End
   Begin VB.CommandButton btAdiciona 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Adiciona"
      Height          =   360
      Left            =   90
      TabIndex        =   28
      Top             =   5160
      Width           =   1245
   End
   Begin VB.CommandButton Apaga 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&Apagar"
      Height          =   360
      Left            =   1417
      TabIndex        =   29
      Top             =   5160
      Width           =   1245
   End
   Begin VB.TextBox Lugar 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5340
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4740
      Width           =   1320
   End
   Begin VB.TextBox UltForn 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2700
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4740
      Width           =   1200
   End
   Begin VB.TextBox Vend 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1620
      MaxLength       =   30
      TabIndex        =   27
      Top             =   3900
      Width           =   5040
   End
   Begin VB.TextBox Fax 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   23
      Top             =   3480
      Width           =   1200
   End
   Begin VB.TextBox Fone 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1620
      MaxLength       =   10
      TabIndex        =   21
      Top             =   3480
      Width           =   1200
   End
   Begin VB.TextBox Ender 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1620
      MaxLength       =   30
      TabIndex        =   17
      Top             =   2700
      Width           =   5040
   End
   Begin VB.TextBox EdtFornec 
      BackColor       =   &H0000FF00&
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1620
      MaxLength       =   30
      TabIndex        =   1
      Top             =   120
      Width           =   5040
   End
   Begin MSMAPI.MAPIMessages MAPIMessages 
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   0   'False
      NewSession      =   0   'False
      Password        =   "ufrt3753"
      UserName        =   "xeviousbr@gmail.com"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&Email:"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   16
      Left            =   465
      TabIndex        =   38
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "&Conta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   15
      Left            =   4680
      TabIndex        =   8
      Top             =   1020
      Width           =   510
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "&Nr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   14
      Left            =   4980
      TabIndex        =   4
      Top             =   660
      Width           =   210
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&Agência"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   13
      Left            =   465
      TabIndex        =   6
      Top             =   1020
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&Banco"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   12
      Left            =   465
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "CEP:"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   11
      Left            =   4740
      TabIndex        =   24
      Top             =   3600
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&Cidade:"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   10
      Left            =   465
      TabIndex        =   12
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&Bairro:"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   9
      Left            =   465
      TabIndex        =   14
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "CNPj/CPF"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   8
      Left            =   105
      TabIndex        =   18
      Top             =   3180
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Razão Social"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   7
      Left            =   465
      TabIndex        =   10
      Top             =   1500
      Width           =   1095
   End
   Begin VB.Label Men 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   60
      TabIndex        =   34
      Top             =   5580
      Width           =   6480
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "Número"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   4260
      TabIndex        =   33
      Top             =   4860
      Width           =   660
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "Último Recibo"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   5
      Left            =   1080
      TabIndex        =   31
      Top             =   4860
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&Vendedor:"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   4
      Left            =   465
      TabIndex        =   25
      Top             =   4020
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "Fa&x:"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   2880
      TabIndex        =   22
      Top             =   3600
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "F&one:"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   465
      TabIndex        =   20
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&Endereço:"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   465
      TabIndex        =   16
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "&Nome:"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   465
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Forneced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'5.0.3 Impressão para matricial e usb do boleto
'4.6.3 Conserto do boleto em branco
'4.6.2 Envio de Email no Boleto [parte 1]
'4.5.2 Relatório de Boletos com mais de uma parcela
'4.5.1 Campos DOC, Obs e CPF no Boleto
'4.5.0 Relatório Boleto apartir dos fornecedores
'4.4.9 Relatório Boleto
'4.4.2 Cadastro de fornecedores

Option Explicit

Private VezMensEmail    As Integer
Private PosForn         As Integer
Private l_Mudado        As Boolean
Private Redimencionou   As Boolean
Private TemQueOrdenar   As Boolean
Private Adicionando     As Boolean
Private CadFornVazio    As Boolean
Private Mudando         As Boolean
Private BancoNovo       As Boolean
Private Achado()        As Long
Private MaxCodiFor      As String
Private SQL             As String
Private Bco(100)        As String
Private TbFor           As Recordset

'4.5.0 Relatório Boleto apartir dos fornecedores
Private Sub btBoleto_Click()
Dim envEmail    As Boolean
Dim OK          As Boolean
Dim a           As Integer
Dim idPagtoForn As Long
Dim Valor       As Currency
Dim sValor      As String
Dim DOC         As String
Dim Obs         As String
Dim rs          As Recordset

Load PagtoForn
PagtoForn.Fornec = TbFor!codi
PagtoForn.Show 1
OK = PagtoForn.OK
Valor = PagtoForn.Valor
sValor = VlrSql(Trim(STR(Valor)))
idPagtoForn = PagtoForn.idPagtoForn

'4.5.1 Campos DOC, Obs e CPF no Boleto
DOC = PagtoForn.DOC
Obs = PagtoForn.Obs

envEmail = PagtoForn.ckEmail

Unload PagtoForn

Dim X As Long
X = DoEvents()

If OK Then

    SQL = "SELECT Data, Valor, Auto "
    SQL = SQL & "From ParcelasTemp "
    SQL = SQL & "WHERE IDPC = " & INI.PC
    SQL = SQL & " Order by Auto "
    AbreTB rs, SQL, dbOpenSnapshot
    rs.MoveFirst

    a = 1
    Do While rs.EOF = False

        SQL = "Insert Into ParcelasPagtoForn (idPagtoForn, Nr, Valor, Data, Pago) Values ("
        SQL = SQL & idPagtoForn & ", "
        SQL = SQL & STR(a) & " , "
        SQL = SQL & VlrSql(rs!Valor) & ", "
        SQL = SQL & DTSqld(rs!Data) & ", "
        SQL = SQL & "0) "
        ExecSql SQL
        rs.MoveNext
        a = a + 1
    Loop
    
    SQL = "Insert Into PagtoForn (idForn, Valor, Data, DOC, Obs, QtdParc) Values ("
    SQL = SQL & TbFor("codi") & ", "
    SQL = SQL & sValor & ", "
    SQL = SQL & DTSqld(Now) & ", "
    SQL = SQL & FA(DOC) & ", "
    SQL = SQL & FA(Obs) & ", "
    SQL = SQL & (a - 1) & ")"
    ExecSql SQL
    
    '4.5.2 Relatório de Boletos com mais de uma parcela
    SQL = "SELECT PagtoForn.idForn, PagtoForn.Valor, PagtoForn.Data, PagtoForn.DOC, PagtoForn.Obs, Fornecedores.Nome, Fornecedores.Ender, "
    SQL = SQL & "Fornecedores.Telefone, Fornecedores.Vendedor, Fornecedores.CGC, Fornecedores.Cid, Fornecedores.CEP, Bancos.Nome as Banco, Fornecedores.Age, "
    SQL = SQL & "Fornecedores.Conta, Fornecedores.email, PagtoForn.idPagtoForn, "
    SQL = SQL & "Str(ParcelasPagtoForn.Nr)+'/'+Str(PagtoForn.QtdParc) AS Nr, "
    SQL = SQL & "ParcelasPagtoForn.Valor AS ValorParc, ParcelasPagtoForn.Data AS DataParc "
        
    '4.6.3 Conserto do boleto em branco
    SQL = SQL & "FROM (Bancos RIGHT JOIN Fornecedores ON Bancos.Nr = Fornecedores.idBanco) "
    'SQL = SQL & "FROM (Bancos INNER JOIN Fornecedores ON Bancos.Nr = Fornecedores.idBanco) "
    
    SQL = SQL & "INNER JOIN (PagtoForn INNER JOIN ParcelasPagtoForn ON PagtoForn.idPagtoForn = ParcelasPagtoForn.idPagtoForn) "
    SQL = SQL & "ON Fornecedores.codi = PagtoForn.idForn "
    SQL = SQL & "WHERE PagtoForn.idPagtoForn = " & idPagtoForn
    
    '5.0.3 Impressão para matricial e usb do boleto
    If INI.TpImpress = 2 Then
        ImprUsb SQL
        
    Else
        Load RelBoleto
        RelBoleto.Dados.RecordSource = SQL
        RelBoleto.Dados.DatabaseName = App.Path + "\OrCarro.mdb"
        RelBoleto.Show 1
    End If
    
    '4.6.2 Envio de Email no Boleto [parte 1]
    If envEmail Then
        If Trim(txEmail.Text) = "" Then
            OK = False
            Do While OK = False
                txEmail.Tag = "PediuEmail"
                txEmail.Text = InputBox("Informe o email do fornecedor")
                If Trim(txEmail.Text) = "" Then
                    VezMensEmail = VezMensEmail + 1
                    If VezMensEmail = 1 Then
                        MsgBox "Se cancelar o email agora não poderá enviar o email deste boleto", vbCritical
                    Else
                        VezMensEmail = 0
                        Exit Sub
                    End If
                Else
                    OK = CriticaEmail(txEmail.Text)
                    If OK = False Then
                        txEmail.SetFocus
                    Else
                        If txEmail.Tag = "PediuEmail" Then
                            ExecSql "Update Fornecedores Set email = " & FA(txEmail.Text)
                            txEmail.Tag = ""
                        End If
                    End If
                End If
            Loop
        End If
        
        EnviarEmail
        
    End If
End If
End Sub

'5.0.3 Impressão para matricial e usb do boleto
Private Sub ImprUsb(SQL As String)
Dim rsBoleto As Recordset

AbreTB rsBoleto, SQL, dbOpenSnapshot
ImprBuferizada_Inicializa
Do While rsBoleto.EOF = False
    LPRINT CENTRAL(SemAcento(INI.Empresa), TamFita / 2)
    LPRINT INI.Fones
    LPRINT "EMAIL:denisbonifacio@gmail.com"
    LPRINT "BOLETO SOMENTE PARA DEPOSITO EM CONTA"
    LPRINT "BANCO             : " & SN(rsBoleto!Banco, vbString)
    LPRINT "AGÊNCIA           : " & SN(rsBoleto!Age, vbString)
    LPRINT "CONTA             : " & SN(rsBoleto!Conta, vbString)
    LPRINT "DATA COMPRA       : " & Format(rsBoleto!Data, "DD/MM/YYYY")
    LPRINT "VALOR COMPRA      : " & Format(rsBoleto!Valor, "#,###,##0.00")
    LPRINT "Nr DO DOC         : " & SN(rsBoleto!DOC, vbString)
    LPRINT "EMPRESA           : " & rsBoleto!Nome
    LPRINT "CNPj/CPF          : " & SN(rsBoleto!CGC, vbString)
    LPRINT "ENDEREÇO          : " & SN(rsBoleto!Ender, vbString)
    LPRINT "CIDADE            : " & SN(rsBoleto!Cid, vbString)
    LPRINT "CEP               : " & SN(rsBoleto!CEP, vbString)
    LPRINT "E-MAIL            : " & SN(rsBoleto!email, vbString)
    LPRINT "VENDEDOR          : " & SN(rsBoleto!Vendedor, vbString)
    LPRINT "FONE              : " & SN(rsBoleto!Telefone, vbString)
    LPRINT "VALOR DAS PARCELAS: " & rsBoleto!ValorParc
    LPRINT "Nr DA PARCELA     : " & rsBoleto!Nr
    LPRINT "DATA VENCIMENTO   : " & rsBoleto!DataParc
    LPRINT "OBSERVAÇÃO        : " & rsBoleto!Obs
    LPRINT " "
    LPRINT " "
    rsBoleto.MoveNext
Loop
ImprBuferizada_Finaliza
End Sub

Private Sub EnviarEmail()
Dim clsEmail As New cEmail

clsEmail.objMAPIMessage = MAPIMessages
clsEmail.objMAPISession = MAPISession
clsEmail.nmCliente = EdtFornec.Text
clsEmail.enderMailCli = txEmail.Text
clsEmail.EnviaEmail
End Sub

Private Sub btPedido_Click()
Unload Me
End Sub

Private Sub btPedido_KeyDown(KeyCode As Integer, Shift As Integer)
Dim X As Integer

If KeyCode = 27 Then
    If Mudado Then GRAVA
    Forneced.Hide
    Exit Sub
End If
X = TRATATECLA(KeyCode)
End Sub

Private Sub cbBanco_Click()
If Mudando = False Then
    Mudado = True
    txNroBco.Text = cbBanco.ItemData(cbBanco.ListIndex)
End If
End Sub

Private Sub cbBanco_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    If Mudado Then GRAVA
    Forneced.Hide
    Exit Sub
Else
    If Mudando = False Then
        If Mudado = False Then
            Mudado = True
            txNroBco.Enabled = True
            txNroBco.Text = ""
            BancoNovo = True
        End If
    End If
End If
End Sub

Private Sub Command1_Click()
Load frmReciboAvulso
frmReciboAvulso.Fornec
frmReciboAvulso.txRecebe = EdtFornec.Text
frmReciboAvulso.txIdent = txForn(3).Text
frmReciboAvulso.txEndereco = Ender.Text
frmReciboAvulso.Show
End Sub

Private Sub Ender_GotFocus()
Seleciona
Men.Caption = "Digite o endereco com ate 30 letras"
End Sub

Private Sub Ender_KeyDown(KeyCode As Integer, Shift As Integer)
Mudado = Mudado + TRATATECLA(KeyCode)
End Sub

Private Sub Fax_GotFocus()
Seleciona
Men.Caption = "Digite o FAX com ate 10 letras"
End Sub

Private Sub Fax_KeyDown(KeyCode As Integer, Shift As Integer)
Mudado = Mudado + TRATATECLA(KeyCode)
End Sub

Private Sub Fone_GotFocus()
Seleciona
Men.Caption = "Digite o TELEFONE com ate 10 letras"
End Sub

Private Sub Fone_KeyDown(KeyCode As Integer, Shift As Integer)
Mudado = Mudado + TRATATECLA(KeyCode)
End Sub

Public Sub MostraLugar()
MaxCodiFor = Trim(STR(Consulta("Select Max(Codi) From Fornecedores")))
Lugar.Text = STR$(TbFor("codi")) + MaxCodiFor
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    SendKeys "{Tab}"
End If
End Sub

Private Sub Form_Load()
Dim a     As Integer
Dim TbBco As Recordset

Mudando = True

'4.6.2 Envio de Email no Boleto [parte 1]
VezMensEmail = 0
txEmail.Tag = ""

InicForm Me
AbreTB TbBco, "Select Nome, Nr From Bancos Order By Nome", dbOpenForwardOnly
Do While TbBco.EOF = False
    Bco(a) = TbBco!Nome
    TextCombo Bco(a), cbBanco, 1, Val(TbBco!Nr)
    TbBco.MoveNext
Loop
cbBanco.ListIndex = -1
Mudando = False
BancoNovo = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then Sai_Click
End Sub

Private Sub GRAVA()
If EdtFornec.Text > "" Then

    If cbBanco.ListIndex > -1 Then
        If Val(txNroBco.Text) = 0 Then
            MsgBox "É necessário informar o número do banco", vbExclamation, "Estoq"
            txNroBco.SetFocus
            Exit Sub
        End If
    End If

    If Adicionando = True Then
        Adicionando = False
        If TbFor Is Nothing Then
            AbreTB TbFor, "Fornecedores", dbOpenTable
        End If
        TbFor.AddNew
        TbFor("codi") = PosForn
        MaxCodiFor = "/" + Trim(STR(PosForn))
        Lugar.Text = STR$(TbFor("codi")) + MaxCodiFor
    Else
        On Local Error GoTo EdicaoInvalida
        TbFor.Edit
        On Local Error GoTo 0
    End If
    TbFor("Nome") = EdtFornec.Text
    TbFor("Ender") = Ender.Text
    TbFor("Telefone") = Fone.Text
    TbFor("Fax") = Fax.Text
    TbFor("Vendedor") = Vend.Text
    TbFor("RazaoDoc") = txForn(0).Text
    TbFor("Cid") = txForn(1).Text
    TbFor("Bai") = txForn(2).Text
    TbFor("CGC") = txForn(3).Text
    TbFor("CEP") = txForn(4).Text
    
    TbFor("Age") = txForn(6).Text
    TbFor("Conta") = txForn(7).Text
    
    TbFor("email") = txEmail.Text
    
    If BancoNovo Then
        ExecSql "Insert into Bancos (Nr, Nome) Values (" & FA(txNroBco.Text) & ", " & FA(cbBanco.Text) & ")"
        TbFor("idBanco") = Val(txNroBco.Text)
        BancoNovo = False
    Else
        If Val(cbBanco.Tag) <> cbBanco.ListIndex Then
            If cbBanco.ListIndex = -1 Then
                TbFor("idBanco") = Null
            Else
                TbFor("idBanco") = Val(txNroBco.Text)
            End If
        End If
    End If
    TbFor.Update
End If
Mudado = 0
Exit Sub
EdicaoInvalida:
End Sub

Private Sub Imprime_GotFocus()
Men.Caption = "Imprime uma ficha com os dados deste fornecedor"
End Sub

Private Sub Imprime_KeyDown(KeyCode As Integer, Shift As Integer)
Dim X As Integer

If KeyCode = 27 Then
    If Mudado Then GRAVA
    Forneced.Hide
    Exit Sub
End If
X = TRATATECLA(KeyCode)
End Sub

Public Sub Mostrar(Optional Forn As String = "")
Mudando = True
If Forn > "" Then
    AbreTB TbFor, "Select * From Fornecedores Where Nome = " & FA(Forn), dbOpenDynaset
End If
EdtFornec.Text = RTrim$(TbFor("Nome"))
Ender.Text = RTrim$(SN(TbFor("Ender"), vbString))
Fone.Text = RTrim$(SN(TbFor("Telefone"), vbString))
Fax.Text = RTrim$(SN(TbFor("Fax"), vbString))
Vend.Text = RTrim$(SN(TbFor("Vendedor"), vbString))
UltForn.Text = SN(TbFor("UltForn"), vbString)
Lugar.Text = STR$(PosForn) + "/" + Trim(STR$(TbFor.RecordCount))
txForn(0) = SN(TbFor("RazaoDoc"), vbString)
txForn(1).Text = SN(TbFor("Cid"), vbString)
txForn(2).Text = SN(TbFor("Bai"), vbString)
txForn(3).Text = SN(TbFor("CGC"), vbString)
txForn(4).Text = SN(TbFor("CEP"), vbString)
PosForn = Val(TbFor("Codi"))

txForn(6).Text = SN(TbFor("Age"), vbString)
txForn(7).Text = SN(TbFor("Conta"), vbString)

If Val(SN(TbFor("idBanco"))) = 0 Then
    cbBanco.ListIndex = -1
    txNroBco.Enabled = True
    cbBanco.Tag = "-1"
Else
    PoemBanco TbFor("idBanco")
    txNroBco.Enabled = False
End If

txEmail.Text = SN(TbFor("email"), vbString)

Mudado = 0
Mudando = False
End Sub

Private Sub PoemBanco(idBanco As Integer)
Dim rsBanco As Recordset

AbreTB rsBanco, "Select Nome From Bancos Where Nr = " & idBanco, dbOpenSnapshot
TextCombo rsBanco!Nome, cbBanco, 1, idBanco
cbBanco.Tag = STR(idBanco)
txNroBco.Text = idBanco
End Sub

Private Sub Sai_Click()
Unload Me
End Sub

Private Function TRATATECLA(Tecla As Integer, Optional NoNome As Boolean) As Integer
'Teclas a tratar ESC, Page-Up, Page-Down"
Dim a As Integer
TRATATECLA = 1
Select Case Tecla
    Case Is = 27
        If Mudado Then GRAVA
        Sai_Click
        TRATATECLA = 0
    Case Is = 33
        'Page-Up
        
        If TbFor.BOF = False Then
            TbFor.MovePrevious
            If TbFor.BOF = False Then
                Mostrar
            End If
        End If
        TRATATECLA = 0
    Case Is = 34
        'Page-Down
        If Mudado Then GRAVA
        If TbFor.EOF = False Then
            TbFor.MoveNext
            If TbFor.EOF = False Then
                Mostrar
            End If
        End If
        TRATATECLA = 0
    Case Else
        If CadFornVazio Then
            Adicionar
        Else
            btAdiciona.Caption = "&Grava"
        End If
End Select
End Function

Private Sub Imprime_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    If Mudado Then GRAVA
    Sai_Click
End If
End Sub

Private Sub txEmail_KeyDown(KeyCode As Integer, Shift As Integer)
Mudado = Mudado + TRATATECLA(KeyCode)
End Sub

Private Sub txEmail_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    If Mudado Then GRAVA
    Sai_Click
End If
End Sub

Private Sub txForn_GotFocus(Index As Integer)
Seleciona
Men.Caption = txForn(Index).ToolTipText
End Sub

Private Sub txForn_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Mudado = Mudado + TRATATECLA(KeyCode)
End Sub

Private Sub txForn_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub Vend_GotFocus()
Seleciona
Men.Caption = "Digite o nome do VENDEDOR ou CONTATO com ate 30 letras"
End Sub

Private Sub Vend_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And PosForn > 0 Then
    If btAdiciona.Caption = "&Grava" Then
        btAdiciona.SetFocus
    Else
        EdtFornec.SetFocus
    End If
End If
Mudado = Mudado + TRATATECLA(KeyCode)
End Sub

Private Sub EdtFornec_GotFocus()
Seleciona
Men.Caption = "Digite o FORNECEDOR tecle Page(Up/Down) ou TAB"
End Sub

Private Sub EdtFornec_KeyDown(KeyCode As Integer, Shift As Integer)
TRATATECLA KeyCode, True
End Sub

Private Sub EdtFornec_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub Ender_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub Vend_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub btAdiciona_Click()
If btAdiciona.Caption = "&Grava" Then
    GRAVA
    btAdiciona.Caption = "A&diciona"
    EdtFornec.SetFocus
Else
    Adicionar
End If
End Sub

Public Sub Adicionar()
Mudando = True
btAdiciona.Caption = "&Grava"
ZeraTela
PosForn = Consulta("Select Max(Codi) From Fornecedores") + 1
Lugar.Text = STR$(PosForn) + "/" + STR$(PosForn)
UltForn.Text = Format$(Now, "dd/mm/yyyy")
btBoleto.Enabled = False
Command1.Enabled = False
Apaga.Enabled = False
cbBanco.ListIndex = -1
Adicionando = True
Mudando = False
EdtFornec.SetFocus
End Sub

Private Sub btAdiciona_KeyDown(KeyCode As Integer, Shift As Integer)
TRATATECLA KeyCode
End Sub

Private Sub ZeraTela(Optional NaoApagaNome As Boolean = False)
If NaoApagaNome = False Then EdtFornec.Text = ""
Ender.Text = ""
Fone.Text = ""
Fax.Text = ""
Vend.Text = ""
txForn(0).Text = ""
txForn(1).Text = ""
txForn(2).Text = ""
txForn(3).Text = ""
txForn(4).Text = ""
End Sub

Private Property Get Mudado() As Boolean
Mudado = l_Mudado
End Property

Private Property Let Mudado(ByVal vNewValue As Boolean)
If Mudado = False And vNewValue = True Then
    btAdiciona.Caption = "&Grava"
End If
l_Mudado = vNewValue
End Property
