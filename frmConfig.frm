VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmConfig 
   Caption         =   "Configuração"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   5850
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txNrFonte 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   1800
      Width           =   1965
   End
   Begin VB.CheckBox ckResolucao 
      Caption         =   "Adaptar Resolução"
      Height          =   195
      Left            =   1560
      TabIndex        =   34
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CheckBox chGrvAutom 
      Caption         =   "Grava itens do orçamento automaticamente"
      Height          =   255
      Left            =   1560
      TabIndex        =   31
      ToolTipText     =   "Grava itens de orçamento automáticamente"
      Top             =   2760
      Width           =   3435
   End
   Begin VB.ComboBox cbOperacao 
      Height          =   315
      ItemData        =   "frmConfig.frx":0000
      Left            =   1560
      List            =   "frmConfig.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   3360
      Width           =   2025
   End
   Begin VB.Frame frLog 
      Caption         =   "Log"
      Height          =   555
      Left            =   1560
      TabIndex        =   26
      Top             =   3720
      Width           =   4215
      Begin VB.CheckBox ckLogEmRede 
         Alignment       =   1  'Right Justify
         Caption         =   "Log em Rede"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2760
         TabIndex        =   28
         Top             =   240
         Width           =   1275
      End
      Begin VB.CheckBox ckLog 
         Caption         =   "Ativar Log"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox TxCGC 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   780
      Width           =   1965
   End
   Begin VB.Frame frComiss 
      Caption         =   "Comissões"
      Height          =   1035
      Left            =   1560
      TabIndex        =   21
      Top             =   4320
      Width           =   4215
      Begin VB.TextBox txQtdCarr 
         Height          =   285
         Left            =   3060
         TabIndex        =   33
         ToolTipText     =   "Quantidade mínima de carros atendidos para liberar a comissão"
         Top             =   660
         Width           =   525
      End
      Begin VB.TextBox txVlrgatComiss 
         Height          =   285
         Left            =   3060
         TabIndex        =   23
         ToolTipText     =   "Valor mínimo para liberar as comissões"
         Top             =   360
         Width           =   825
      End
      Begin VB.CheckBox ckComiss 
         Caption         =   "Utiliza Comissões"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   180
         Width           =   1515
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nr de carros para liberar"
         Height          =   195
         Index           =   8
         Left            =   1320
         TabIndex        =   32
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor de gatilho para pagto de comissões"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   2910
      End
   End
   Begin VB.CheckBox chImprBranco 
      Caption         =   "Imprimir valores em branco"
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   1800
      Width           =   2235
   End
   Begin VB.TextBox txGarantia 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   2460
      Width           =   465
   End
   Begin VB.CommandButton btImagem 
      Appearance      =   0  'Flat
      Caption         =   "&Imagem"
      Height          =   300
      Left            =   4380
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Define a cor do programa"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txLinhasApos 
      Height          =   285
      Left            =   1560
      TabIndex        =   10
      Top             =   3060
      Width           =   465
   End
   Begin VB.TextBox txtImpressora 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Top             =   2130
      Width           =   3705
   End
   Begin VB.ComboBox cbImpress 
      Height          =   315
      ItemData        =   "frmConfig.frx":002F
      Left            =   1560
      List            =   "frmConfig.frx":003C
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1470
      Width           =   2025
   End
   Begin VB.CommandButton btDefCor 
      Appearance      =   0  'Flat
      Caption         =   "&Cor"
      Height          =   300
      Left            =   2250
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Define a cor do programa"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtEmpresa 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   60
      Width           =   3705
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   300
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   11
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtTelefone 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   1140
      Width           =   1965
   End
   Begin VB.TextBox txtEndereco 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   390
      Width           =   3705
   End
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   4140
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Defina a imagem de fundo para o programa"
      Filter          =   $"frmConfig.frx":005F
      MaxFileSize     =   255
   End
   Begin VB.Label lbNrFonte 
      Alignment       =   1  'Right Justify
      Caption         =   "Nr da Fonte: "
      Enabled         =   0   'False
      Height          =   225
      Left            =   600
      TabIndex        =   5
      Top             =   1860
      Width           =   915
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Modo de Operação:"
      Height          =   195
      Index           =   7
      Left            =   60
      TabIndex        =   29
      Top             =   3420
      Width           =   1425
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "CNPj"
      Height          =   225
      Index           =   6
      Left            =   660
      TabIndex        =   25
      Top             =   810
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "Meses de Garantia"
      Height          =   225
      Index           =   4
      Left            =   2100
      TabIndex        =   20
      Top             =   2550
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Linhas após a impressão"
      Height          =   225
      Index           =   3
      Left            =   2100
      TabIndex        =   19
      Top             =   3150
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Impressora: "
      Height          =   225
      Index           =   2
      Left            =   480
      TabIndex        =   18
      Top             =   2160
      Width           =   1035
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Impressão:"
      Height          =   225
      Index           =   1
      Left            =   720
      TabIndex        =   17
      Top             =   1500
      Width           =   795
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Telefone: "
      Height          =   225
      Index           =   0
      Left            =   720
      TabIndex        =   16
      Top             =   1170
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Endereço: "
      Height          =   225
      Left            =   720
      TabIndex        =   15
      Top             =   420
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Empresa: "
      Height          =   225
      Left            =   720
      TabIndex        =   14
      Top             =   90
      Width           =   795
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'4.9.5 Definir o nr do tamanho da fonte na impressão USB
'4.9.0 Previsão para impressora em Fita via Usb
'4.6.5 Retirada da configuração do modo de impressão a nivel de banco de dados
'3.8.4 Adaptação da resolução
'3.3.0 Critério de quantidade de carros para liberar as comissões
'3.2.3 Gravação automática dos itens
'3.1.0 Modo Balcão
'3.0.2 Log em Rede
'2.7.4-5 Campo de CGC na configuracao
'2.7.4 Opção para utilizar ou não as comissões
'2.7.3 Configuração para valor de gatilho para comissão
'2.5.6 Modo Restrito
'2.0.7 Garantia na configuração
'2.0.6 Alteração do funcionamento interno das variáveis de configuração
'2.0.5 Implantação do Log
'2.0.2 Aumentar os campos automaticamente
'2.0.2 Configuração para linhas após na impressão em fita
'2.0.1 Tela de configuração

Option Explicit

Private Mudando  As Boolean
Private lcVlrGat As Double

Private Sub btDefCor_Click()
Dialogo.Filter = "Todas Imagens (*.bmp, *.dib, *.gif, *.jpg, *.wmf, *.emf)|*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf|Bitmap (*.bmp)|*.bmp|Dib's (*.dib)|*.dib|Gif's (*.gif)|*gif|Jpeg's (*.jpg)|*.jpg|Metafile's (*.wmf)|*.wmf|(*.emf)|*.emf"
Dialogo.DialogTitle = "Escolha a cor do programa"
Dialogo.Color = btDefCor.BackColor
Dialogo.ShowColor

'2.0.6 Alteração do funcionamento interno das variáveis de configuração
INI.Cor = Dialogo.Color
CorSelec = 16777215 - INI.Cor
btDefCor.BackColor = INI.Cor
btDefCor.MaskColor = CorSelec
End Sub

Private Sub btImagem_Click()
Dim Diret      As String
Dim PathImagem As String
Dim TbConfig   As Recordset

Diret = CurDir()
Dialogo.ShowOpen
On Local Error GoTo Sai
If Dialogo.FileName <> "" Then
    PathImagem = Dialogo.FileName
    FrmMenu.Imagem.Picture = LoadPicture(PathImagem)
    
    '2.0.5 Implantação do Log
    AbreTB TbConfig, "Config", dbOpenTable
    'Set TbConfig = db.OpenRecordset("Config")
    
    AdapTamCampo PathImagem, TbConfig!Imagem, TbConfig, "Config"
    TbConfig.Edit
    TbConfig!Imagem = PathImagem
    TbConfig.Update
    TbConfig.Close
End If

Sai:
ChDir Diret
End Sub

Private Sub cbImpress_Click()
'4.9.5 Definir o nr do tamanho da fonte na impressão USB
Select Case cbImpress.ListIndex
    Case 0
        chImprBranco.Enabled = True
        txtImpressora.Enabled = False
        lbNrFonte.Enabled = False
        txNrFonte.Enabled = False
    Case 1
        chImprBranco.Enabled = False
        txtImpressora.Enabled = True
        lbNrFonte.Enabled = False
        txNrFonte.Enabled = False
    Case 2
        chImprBranco.Enabled = False
        txtImpressora.Enabled = False
        lbNrFonte.Enabled = True
        txNrFonte.Enabled = True
End Select
txLinhasApos.Enabled = txtImpressora.Enabled
End Sub

Private Sub ckComiss_Click()
'2.7.4 Opção para utilizar ou não as comissões
If Mudando = False Then
    If ckComiss.Value = 1 Then
        If lcVlrGat > 0 Then
            MostraValor txVlrgatComiss, lcVlrGat
        End If
    Else
        VeValor txVlrgatComiss.Text, lcVlrGat, txVlrgatComiss, 0
        txVlrgatComiss.Text = ""
    End If
    txVlrgatComiss.Text = IIf(ckComiss.Value = 1, txVlrgatComiss.Text, "")
End If
End Sub

Private Sub ckLog_Click()
'3.0.2 Log em Rede
If Mudando = False Then
    ckLogEmRede.Enabled = (ckLog.Value = 1)
End If
End Sub

Private Sub cmdGravar_Click()
Dim Result As Boolean
Dim SQL    As String

'2.0.2 Aumentar os campos automaticamente
Dim TbConfig As Recordset

Dim sVlrComiss As Double

'3.1.0 Modo Balcão
''2.5.6 Modo Restrito
'If INI.Restrito = True Then
'    Load frmSenha
'    frmSenha.Show 1
'    Result = frmSenha.Resultado
'    Unload frmSenha
'    If Result = False Then
'        Exit Sub
'    End If
'End If

'2.0.5 Implantação do Log
AbreTB TbConfig, "Config", dbOpenTable
'Set TbConfig = db.OpenRecordset("Config")

AdapTamCampo txtEmpresa.Text, TbConfig!Empresa, TbConfig, "Config"
AdapTamCampo txtEndereco.Text, TbConfig!Endereco, TbConfig, "Config"
AdapTamCampo txtTelefone.Text, TbConfig!Fones, TbConfig, "Config"
AdapTamCampo txtImpressora.Text, TbConfig!Empresa, TbConfig, "Config"

'2.0.6 Alteração do funcionamento interno das variáveis de configuração
INI.Empresa = txtEmpresa.Text
INI.Endereco = txtEndereco.Text
INI.Fones = txtTelefone.Text
INI.Cor = btDefCor.BackColor
INI.TpImpress = cbImpress.ListIndex
INI.LinhasApos = Val(txLinhasApos.Text)

'2.0.7 Garantia na configuração
INI.Garantia = Val(txGarantia.Text)

'4.6.5 Retirada da configuração do modo de impressão a nivel de banco de dados
SQL = "Update Config Set Empresa = '" & txtEmpresa.Text & _
    "', Endereco = '" & txtEndereco.Text & _
    "', Fones = '" & txtTelefone.Text & _
    "', Cor = " & INI.Cor & _
    ", Garantia = " & INI.Garantia
'3.3.4 Nr de linhas apos a impressao passa a ser informação local
'SQL = "Update Config Set Empresa = '" & txtEmpresa.Text & _
    "', Endereco = '" & txtEndereco.Text & _
    "', Fones = '" & txtTelefone.Text & _
    "', Cor = " & INI.Cor & _
    ", TpImpress = " & cbImpress.ListIndex & _
    ", Garantia = " & INI.Garantia

'2.0.5 Implantação do Log
ExecSql SQL
'db.Execute SQL
    
INI.AbreChave
INI.Impressora = txtImpressora.Text

'2.0.5 Implantação do Log
If ckLog.Value = 1 Then
    INI.Log = True
    
    '3.0.2 Log em Rede
    INI.LogEmRede = (ckLogEmRede.Value = 1)
End If
'INI.Log = (ckLog.Value = 1)

'3.1.0 Modo Balcão
'2.5.6 Modo Restrito
'VeRestrito

'2.7.2 Deixar de imprimir valores em zero, para impressão A4
INI.ImprEmBrano = chImprBranco.Value

'2.7.4 Opção para utilizar ou não as comissões
INI.UtComissoes = ckComiss.Value
If ckComiss.Value Then

    '2.7.3 Configuração para valor de gatilho para comissão
    VeValor txVlrgatComiss.Text, sVlrComiss, txVlrgatComiss, 1
    INI.VlrGatComiss = sVlrComiss
End If

'2.7.4-5 Campo de CGC na configuracao
INI.CGC = TxCGC.Text

'3.1.0 Modo Balcão
INI.ModoOperacao = cbOperacao.ListIndex
DefineMenus INI.ModoOperacao

'3.2.3 Gravação automática dos itens
INI.GravaAutom = chGrvAutom.Value

'3.3.0 Critério de quantidade de carros para liberar as comissões
INI.QtdCarrComiss = Val(txQtdCarr.Text)

'3.8.4 Adaptação da resolução
INI.Resolucao = ckResolucao.Value

'4.9.5 Definir o nr do tamanho da fonte na impressão USB
INI.NrFonte = Val(txNrFonte.Text)

INI.FechaChave

Unload Me
End Sub

'3.1.0 Modo Balcão
'Private Sub VeRestrito()
'Dim MudouRestrito As Boolean
'
''2.5.6 Modo Restrito
'If INI.Restrito = True Then
'    If Check1.Value = 0 Then
'        INI.Restrito = False
'        MudouRestrito = True
'    End If
'Else
'    If Check1.Value = 1 Then
'        INI.Restrito = True
'        MudouRestrito = True
'    End If
'End If
'If MudouRestrito = True Then
'    PoemTituloAplicacao FrmMenu
'    ModoRestrito INI.Restrito
'End If
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Dim TbConfig As Recordset
Dim sEmpresa As String

Mudando = True

InicForm Me

'2.0.5 Implantação do Log
AbreTB TbConfig, "Config", dbOpenTable
'Set TbConfig = db.OpenRecordset("Config")

'2.0.6 Alteração do funcionamento interno das variáveis de configuração
txtEmpresa.Text = IIf(INI.Empresa = "", PegaNmCliente, INI.Empresa)
txtEndereco.Text = INI.Endereco
txtTelefone.Text = INI.Fones
cbImpress.ListIndex = INI.TpImpress

'4.9.5 Definir o nr do tamanho da fonte na impressão USB
txNrFonte.Text = INI.NrFonte

'2.0.2 Configuração para linhas após na impressão em fita
txLinhasApos.Text = STR(INI.LinhasApos)

'2.0.7 Garantia na configuração
txGarantia.Text = Trim(STR(INI.Garantia))

TbConfig.Close

INI.AbreChave

'2.0.5 Implantação do Log
If INI.Log Then
    ckLog.Value = 1
    
    '3.0.2 Log em Rede
    ckLogEmRede.Enabled = True
    ckLogEmRede.Value = IIf(INI.LogEmRede = True, 1, 0)
End If
'ckLog.Value = IIf(INI.Log = True, 1, 0)

txtImpressora.Text = INI.Impressora

'3.1.0 Modo Balcão
'2.5.6 Modo Restrito
'Check1.Value = IIf(INI.Restrito = True, 1, 0)

'2.7.2 Deixar de imprimir valores em zero, para impressão A4
chImprBranco.Value = INI.ImprEmBrano

'2.7.4 Opção para utilizar ou não as comissões
ckComiss.Value = INI.UtComissoes
lcVlrGat = INI.VlrGatComiss
If ckComiss.Value Then

    '2.7.3 Configuração para valor de gatilho para comissão
    MostraValor txVlrgatComiss, lcVlrGat
End If

'2.7.4-5 Campo de CGC na configuracao
TxCGC.Text = INI.CGC

'3.1.0 Modo Balcão
cbOperacao.ListIndex = INI.ModoOperacao

'3.2.3 Gravação automática dos itens
chGrvAutom.Value = INI.GravaAutom

'3.3.0 Critério de quantidade de carros para liberar as comissões
txQtdCarr.Text = Trim(STR(INI.QtdCarrComiss))

'3.8.4 Adaptação da resolução
ckResolucao.Value = INI.Resolucao

INI.FechaChave
btImagem.BackColor = cmdGravar.BackColor

Mudando = False
End Sub

Private Sub txVlrgatComiss_Click()
'2.7.4 Opção para utilizar ou não as comissões
txVlrgatComiss.SelStart = 0
txVlrgatComiss.SelLength = Len(txVlrgatComiss.Text)
End Sub
