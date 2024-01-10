VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form CadMecanicos 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Cadastro de Mecânicos"
   ClientHeight    =   3840
   ClientLeft      =   1320
   ClientTop       =   360
   ClientWidth     =   6300
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
   ScaleHeight     =   3840
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4020
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadMecanicos.frx":0000
            Key             =   "Adição"
            Object.Tag             =   "Adicao"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadMecanicos.frx":059A
            Key             =   "Vendedor"
            Object.Tag             =   "Vendedor"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadMecanicos.frx":0F74
            Key             =   "Ficha"
            Object.Tag             =   "Ficha"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadMecanicos.frx":13C6
            Key             =   "Exclusão"
            Object.Tag             =   "Exclusao"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadMecanicos.frx":7128
            Key             =   "Pedido"
            Object.Tag             =   "Pedido"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadMecanicos.frx":83AA
            Key             =   "Orçamento"
            Object.Tag             =   "Orcamento"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadMecanicos.frx":86C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CadMecanicos.frx":89DE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   60
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cheques"
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frCliente 
      Caption         =   "Mecânico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2955
      Left            =   60
      TabIndex        =   8
      Top             =   840
      Width           =   6165
      Begin VB.ComboBox cbTipRecebimento 
         Height          =   360
         ItemData        =   "CadMecanicos.frx":24080
         Left            =   2040
         List            =   "CadMecanicos.frx":2408A
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2460
         Width           =   1425
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   5
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1320
         Width           =   1245
      End
      Begin VB.CheckBox chRecebe 
         Caption         =   "Recebe Pagamentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   15
         Top             =   2100
         Width           =   2115
      End
      Begin VB.ComboBox cbOper 
         Height          =   360
         ItemData        =   "CadMecanicos.frx":2409F
         Left            =   2040
         List            =   "CadMecanicos.frx":240AC
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2040
         Width           =   1425
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   4485
         MaxLength       =   8
         TabIndex        =   7
         Top             =   1680
         Width           =   945
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   3
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1680
         Width           =   645
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   2
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   4
         Top             =   960
         Width           =   2265
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   1
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   3
         Top             =   600
         Width           =   3300
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00C0C0C0&
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
         Height          =   300
         Index           =   0
         Left            =   2040
         MaxLength       =   48
         TabIndex        =   1
         Top             =   240
         Width           =   3300
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Recebimento:"
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
         Index           =   6
         Left            =   60
         TabIndex        =   18
         Top             =   2520
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "RG:"
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
         Index           =   5
         Left            =   1590
         TabIndex        =   16
         Top             =   1380
         Width           =   345
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Operador Tipo:"
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
         Index           =   4
         Left            =   690
         TabIndex        =   13
         Top             =   2100
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Senha: "
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
         Index           =   3
         Left            =   3780
         TabIndex        =   12
         Top             =   1740
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Perc.Comissão:"
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
         Index           =   2
         Left            =   675
         TabIndex        =   11
         Top             =   1770
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Telefones:"
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
         Index           =   1
         Left            =   1050
         TabIndex        =   9
         Top             =   1050
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Endereço:"
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
         Index           =   24
         Left            =   1110
         TabIndex        =   2
         Top             =   690
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Nome:"
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
         Index           =   0
         Left            =   1425
         TabIndex        =   0
         Top             =   330
         Width           =   555
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   780
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   1376
      ButtonWidth     =   1799
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Adição"
            Key             =   "Adição"
            Description     =   "Adição"
            Object.ToolTipText     =   "Adição"
            Object.Tag             =   "Adição"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exclusão"
            Key             =   "Exclusão"
            Description     =   "Exclusão"
            Object.ToolTipText     =   "Exclusão"
            Object.Tag             =   "Exclusão"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Gravar"
            Key             =   "Gravar"
            Description     =   "Gravar"
            Object.ToolTipText     =   "Gravar"
            Object.Tag             =   "Gravar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Recibo"
            Key             =   "Recibo"
            Description     =   "Recibo"
            Object.ToolTipText     =   "Recibo"
            Object.Tag             =   "Recibo"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ferramentas"
            Key             =   "Ferramentas"
            Description     =   "Ferramentas"
            Object.ToolTipText     =   "Ferramentas"
            Object.Tag             =   "Ferramentas"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.Menu Mnu_Adi 
      Caption         =   "&Adição"
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_Visu 
      Caption         =   "&Visualização"
      Visible         =   0   'False
      Begin VB.Menu MnuVis 
         Caption         =   "&Pedido"
         Index           =   0
      End
      Begin VB.Menu MnuVis 
         Caption         =   "&Orçamento"
         Index           =   1
      End
   End
   Begin VB.Menu Mnu_Ficha 
      Caption         =   "&Ficha"
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_Exclusao 
      Caption         =   "&Exclusão"
      Visible         =   0   'False
   End
   Begin VB.Menu Mnu_Vend 
      Caption         =   "&Vendedor"
      Visible         =   0   'False
      Begin VB.Menu Mnu_EscVend 
         Caption         =   "&Escolher Vendedor"
      End
   End
   Begin VB.Menu MnuEsq 
      Caption         =   "&Mnu"
      Visible         =   0   'False
      Begin VB.Menu MnuPOP 
         Caption         =   "titulo"
         Index           =   0
      End
      Begin VB.Menu MnuPOP 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MnuPOP 
         Caption         =   "Apa&gar"
         Index           =   2
      End
      Begin VB.Menu MnuPOP 
         Caption         =   "Mover para o cliente &anterior"
         Index           =   3
      End
      Begin VB.Menu MnuPOP 
         Caption         =   "Mover para o cliente &posterior"
         Index           =   4
      End
   End
End
Attribute VB_Name = "CadMecanicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'4.8.4 ReImpressão dos Vales
'3.8.5 Previsão para quando se clica nas ferramentas e ainda não tem o nr de mecânico
'3.7.7 Quando cadastra novo mecânico já permite adicionar ferramentas
'3.7.6 Mecânico com recebimento semanal
'3.7.0 Ferramentas
'3.6.3 Gravar o operador como ativo ao cadastrar novo operador
'3.6.3 Impedir de gravar operador sem informar o tipo
'3.5.1 Não excluir fisicamente Mecânico
'3.4.4 Vale Transporte
'3.3.7 Ajustes na adição de operadores
'3.3.3 Só permitir recebimentos de balconistas com permissão
'3.1.2 Ajuste na gravação da senha do operador
'3.1.0 Modo Balcão
'3.1.0 Cadastro de operadores
'2.7.4 Campo de senha para o mecânico
'2.7.3 Campo de comissões para Mecânicos
'2.7.2 Logar todas mensagens
'2.6.8 Log para verificar o endereço na impressão do orçamento
'2.6.7 Ajuste a primeira operação com cadastro de mecânicos
'2.6.4 Ajuste no rolamento dos mecânicos
'2.6.4 Atualizar o cadastro do mecânico sem precisar fechar a tela
'2.6.4 Impressão do recibo do mecânico
'2.6.3 Cadastro de Mecânicos

Option Explicit

Private Adicionando As Boolean
Private Mudado      As Integer
Private Mudando     As Integer
Private FimMec      As Integer
Private MecAtual    As Integer
Private TotalMec    As Integer
Private AtuNome     As Integer
Private PosMec      As Long
Private TbMec       As Recordset

Private Const Tabela = "Mecanicos"

Private Sub cbOper_Click()
'3.1.0 Cadastro de operadores
If Mudando = False Then
    Mudado = True
    cCli_evMudou True
End If
Label1(2).Enabled = (cbOper.ListIndex = 0)
Text(3).Enabled = (cbOper.ListIndex = 0)
End Sub

Private Sub cbTipRecebimento_Click()
'3.7.6 Mecânico com recebimento semanal
If Mudando = False Then
    Mudado = True
    cCli_evMudou True
End If
End Sub

Private Sub chRecebe_Click()
If Mudando = False Then
    Mudado = True
    cCli_evMudou True
End If
End Sub

Private Sub Form_Load()
Dim Sair        As Boolean
Dim TbMecEOF    As Boolean
Dim SQL         As String

Mudando = True
AtuNome = 0
FimMec = FunFimMec

If MecAtual > FimMec Or MecAtual < 1 Then
    MecAtual = 1
End If

FimMec = FunFimMec
Sair = False
If FimMec = 0 Then
    Adicionando = True
Else
    MecAtual = 0
    Do While Sair = False
        MecAtual = MecAtual + 1
        SQL = "Select Count(*) From " & Tabela & " Where codi = " & MecAtual
        If Consulta(SQL) > 0 Then
            SQL = "Select * From " & Tabela & " Where codi = " & MecAtual
            AbreTB TbMec, SQL, dbOpenDynaset
            Sair = True
        End If
    Loop
End If
If FimMec > 0 Then
    Mostra
End If
Mudando = False
Screen.MousePointer = vbDefault
Left = 500
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If TRATATECLA(KeyCode) Then
   If Index = 0 Then
      AtuNome = 1
   End If
   cCli_evMudou True
   Mudado = 1
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Tag
    Case "Adição"
        Adicao
    Case "Exclusão"
        Exclusao
    Case "Gravar"
        If GRAVA Then
            Toolbar1.Buttons(3).Enabled = False
        End If
    Case "Recibo"
        '3.4.7 Gravação dos recibos
        Load RecibosRealizados
        
        '3.5.1 Listagem dos vales dos funcionario
        RecibosRealizados.Funcionario% = MecAtual
        RecibosRealizados.Nome = Trim(TbMec("Nome"))
        
        '4.8.4 ReImpressão dos Vales
        RecibosRealizados.Endereco = Text(1).Text
        
        RecibosRealizados.Show
        
    '3.7.0 Ferramentas
    Case "Ferramentas"
        Load FerrMec
        FerrMec.Nome = Text(0).Text
        
        '3.8.5 Previsão para quando se clica nas ferramentas e ainda não tem o nr de mecânico
        On Error GoTo SemMecanico
        FerrMec.nrFunc = TbMec("codi")
        On Error GoTo 0
        FerrMec.Show
End Select
Exit Sub

'3.8.5 Previsão para quando se clica nas ferramentas e ainda não tem o nr de mecânico
SemMecanico:
MsgBox "Não será possível mostrar o cadastro de ferramentas neste momento", vbInformation, "OrCarro"
End Sub

Private Sub Exclusao()
If LTrim$(TbMec("Nome")) > "" Then
    '2.7.2 Logar todas mensagens
    If msgboxL("Confirma a exclusao do Mecânico", 4 + 256, "Exclusao de Mecânicos") = 7 Then Exit Sub
End If

'3.5.1 Não excluir fisicamente Mecânico
ExecSql "Update Mecanicos Set Ativo = False Where Codi = " & MecAtual
'ExecSql "Delete From " & Tabela & " Where Codi = " & MecAtual

'2.6.4 Impressão do recibo do mecânico
Toolbar1.Buttons(4).Enabled = False

Mudado = 0
TRATATECLA 34
TotalMec = 0
End Sub

Private Sub Adicao()
Dim SQL As String

If Mudado Then GRAVA
ZeraTela

SQL = "Select Max(Codi) From Clientes"
FimMec = FunFimMec + 1

MecAtual = FimMec

Loga "Adição - FimMec =" + STR(FimMec)

'3.3.7 Ajustes na adição de operadores
chRecebe.Value = 0

Text(0).SetFocus
Adicionando = True
End Sub

Public Function GRAVA() As Boolean
Dim TbMeccodi As Integer
Dim SQL       As String
Dim Aux#
Dim AuxVlrComiss As String

'3.1.2 Ajuste na gravação da senha do operador
Text(4).Text = Trim(Text(4).Text)

'3.1.0 Modo Balcão
If Text(4).Text > "" Then
    
    '3.5.1 Não excluir fisicamente Mecânico
    If Consulta("Select Count(*) From Mecanicos Where Senha = '" & Text(4).Text & "' and Nome <> '" & Text(0).Text & "' and Ativo = True ") > 0 Then
    'If Consulta("Select Count(*) From Mecanicos Where Senha = '" & Text(4).Text & "' and Nome <> '" & Text(0).Text & "'") > 0 Then
    
        MsgBox "Já existe outro mecânico com esta senha"
        Exit Function
    End If
End If

'3.6.3 Impedir de gravar operador sem informar o tipo
If cbOper.Text = "" Then
    MsgBox "É necessário informar o tipo de operador do sistema"
    cbOper.SetFocus
    Exit Function
End If

'3.7.6 Mecânico com recebimento semanal
If cbTipRecebimento.Text = "" Then
    MsgBox "É necessário informar o tipo recebimento do funcionário"
    cbTipRecebimento.SetFocus
    Exit Function
End If

1200 Mudado = 0
1201 If Text(0).Text > "" Then

    AuxVlrComiss = VlrSql(Replace(Text(3).Text, "%", ""))
            
1202   If Adicionando Then
1203      Adicionando = False
   
      'Novo
        
      SQL = "Select Count(*) From " & Tabela & " Where Nome = '" & Trim(Text(0).Text) & "'"
      If Consulta(SQL) > 0 Then
              '2.7.2 Logar todas mensagens
1209          msgboxL "Mecânico já cadastrado"
1210          Exit Function
1211       End If

          If MecAtual > 0 Then
            SQL = "Select Count(*) From " & Tabela & " Where Nome = '" & Trim(Text(0).Text) & "' and codi <> " & MecAtual
            If Consulta(SQL) > 0 Then
                '2.7.2 Logar todas mensagens
                msgboxL "Já existe um Mecânico com o nome que voce esta alterando"
                Text(0).Text = TbMec("Nome")
                Exit Function
            End If
          End If
                    
           FimMec = FimMec + 1
           MecAtual = FimMec
                      
           '3.7.6 Mecânico com recebimento semanal
           SQL = "Insert Into " & Tabela & "(codi, Nome, Ende, PercComiss, Senha, Telefone, Oper, Recebe, TpRec, RG, Ativo) Values (" & MecAtual & " ,"
           '3.6.3 Gravar o operador como ativo ao cadastrar novo operador
           'SQL = "Insert Into " & Tabela & "(codi, Nome, Ende, PercComiss, Senha, Telefone, Oper, Recebe, RG, Ativo) Values (" & MecAtual & " ,"
           
           SQL = SQL & "'" & Text(0).Text & "',"
           SQL = SQL & "'" & Text(1).Text & "',"
           
           '2.7.3 Campo de comissões para Mecânicos
           SQL = SQL & AuxVlrComiss & " ,'"
           
           '2.7.4 Campo de senha para o mecânico
           SQL = SQL & Text(4).Text & "'"
           
           '3.3.7 Ajustes na adição de operadores
           SQL = SQL & ", '" & Text(2).Text & "'"
           
           '3.1.0 Cadastro de operadores
           SQL = SQL & ", " & cbOper.ListIndex
           
           '3.3.7 Ajustes na adição de operadores
           SQL = SQL & ", " & chRecebe.Value
                      
           '3.7.6 Mecânico com recebimento semanal
           SQL = SQL & ", " & cbTipRecebimento.ListIndex
                      
           '3.6.3 Gravar o operador como ativo ao cadastrar novo operador
           SQL = SQL & ",'" & Text(5).Text
           SQL = SQL & "', true)"
           '3.4.4 Vale Transporte
           'SQL = SQL & ",'" & Text(5).Text & "')"
           
           ExecSql SQL
        
1213       Loga "TbMec('Codi' = " + STR(FimMec) + " )"
           TotalMec = 0
           
           '3.7.7 Quando cadastra novo mecânico já permite adicionar ferramentas
           If cbOper.ListIndex = 0 Then
             Toolbar1.Buttons(5).Enabled = True
           End If
           
       Else
          'EDIÇÂO
          If FimMec > 0 Then
            If Trim(TbMec("Nome")) <> Trim(Text(0).Text) Then
                '2.7.2 Logar todas mensagens
                If msgboxL("Confirma a alteração no nome", vbQuestion + vbYesNo + vbDefaultButton2, "Alteração de nome do Mecânico") = vbNo Then
                    Text(0).Text = TbMec("Nome")
                    Exit Function
                End If
            End If
            SQL = "Update " & Tabela
            SQL = SQL & " Set Nome = '" & Text(0).Text & "' , "
            SQL = SQL & "Ende = '" & Text(1).Text & "' , "
            SQL = SQL & "Telefone = '" & Text(2).Text & "'"
            
            '2.7.3 Campo de comissões para Mecânicos
            SQL = SQL & ",PercComiss = " & AuxVlrComiss & " , "
            
            '2.7.4 Campo de senha para o mecânico
            SQL = SQL & "Senha = '" & Text(4).Text & "'"
            
            '3.1.0 Cadastro de operadores
            SQL = SQL & ", Oper = " & cbOper.ListIndex
            
            '3.3.3 Só permitir recebimentos de balconistas com permissão
            SQL = SQL & ", Recebe = " & chRecebe.Value
            
            '3.4.4 Vale Transporte
            SQL = SQL & ", RG = '" & Text(5).Text & "' "
            
            '3.7.6 Mecânico com recebimento semanal
            SQL = SQL & ", TpRec = " & cbTipRecebimento.ListIndex
            
            SQL = SQL & " Where codi = " & MecAtual
            ExecSql SQL
          End If
End If

'2.6.7 Ajuste a primeira operação com cadastro de mecânicos
On Local Error Resume Next
'2.6.4 Atualizar o cadastro do mecânico sem precisar fechar a tela
TbMec.Requery
On Local Error GoTo 0

Ret_Erro_GRAVA:
1278    If MecAtual = 0 Then
            MecAtual = 1
        End If
1280 End If
GRAVA = True
1281 Exit Function

Erro_GRAVA:
1283 Resume Ret_Erro_GRAVA

ErrNaGravacao:
Loga "erro de tipo " & Err & " = " & Error(Err) & " na linha " & Erl, lERR
'2.7.2 Logar todas mensagens
msgboxL "Erro de tipo " & vbCrLf & Error(Err) & vbCrLf & " na linha " & Erl, vbCritical, "Erro na gravação do Mecânico"
Resume Next
End Function

Private Sub cCli_evMudou(Estado As Boolean)
Toolbar1.Buttons(3).Enabled = Estado
End Sub

Sub ZeraTela()
Dim C As Control

For Each C In Screen.ActiveForm
   If TypeOf C Is TextBox Then
      C.Text = ""
   ElseIf TypeOf C Is ComboBox Then
      C.Tag = "M"
      C.ListIndex = -1
      C.Tag = ""
   ElseIf TypeOf C Is ListBox Then
      C.Clear
   End If
Next
End Sub

Public Function FunFimMec() As Integer
FunFimMec = Consulta("Select Max(Codi) from " & Tabela)
End Function

Public Function TRATATECLA(Tecla As Integer, Optional NaoGrava As Boolean) As Integer
Select Case Tecla
    Case Is = 27
        'Esc
        If Mudado Then
            Mostra
        Else
            Unload CadMecanicos
        End If
    Case Is = 33
PageUp:
        If NaoGrava = False Then
        
            '2.6.4 Ajuste no rolamento dos mecânicos
            PosMec = Consulta("Select TOP 1 codi From " & Tabela & " Where Nome < '" & Text(0).Text & "' Order by Nome desc")
            'PosMec = Consulta("Select codi From " & Tabela & " Where Nome < '" & Text(0).Text & "'")
            'PosMec = MecAtual - 1
            If PosMec > 0 Then
                If Mudado Then GRAVA
                
                LeCliente PosMec
                If StringBoa(TbMec("Nome"), True) = "" Then GoTo PageUp
                Mostra
            End If
        End If
    Case Is = 34
PageDown:
        If NaoGrava = False Then
            If Mudado Then GRAVA
            'PosMec = MecAtual + 1
            
            '2.6.4 Ajuste no rolamento dos mecânicos
            PosMec = Consulta("Select top 1 codi From " & Tabela & " Where Nome > '" & Text(0).Text & "' order by Nome")
            'PosMec = Consulta("Select codi From " & Tabela & " Where Nome > '" & Text(0).Text & "'")
            
            If PosMec > 0 Then
                LeCliente PosMec
                If StringBoa(TbMec("Nome"), True) = "" Then GoTo PageDown
                Mostra
            End If
            TRATATECLA = 0
        End If
    Case Is = 16, 17, 45
        'Somente para nao marcar para salvar
    Case Else
        Mudado = 1
        TRATATECLA = 1
End Select
End Function

Private Sub Mostra()
Mudando = True
Text(0).Text = Trim(TbMec("Nome"))
Text(1).Text = Trim(SN(TbMec("Ende"), vbString))
Text(2).Text = SN(TbMec("Telefone"), vbString)
Mudando = False
Mudado = 0

'2.6.4 Impressão do recibo do mecânico
Toolbar1.Buttons(4).Enabled = True

'2.7.3 Campo de comissões para Mecânicos
MostraValor Text(3), TbMec!PercComiss, False

'2.7.4 Campo de senha para o mecânico
Text(4).Text = SN(TbMec!Senha, vbString)

Text(3).Text = Text(3).Text & "%"

'3.1.0 Cadastro de operadores
cbOper.ListIndex = TbMec!Oper

'3.3.3 Só permitir recebimentos de balconistas com permissão
chRecebe.Value = SN(TbMec!Recebe, vbInteger)

'3.4.4 Vale Transporte
Text(5).Text = SN(TbMec!RG, vbString)

'3.7.0 Ferramentas
Toolbar1.Buttons(5).Enabled = (TbMec!Oper = 0)

'3.7.6 Mecânico com recebimento semanal
cbTipRecebimento.ListIndex = TbMec!tpRec

cCli_evMudou False
End Sub

Private Sub LeCliente(a As Long)
Static PosAnt As Integer
Dim Sair      As Boolean
Dim SQL       As String

Sair = False
Do While Sair = False
    SQL = "Select Count(*) From " & Tabela & " Where codi = " & a
    If Consulta(SQL) = 0 Then
        If a > PosAnt Then
            a = a + 1
        Else
            a = a - 1
            If a < 1 Then
                Sair = True
            End If
        End If
    Else
        Sair = True
    End If
Loop
MecAtual = a
PosAnt = a
AbreTB TbMec, "Select * From " & Tabela & " Where codi = " & a
End Sub

Public Sub Mostrar(Nome As String)
Dim nrCodi As Long

nrCodi = Consulta("Select codi from " & Tabela & " Where Nome = '" & Nome & "'")
LeCliente nrCodi
Mostra
End Sub
