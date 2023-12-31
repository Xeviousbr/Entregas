VERSION 5.00
Object = "{00028C4A-0000-0000-0000-000000000046}#5.0#0"; "TDBG5.OCX"
Begin VB.Form frmModeloItens 
   Caption         =   "Configuração do modelo de itens"
   ClientHeight    =   9390
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9390
   ScaleMode       =   0  'User
   ScaleWidth      =   13095
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Impr.antiga"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2100
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   360
      Top             =   2880
   End
   Begin VB.Data DataY 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Z:\Share\Orcarro\OrCarro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   3
      Left            =   10200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * From ConfigModelo Where Coluna = 4 Order By Linha"
      Top             =   8880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data DataY 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Z:\Share\Orcarro\OrCarro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   2
      Left            =   7260
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * From ConfigModelo Where Coluna = 3 Order By Linha"
      Top             =   8880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data DataY 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Z:\Share\Orcarro\OrCarro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   4260
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * From ConfigModelo Where Coluna = 2 Order By Linha"
      Top             =   8880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1620
      Width           =   1095
   End
   Begin VB.Data DataY 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Z:\Share\Orcarro\OrCarro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Index           =   0
      Left            =   1380
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select * From ConfigModelo Where Coluna = 1 Order By Linha"
      Top             =   8880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1140
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   660
      Width           =   1095
   End
   Begin VB.TextBox txTitulo 
      Height          =   315
      Index           =   3
      Left            =   10140
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox txTitulo 
      Height          =   315
      Index           =   2
      Left            =   7200
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox txTitulo 
      Height          =   315
      Index           =   1
      Left            =   4260
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox txTitulo 
      Height          =   315
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin TrueDBGrid50.TDBGrid TDBGridX 
      Bindings        =   "frmModeloItens.frx":0000
      Height          =   8655
      Index           =   0
      Left            =   1320
      OleObjectBlob   =   "frmModeloItens.frx":0017
      TabIndex        =   7
      Top             =   480
      Width           =   2895
   End
   Begin TrueDBGrid50.TDBGrid TDBGridX 
      Bindings        =   "frmModeloItens.frx":2209
      Height          =   8655
      Index           =   1
      Left            =   4260
      OleObjectBlob   =   "frmModeloItens.frx":2220
      TabIndex        =   9
      Top             =   480
      Width           =   2895
   End
   Begin TrueDBGrid50.TDBGrid TDBGridX 
      Bindings        =   "frmModeloItens.frx":440E
      Height          =   8655
      Index           =   2
      Left            =   7200
      OleObjectBlob   =   "frmModeloItens.frx":4425
      TabIndex        =   10
      Top             =   480
      Width           =   2895
   End
   Begin TrueDBGrid50.TDBGrid TDBGridX 
      Bindings        =   "frmModeloItens.frx":660B
      Height          =   8655
      Index           =   3
      Left            =   10140
      OleObjectBlob   =   "frmModeloItens.frx":6622
      TabIndex        =   11
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Titulos"
      Height          =   195
      Left            =   780
      TabIndex        =   1
      Top             =   180
      Width           =   495
   End
End
Attribute VB_Name = "frmModeloItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'3.2.7 Gravação do valor da mão de obra
'3.1.0 Modo Balcão
'2.5.9 Impressao dos modelo de itens em novo formato``
'2.5.6 Modo Restrito
'2.5.1 Ajuste na localização da base de dados na tela do modelo de itens de orçamento
'2.5.0 Configuração do modelo de itens

Private Adicionando As Boolean

Private Sub Command1_Click()
Dim SQL As String

SQL = "Update Config Set TitModelo1 = '" & txTitulo(0).Text
SQL = SQL & "', TitModelo2 = '" & txTitulo(1).Text
SQL = SQL & "', TitModelo3 = '" & txTitulo(2).Text
SQL = SQL & "', TitModelo4 = '" & txTitulo(3).Text & "'"
ExecSql SQL
End Sub

Private Sub Impressao1(Quant As Integer)
Dim a        As Integer

Load frmImprOpOrc
For a = 1 To Quant
    frmImprOpOrc.PrintForm
Next

Unload frmImprOpOrc
Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Load frmImprOpOrc
frmImprOpOrc.PrintForm

End Sub

Private Sub Command2_Click()
'2.5.9 Impressao dos modelo de itens em novo formato``
Dim vezes     As Integer
Dim Resposta  As String

Resposta = InputBox("Quantas Páginas deseja imprimir", "Impressão do Modelo", 1)
If Resposta > "" Then
    For vezes = 1 To Val(Resposta)
        ImprNova
    Next
End If
End Sub

Private Sub ImprNova()
'2.5.9 Impressao dos modelo de itens em novo formato``
Dim lin       As Integer
Dim posY      As Integer
Dim Sair      As Boolean
Dim Impresso  As Boolean
Dim Linha     As String
Dim rsTitulos As Recordset
Dim RsTab(4)  As Recordset

Const Marg = "    "
Const Marg2 = "            "
Const Marg3 = 1250
Const TamCol = 2800
Const Altura = 300

Printer.FontSize = 18
Printer.FontBold = True
Printer.Print Marg & CENTRAL("Opções de Orçamento para Mecânica", 40)
Printer.FontSize = 14
Printer.Print CENTRAL(Consulta("Select Empresa From Config"), 70)
Printer.Print
Printer.Print Marg2 & "Nome: _____________________________________________________"
Printer.Print
Printer.Print Marg2 & "Endereço: _________________________________________________"
Printer.Print
Printer.Print Marg2 & "Bairro: __________________________  Telefone: _____________"
Printer.Print
Printer.Print Marg2 & "Carro: _________________ Cor: ___________ Placa: __________"
posY = 3600
AbreTB rsTitulos, "Select TitModelo1, TitModelo2, TitModelo3, TitModelo4 From Config "
For a = 0 To 3
    If IsNull(rsTitulos.Fields(a).Value) = False Then
        Printer.CurrentX = Marg3 + TamCol * a
        Printer.CurrentY = posY
        Printer.Print rsTitulos.Fields(a).Value
    End If
Next
rsTitulos.Close
Printer.FontSize = 10

posY = 4200
Do While Sair = False
    Impresso = False
    For a = 0 To 3
        If lin = 0 Then
            AbreTB RsTab(a), DataY(a).RecordSource
        End If
        If RsTab(a).EOF = False Then
            If RsTab(a)!Conteudo > "" Then
                Printer.CurrentX = Marg3 + TamCol * a
                Printer.CurrentY = posY
                Printer.Print "[   ] " & RsTab(a)!Conteudo
                RsTab(a).MoveNext
                Impresso = True
            End If
        End If
    Next
    lin = lin + 1
    Sair = Not (Impresso)
    posY = posY + Altura
Loop

'3.1.3 Itens do Orçamento acabamento
Linha = "____________________________________________________________________________________"
Printer.FontSize = 10
Printer.Print
Printer.Print Marg2 & Linha
Printer.Print
Printer.Print Marg2 & Linha
Printer.Print
Printer.Print Marg2 & Linha
Printer.Print
Printer.FontSize = 14
Printer.Print Marg2 & "Ass. Cliente ________________________________________________"
Printer.EndDoc
End Sub

Private Sub Form_Load()
Dim a%, b%
Dim rsTitulos As Recordset

AbreATabela:
AbreTB rsTitulos, "Select TitModelo1, TitModelo2, TitModelo3, TitModelo4 From Config "

'2.5.1 Ajuste na localização da base de dados na tela do modelo de itens de orçamento
DataY(0).DatabaseName = App.Path & "\OrCarro.mdb"
DataY(1).DatabaseName = DataY(0).DatabaseName
DataY(2).DatabaseName = DataY(0).DatabaseName
DataY(3).DatabaseName = DataY(0).DatabaseName

If IsNull(rsTitulos!TitModelo1) = False Then
    txTitulo(0).Text = rsTitulos!TitModelo1
    DataY(0).Enabled = True
End If
If IsNull(rsTitulos!TitModelo2) = False Then
    txTitulo(1).Text = rsTitulos!TitModelo2
    DataY(1).Enabled = True
End If
If IsNull(rsTitulos!TitModelo3) = False Then
    txTitulo(2).Text = rsTitulos!TitModelo3
    DataY(2).Enabled = True
End If
If IsNull(rsTitulos!TitModelo4) = False Then
    txTitulo(3).Text = rsTitulos!TitModelo4
    DataY(3).Enabled = True
End If

'3.1.0 Modo Balcão
If INI.ModoOperacao = tpBalcao Then
'2.5.6 Modo Restrito
'If INI.Restrito Then
    
    '3.2.7 Gravação do valor da mão de obra
    Me.Caption = "Tabela de Preços"
    For b% = 0 To 3
        txTitulo(b%).Locked = True
        For a% = 2 To 3
            TDBGridX(b%).Columns(a%).Locked = True
        Next a%
    Next b%
    
'    TDBGridX(0).Enabled = False
'    TDBGridX(1).Enabled = False
'    TDBGridX(2).Enabled = False
'    TDBGridX(3).Enabled = False

End If
End Sub

Private Sub TDBGridX_AfterUpdate(Index As Integer)
Dim SQL As String

If Adicionando Then
    SQL = "Update ConfigModelo Set Coluna = " & (Index + 1) & ", Linha = " & DataY(Index).Recordset.RecordCount
    SQL = SQL & " Where Coluna is null and Linha is null "
    ExecSql SQL
    Adicionando = False
End If
End Sub

Private Sub TDBGridX_BeforeInsert(Index As Integer, Cancel As Integer)
Adicionando = True
End Sub

Private Sub txTitulo_Change(Index As Integer)
If txTitulo(Index).Text > "" Then
    If DataY(Index).Enabled = False Then
        DataY(Index).Enabled = True
    End If
        
    '3.1.0 Modo Balcão
    If INI.ModoOperacao = tpEscritorio Then
    '2.5.6 Modo Restrito
    'If INI.Restrito = False Then
    
        Command1.Enabled = True
    End If
End If
End Sub
