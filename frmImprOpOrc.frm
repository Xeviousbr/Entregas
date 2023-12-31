VERSION 5.00
Object = "{00028C4A-0000-0000-0000-000000000046}#5.0#0"; "TDBG5.OCX"
Begin VB.Form frmImprOpOrc 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opções de Orçamento para Mecânica"
   ClientHeight    =   9195
   ClientLeft      =   150
   ClientTop       =   -1575
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   14715
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Dados 
      Caption         =   "DataT"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Z:\Share\Orcarro\OrCarro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   2  'Add New
      Exclusive       =   0   'False
      Height          =   435
      Index           =   3
      Left            =   11340
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select Valor, Conteudo From ItensConcertoTemp Where Coluna = 2"
      Top             =   660
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Dados 
      Caption         =   "DataT"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Z:\Share\Orcarro\OrCarro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   2  'Add New
      Exclusive       =   0   'False
      Height          =   435
      Index           =   2
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select Valor, Conteudo From ItensConcertoTemp Where Coluna = 2"
      Top             =   660
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salvar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   1215
   End
   Begin VB.Data Dados 
      Caption         =   "DataT"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Z:\Share\Orcarro\OrCarro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   2  'Add New
      Exclusive       =   0   'False
      Height          =   435
      Index           =   1
      Left            =   4020
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select Valor, Conteudo From ItensConcertoTemp Where Coluna = 2"
      Top             =   660
      Visible         =   0   'False
      Width           =   2775
   End
   Begin TrueDBGrid50.TDBGrid TDBGrid1 
      Bindings        =   "frmImprOpOrc.frx":0000
      Height          =   8355
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "frmImprOpOrc.frx":0017
      TabIndex        =   0
      Top             =   780
      Width           =   3615
   End
   Begin VB.Data Dados 
      Caption         =   "DataT"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Z:\Share\Orcarro\OrCarro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   2  'Add New
      Exclusive       =   0   'False
      Height          =   435
      Index           =   0
      Left            =   780
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select Valor, Conteudo From ItensConcertoTemp Where Coluna = 1 and IDPC = 8"
      Top             =   720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin TrueDBGrid50.TDBGrid TDBGrid1 
      Bindings        =   "frmImprOpOrc.frx":2391
      Height          =   8355
      Index           =   1
      Left            =   3720
      OleObjectBlob   =   "frmImprOpOrc.frx":23A8
      TabIndex        =   1
      Top             =   780
      Width           =   3615
   End
   Begin TrueDBGrid50.TDBGrid TDBGrid1 
      Bindings        =   "frmImprOpOrc.frx":4722
      Height          =   8355
      Index           =   2
      Left            =   7380
      OleObjectBlob   =   "frmImprOpOrc.frx":4739
      TabIndex        =   8
      Top             =   780
      Width           =   3615
   End
   Begin TrueDBGrid50.TDBGrid TDBGrid1 
      Bindings        =   "frmImprOpOrc.frx":6AB3
      Height          =   8355
      Index           =   3
      Left            =   11040
      OleObjectBlob   =   "frmImprOpOrc.frx":6ACA
      TabIndex        =   9
      Top             =   780
      Width           =   3615
   End
   Begin VB.Label lbTit 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   11040
      TabIndex        =   6
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label lbTit 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   7380
      TabIndex        =   5
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label lbTit 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   3720
      TabIndex        =   4
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label lbTit 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "frmImprOpOrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'3.1.3 Itens do Orçamento acabamento
'2.5.0 Configuração do modelo de itens

Option Explicit

Private lcNrOrc As Long

Private Sub Command1_Click()
Dim a   As Integer
Dim SQL As String

ExecSql "Delete From IC_Orc Where Orc = " & lcNrOrc
For a = 0 To 3
    Dados(a).Recordset.MoveFirst
    Do While Dados(a).Recordset.EOF = False
        If Dados(a).Recordset!Valor = 1 Then
            SQL = "Insert Into IC_Orc (Orc, Col, Lin) Values ("
            SQL = SQL & lcNrOrc
            SQL = SQL & ", " & (a + 1)
            SQL = SQL & ", " & Dados(a).Recordset!Linha & ")"
            ExecSql SQL
        End If
        Dados(a).Recordset.MoveNext
    Loop
Next
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub ColocaDados()
Dim a           As Integer
Dim b           As Integer
Dim NrAtu       As Integer
Dim IDPC        As Integer
Dim NaoTemMais  As Boolean
Dim SQL         As String
Dim rsTitulos   As Recordset
Dim rsItens(3)  As Recordset
Dim rsIC_Orc(3) As Recordset

AbreTB rsTitulos, "Select TitModelo1, TitModelo2, TitModelo3, TitModelo4 From Config "
lbTit(0).Caption = rsTitulos!TitModelo1
lbTit(1).Caption = rsTitulos!TitModelo2
lbTit(2).Caption = rsTitulos!TitModelo3
lbTit(3).Caption = rsTitulos!TitModelo4
IDPC = INI.PC
For a = 0 To 3
    Dados(a).DatabaseName = Base
    SQL = "Select Lin From IC_Orc Where Col = " & (a + 1)
    SQL = SQL & " and ORC = " & lcNrOrc
    AbreTB rsIC_Orc(a), SQL
    AbreTB rsItens(a), "Select Conteudo From ConfigModelo Where Conteudo > '' and Coluna = " & (a + 1) & " Order By Linha"
    If rsIC_Orc(a).EOF = False Then
        NrAtu = rsIC_Orc(a).Fields(0).Value
        rsIC_Orc(a).MoveNext
        NaoTemMais = False
    Else
        NaoTemMais = True
    End If
    b = 1
    Do While rsItens(a).EOF = False
        SQL = "Insert Into ItensConcertoTemp (Coluna, Linha, Valor, Conteudo, IDPC) Values ("
        SQL = SQL & (a + 1)
        SQL = SQL & ", " & b
        If b < NrAtu Or NaoTemMais Then
            SQL = SQL & ", " & "0"
        Else
            SQL = SQL & ", " & "1"
            If rsIC_Orc(a).EOF = False Then
                NrAtu = rsIC_Orc(a).Fields(0).Value
                rsIC_Orc(a).MoveNext
            Else
                NaoTemMais = True
            End If
        End If
        SQL = SQL & ", '" & rsItens(a)!Conteudo
        SQL = SQL & "', " & IDPC & ")"
        ExecSql SQL
        rsItens(a).MoveNext
        b = b + 1
    Loop
Next
For a = 0 To 3
    Dados(a).DatabaseName = Base
    SQL = "Select Valor, Conteudo, Linha From ItensConcertoTemp Where Coluna = " & (a + 1)
    SQL = SQL & " and IDPC = " & INI.PC
    SQL = SQL & " Order By Linha"
    Loga SQL
    Dados(a).RecordSource = SQL
    Dados(a).Refresh
Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
ExecSql "Delete From ItensConcertoTemp Where IDPC = " & INI.PC
End Sub

Public Property Let NrOrc(ByVal vNewValue As Long)
lcNrOrc = vNewValue
ColocaDados
End Property

Private Sub TDBGrid1_BeforeColEdit(Index As Integer, ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
Command1.Enabled = True
End Sub
