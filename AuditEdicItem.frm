VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form AuditEdicItem 
   Caption         =   "Auditoria da edição do item"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   10380
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Bindings        =   "AuditEdicItem.frx":0000
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   10398
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
   End
End
Attribute VB_Name = "AuditEdicItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'4.0.2 Identificar quem editou o item do orçamento

Option Explicit

Private JaMostrou As Boolean
Private lcNrItem  As Long

Private Sub Form_Activate()
If JaMostrou = False Then
    JaMostrou = True
    Mostra
End If
End Sub

Private Sub Mostra()
Dim a      As Integer
Dim TbVend As Recordset
Dim SQL    As String

SQL = "SELECT Item, Quant, Valor, DtItOrc, Nome "
SQL = SQL & "from AuditItemOrc "
SQL = SQL & "INNER JOIN Mecanicos ON AuditItemOrc.Balconista = Mecanicos.codi "
SQL = SQL & "Where Item_Orc = " & NrItem
SQL = SQL & " Order By DtItOrc "
Set TbVend = db.OpenRecordset(SQL)

Grid.ColWidth(0) = 3900
Grid.ColWidth(1) = 600
Grid.ColWidth(2) = 800
Grid.ColWidth(3) = 1700
Grid.ColWidth(4) = 3000

Grid.Row = 0: Grid.Col = 0: Grid.Text = "Item"
Grid.Row = 0: Grid.Col = 1: Grid.Text = "Quant"
Grid.Row = 0: Grid.Col = 2: Grid.Text = "Valor"
Grid.Row = 0: Grid.Col = 3: Grid.Text = "Data"
Grid.Row = 0: Grid.Col = 4: Grid.Text = "Operador"

TbVend.MoveFirst
On Error GoTo 0
For a = 0 To TbVend.RecordCount - 1
   Grid.AddItem TbVend("Item"), a + 1
   Grid.Row = a + 1
   Grid.Col = 1: Grid.Text = TbVend("Quant")
   Grid.Col = 2: Grid.Text = Format(TbVend("Valor"), "#,###,##0.00")
   Grid.Col = 3: Grid.Text = TbVend("DtItOrc")
   Grid.Col = 4: Grid.Text = TbVend("Nome")
   TbVend.MoveNext
Next

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
JaMostrou = False
End Sub

Public Property Get NrItem() As Long
NrItem = lcNrItem
End Property

Public Property Let NrItem(ByVal vNewValue As Long)
lcNrItem = vNewValue
End Property
