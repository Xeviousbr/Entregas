VERSION 5.00
Begin VB.Form frmMecOrc 
   Caption         =   "Mecânicos do Orçamento"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5145
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   5145
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   2580
      TabIndex        =   10
      Top             =   960
      Width           =   2355
   End
   Begin VB.CommandButton btRemove 
      Caption         =   "Remover"
      Height          =   315
      Left            =   2580
      TabIndex        =   9
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton btCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   3060
      Width           =   1095
   End
   Begin VB.CommandButton btOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   3060
      Width           =   1095
   End
   Begin VB.CommandButton btAdic 
      Caption         =   "Adicionar"
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   2640
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   180
      TabIndex        =   1
      Top             =   960
      Width           =   2355
   End
   Begin VB.Label lbVlrTot 
      Caption         =   "8.888,88"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   4020
      TabIndex        =   5
      Top             =   120
      Width           =   780
   End
   Begin VB.Label lbVlrTot 
      Caption         =   "Valor a definir: "
      Height          =   195
      Index           =   2
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   1020
   End
   Begin VB.Label lbVlrTot 
      Caption         =   "8.888,88"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1980
      TabIndex        =   3
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Mecânicos disponíveis"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   720
      Width           =   2400
   End
   Begin VB.Label lbVlrTot 
      Caption         =   "Valor total de mão de obra: "
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1920
   End
End
Attribute VB_Name = "frmMecOrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lcNrOrc      As Long
Dim lcVlrMaoObra As Single
Dim VlrComis(20) As Single

Public Property Get NrOrc() As Long
NrOrc = lcNrOrc
End Property

Public Property Let NrOrc(ByVal vNewValue As Long)
lcNrOrc = NrOrc
End Property

Public Property Get VlrMaoObra() As Single
lcVlrMaoObra = VlrMaoObra
End Property

Public Property Let VlrMaoObra(ByVal vNewValue As Single)
lcVlrMaoObra = vNewValue
MostraValor lbVlrTot(1), lcVlrMaoObra
VlrDisp = lcVlrMaoObra
End Property

Private Sub btAdic_Click()
Dim Pesq$

Pesq$ = InputBox("Informe o valor a que se refere a comissão")
If Pesq$ > "" Then
    List2.AddItem List1.Text
    List2.ItemData(List2.ListCount) = List1.ItemData(List1.ListIndex)
    List1.RemoveItem List1.ListIndex
    'VlrComis(
    'Talvez colocar utilizar registro
End If
End Sub

Private Sub btCancel_Click()
Unload Me
End Sub

Private Sub btOk_Click()
Dim a As Integer

ExecSql "Delete From Mec_Orc Where Orc = " & lcNrOrc
Stop
For a = 0 To List2.ListCount
    ExecSql "Insert Into Mec_Orc "
    
    SQL = "Insert Into Mec_Orc (Orc, Mec, Vlr) Values (" & lcNrOrc & ","
    SQL = SQL & List2.ItemData(a) & " , "
    'SQL = SQL & VlrSql(txtFields(5).Text & "',"
    'SQL = SQL & "'" & txtFields(6).Text & "')"
        
    'Mec
    'Vlr
Next
End Sub

Private Sub Form_Load()
Dim ListMecs As Recordset

AbreTB ListMecs, "Select Nome, codi From Mecanicos Order By Nome", dbOpenDynaset
Do While ListMecs.EOF = False
    List1.AddItem ListMecs!Nome
    List1.ItemData(List1.ListCount - 1) = ListMecs!codi
    ListMecs.MoveNext
Loop
ListMecs.Close
End Sub

Public Property Get VlrDisp() As Single
VlrDisp = lcVlrDisp
End Property

Public Property Let VlrDisp(ByVal vNewValue As Single)
lcVlrDisp = vNewValue
MostraValor lbVlrTot(3), lcVlrDisp
End Property
