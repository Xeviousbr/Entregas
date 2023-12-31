VERSION 5.00
Begin VB.Form frmRelComissoes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Comissões"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3645
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      Height          =   615
      Left            =   60
      TabIndex        =   7
      Top             =   420
      Width           =   3375
      Begin VB.CheckBox Check 
         Caption         =   "Pagamento"
         Height          =   375
         Index           =   2
         Left            =   2100
         TabIndex        =   10
         Top             =   180
         Width           =   1155
      End
      Begin VB.CheckBox Check 
         Caption         =   "Comissão"
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   9
         Top             =   180
         Width           =   975
      End
      Begin VB.CheckBox Check 
         Caption         =   "Vale"
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   180
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Relatório"
      Height          =   375
      Left            =   1215
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txDtFIM 
      Height          =   285
      Left            =   2460
      MaxLength       =   20
      TabIndex        =   2
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txDtINI 
      Height          =   285
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   1
      Top             =   60
      Width           =   975
   End
   Begin VB.ComboBox cbVend 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1140
      Width           =   2355
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "até"
      Height          =   195
      Left            =   2100
      TabIndex        =   5
      Top             =   120
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Apartir de"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Vendedores"
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   1020
   End
End
Attribute VB_Name = "frmRelComissoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If IsDate(txDtINI.Text) = False Or IsDate(txDtFIM.Text) = False Then
    msgboxL "Data Inválida"
Else
    If Check(0).Value = 0 And Check(1).Value = 0 And Check(2).Value = 0 Then
        msgboxL "Escolha pelo menos um tipo"
    Else
        'fazer o relatório
        Impressao
    End If
End If
End Sub

Private Sub Form_Load()
Dim TbMec As Recordset

txDtFIM.Text = Format(Now, "DD/MM/YYYY")
txDtINI.Text = "01" + Right(txDtFIM.Text, 8)
AbreTB TbMec, "Select Nome From Mecanicos Where Ativo = True and Nome > '' and Oper = 0 Order by Nome ", dbOpenDynaset
cbVend.AddItem "Todos"
Do While TbMec.EOF = False
    cbVend.AddItem TbMec.Fields("Nome")
    TbMec.MoveNext
Loop
TbMec.Close
cbVend.ListIndex = 0
End Sub

Private Sub Impressao()
Dim a      As Integer
Dim Cont   As Integer
Dim sValor As Single
Dim Aux    As String
Dim SQL    As String
Dim Tpo    As String
Dim Total  As Currency

Const TamFita = 55

ImprBuferizada_Inicializa

Aux = "RELATÓRIO DE PAGAMENTOS "
Aux = Aux & "Tipo"
For a = 0 To 2
    If Check(a).Value Then
        Cont = Cont + 1
        Tpo = Tpo & UCase(Check(a).Caption) & " e "
    End If
Next
If Cont > 1 Then
    Aux = Aux & "s:"
Else
    Aux = Aux & ":"
End If
Aux = Aux & Left(Tpo, Len(Tpo) - 3)

LPRINT "De " & txDtINI.Text & " até " & txDtFIM.Text
If cbVend.ListIndex = 0 Then
    LPRINT "Todos Vendedores"
Else
    LPRINT "Vendedor: " & cbVend.Text
End If
LPRINT String(TamFita, "-")
    
If Total > 0 Then
    LPRINT "Valor: " & Format(Total, "##,###,###,##0.00")
    sValor = Total
End If

LPRINT " "
LPRINT " "
LPRINT String(TamFita, "-")

If ImprBuferizada_Finaliza = False Then
    Exit Sub
End If
End Sub
