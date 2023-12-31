VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form ReciboEdicao 
   Caption         =   "Edição do recibo"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4380
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4380
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txPagto 
      Height          =   315
      Left            =   2760
      TabIndex        =   14
      Top             =   1500
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Command2"
      Height          =   315
      Left            =   4800
      TabIndex        =   13
      Top             =   1380
      Width           =   195
   End
   Begin VB.TextBox txData 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   1500
      Width           =   975
   End
   Begin VB.ComboBox cbMecanico 
      Height          =   315
      ItemData        =   "ReciboEdicao.frx":0000
      Left            =   1080
      List            =   "ReciboEdicao.frx":0010
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   420
      Width           =   2715
   End
   Begin VB.TextBox txSemana 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   1140
      Width           =   2655
   End
   Begin VB.TextBox txValor 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   780
      Width           =   975
   End
   Begin VB.TextBox txFuncionario 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Frame Frame 
      Caption         =   "Observação"
      Height          =   1335
      Left            =   60
      TabIndex        =   10
      Top             =   1920
      Width           =   4215
      Begin RichTextLib.RichTextBox txDet 
         Height          =   1095
         Left            =   60
         TabIndex        =   5
         Top             =   180
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1931
         _Version        =   393217
         TextRTF         =   $"ReciboEdicao.frx":0048
      End
   End
   Begin VB.Label lbPagto 
      Caption         =   "Pagto"
      Height          =   255
      Index           =   5
      Left            =   2220
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "Data"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Período"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Valor"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Funcionário"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "ReciboEdicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'4.8.6 Edição de recibos

Option Explicit

Private ID As Long

Public Property Get NrVale() As Long
NrVale = ID
End Property

Public Property Let NrVale(ByVal vNewValue As Long)
Dim Pq    As Boolean
Dim sTipo As String
Dim Tb    As Recordset

ID = vNewValue
AbreTB Tb, "Select * From Vales Where ID = " & ID, dbOpenSnapshot
sTipo = TpRecs(Tb!Tipo)
Select Case Tb!Tipo
    Case 0
        lbPagto(5).Visible = True
        txPagto.Visible = True
    Case 2
        Pq = True
End Select
TextCombo sTipo, cbMecanico, 1
If Pq Then
    Frame.Visible = False
    ReciboEdicao.Height = ReciboEdicao.Height * 0.7
    Command1.Top = Command1.Top * 0.6
End If
txValor.Text = Format(Tb!Valor, "##,##0.00") & " "
txData.Text = Format(Tb!Data, "DD/MM/YYYY")
txDet.Text = SN(Tb!Obs, vbString)
txSemana.Text = SN(Tb!Periodo, vbString)
If Tb!PAGO > 0 Then
    txPagto.Text = Format(Tb!PAGO, "DD/MM/YYYY")
End If
txFuncionario.Text = Consulta("Select Nome From Mecanicos Where codi = " & Tb!IdOperador)
End Property

Private Sub Command1_Click()
Dim auxDT As Date
Dim SQL   As String

SQL = "Update Vales Set Tipo = " & cbMecanico.ItemData(cbMecanico.ListIndex)
SQL = SQL & " ,Valor = " & VlrSql(STR(txValor.Text))
SQL = SQL & " ,Data = " & DTSqls(txData.Text)
SQL = SQL & " ,Obs = " & FA(txDet.Text)
SQL = SQL & " ,Periodo = " & FA(txSemana.Text)
If txPagto.Text > "" Then
    On Local Error GoTo Erro_Command1_Click
    auxDT = DateValue(txPagto.Text)
    SQL = SQL & " ,Pago = " & DTSqls(Format(auxDT, "DD/MM/YYYY"), True)
End If

Continua_Command1_Click:
SQL = SQL & " Where ID = " & ID
ExecSql SQL
Unload Me
Exit Sub

Erro_Command1_Click:
Resume Continua_Command1_Click
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
