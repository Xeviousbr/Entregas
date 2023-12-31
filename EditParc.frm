VERSION 5.00
Begin VB.Form EditParc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pagamento da Parcela"
   ClientHeight    =   4065
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txData 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   600
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Observação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   60
      TabIndex        =   4
      Top             =   1200
      Width           =   5835
      Begin VB.TextBox txObs 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   420
         Width           =   5595
      End
   End
   Begin VB.TextBox txValor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   780
      TabIndex        =   3
      Text            =   "100"
      Top             =   540
      Width           =   1935
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label lbCli 
      Alignment       =   2  'Center
      Caption         =   "Cliente Fulano da Silva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   5820
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pagamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2820
      TabIndex        =   6
      Top             =   660
      Width           =   1440
   End
   Begin VB.Label lbValor 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   60
      TabIndex        =   2
      Top             =   600
      Width           =   660
   End
End
Attribute VB_Name = "EditParc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'3.9.6 Edição das parcelas

Option Explicit

Private lcValor   As String
Private lcObs     As String
Private lcData    As String
Private lcParcela As String
Private lccValor  As Currency

Public Property Let Parcela(ByVal vNewValue As String)
Dim SQL       As String
Dim rsParcela As Recordset

lcParcela = vNewValue
SQL = "SELECT Clientes.Nome, Parcelas.Valor, Parcelas.Obs "
SQL = SQL & "FROM Parcelas INNER JOIN Clientes ON Parcelas.Cli = Clientes.NrCli "
SQL = SQL & "WHERE Parcelas.idParc = " & lcParcela
AbreTB rsParcela, SQL, dbOpenDynaset
lbCli.Caption = "Cliente " & rsParcela!Nome
MostraValor txValor, rsParcela!Valor
txObs.Text = SN(rsParcela!Obs, vbString)
End Property

Public Property Let DtPago(ByVal vNewValue As String)
lcData = vNewValue
txData.Text = lcData
End Property

Public Property Let Valor(ByVal vNewValue As String)
lcValor = vNewValue
txValor.Text = lcValor
cValor = 0
End Property

Public Property Let cValor(ByVal vNewValue As Currency)
lccValor = vNewValue
End Property

Public Property Get Obs() As Currency
Obs = lcObs
End Property

Public Property Let Obs(ByVal vNewValue As Currency)
lcObs = vNewValue
End Property

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Form_Activate()
OKButton.Enabled = False
End Sub

Private Sub OKButton_Click()
Dim SQL    As String
Dim sValor As String
Dim Resp   As String
Dim sData  As String
Dim Valor  As Double
Dim rsMec  As Recordset

If lcData <> txData.Text And txData.Text > "" Then
    Dim Data  As Date

    sData = txData.Text
    If CriticaData(sData, Data) = 0 Then
        MsgBox "Data Inválida"
        Exit Sub
    End If
End If

frmSenha.Tipo = 1
frmSenha.Show 1
If frmSenha.Resultado = True Then
    Resp = frmSenha.Senha
End If
Unload frmSenha
If Resp = "" Then
    MsgBox "Operador não identificado"
Else
    Set rsMec = BuscaMec(Resp, " and Oper > 0 and Recebe = 1 ")
    If rsMec.EOF Then
        MsgBox "Operador não identificado"
    Else
        VeValor txValor.Text, Valor, txValor, 0
        sValor = VlrSql(STR(Valor))
        SQL = "Update Parcelas Set Valor = " & sValor
        If lcData <> txData.Text Then
            If Trim(txData.Text) = "" Then
                SQL = SQL & ", Pagto = null, BalcRec = null "
            Else
                sData = DTSqls(txData.Text)
                SQL = SQL & ", Pagto = " & sData
                SQL = SQL & ", BalcRec = " & rsMec!codi
            End If
        End If
        SQL = SQL & ", Obs = " & FA(txObs.Text)
        SQL = SQL & " Where idParc = " & lcParcela
        ExecSql SQL
        Unload Me
    End If
End If
End Sub

Private Sub txData_Change()
OKButton.Enabled = True
End Sub

Private Sub txObs_Change()
OKButton.Enabled = True
End Sub

Private Sub txValor_Change()
OKButton.Enabled = True
End Sub
