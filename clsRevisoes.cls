VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRevisoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'5.0.4 Cadastro de Revisões

Option Explicit

Private ArrRev(5) As Integer

Public Sub CarregaRevisoes()
Dim a      As Integer
Dim IdRevC As Integer
Dim SQL    As String
Dim tbRevs As Recordset

SQL = "SELECT Revisoes.ID, RV.IdRevCarros, Revisoes.Nome, RV.Data, RV.Kms "
SQL = SQL & "from Revisoes "
SQL = SQL & "LEFT JOIN (Select * From RevisoesCarros Where Placa = " & FA(frmClientes.txtFields(5).Text) & ") RV ON Revisoes.ID = RV.idRevisao "
AbreTB tbRevs, SQL
Do While tbRevs.EOF = False
    If a < 5 Then
        frmClientes.txRevNome(a).Text = tbRevs!Nome
        
        ArrRev(a) = tbRevs!ID
        
        IdRevC = SN(tbRevs!IdRevCarros, vbInteger)
        If IdRevC > 0 Then
            frmClientes.txRevDt(a).Text = Format(tbRevs!Data, "dd/mm/yyyy")
            frmClientes.ckRev(a).Tag = "1"
            frmClientes.ckRev(a).Value = 1
            frmClientes.txRevDt(a).Tag = Trim(STR(IdRevC))
        Else
            frmClientes.txRevNome(a).Tag = ""
            frmClientes.txRevDt(a).Text = ""
            frmClientes.ckRev(a).Tag = "0"
            frmClientes.ckRev(a).Value = 0
            frmClientes.txRevDt(a).Tag = "0"
        End If
        tbRevs.MoveNext
        a = a + 1
    Else
        Exit Do
    End If
Loop
End Sub

Public Sub ckRev_Click(Index As Integer)
Dim Revisao As String

Revisao = frmClientes.txRevNome(Index).Text
If Revisao = "" Then
    frmClientes.ckRev(Index).Value = 0
Else
    If frmClientes.ckRev(Index).Tag = "cliquei" Then
        frmClientes.ckRev(Index).Tag = Trim(STR(frmClientes.ckRev(Index).Value))
    Else
        If frmClientes.ckRev(Index).Tag = "0" Then
            If msgboxL("Confirma a criação da revisão de " & Revisao, vbYesNo + vbDefaultButton1) = vbYes Then
                CriaRevisao Index
            Else
                frmClientes.ckRev(Index).Tag = "cliquei"
                frmClientes.ckRev(Index).Value = 0
            End If
        Else
            If msgboxL("Tem certeza que quer apagar a revisão de " & Revisao, vbYesNo + vbDefaultButton2) = vbYes Then
                DeletaRevisao Index
            Else
                frmClientes.ckRev(Index).Tag = "cliquei"
                frmClientes.ckRev(Index).Value = 1
            End If
        End If
    End If
End If
End Sub

Private Sub DeletaRevisao(Index As Integer)
ExecSql "Delete From RevisoesCarros Where IdRevCarros = " & frmClientes.txRevDt(Index).Tag
frmClientes.txRevDt(Index).Text = ""
End Sub

Private Sub CriaRevisao(Index As Integer)
Dim SQL As String

SQL = "Insert Into RevisoesCarros (idRevisao, Data, Placa) Values ("
SQL = SQL & ArrRev(Index)
SQL = SQL & ", " & DTSqld(Int(Now))
SQL = SQL & ", " & FA(frmClientes.txtFields(5).Text) & ")"
ExecSql SQL
frmClientes.txRevDt(Index).Text = Format(Now, "dd/mm/yyyy")
End Sub

