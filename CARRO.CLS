VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCarro"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Sub CarregaCarros(cb As ComboBox)
Dim CAR As Recordset
Dim Selec As String
cb.Clear

Selec = "SELECT DISTINCTROW Carros.Modelo "
Selec = Selec + "FROM Clientes INNER JOIN Carros ON Clientes.NrCli = Carros.NrCli "
Selec = Selec + "WHERE ((Clientes.Nome='" + CliAtual + "')); "
Set CAR = DB.OpenRecordset(Selec, dbOpenDynaset)
On Error GoTo DeuGuru
CAR.MoveFirst
CAR.MoveLast
CAR.MoveFirst
nrCar = CAR.RecordCount
For a% = 1 To nrCar
    cb.AddItem CAR(0)
    CAR.MoveNext
Next
CAR.Close
Sai:
Exit Sub

DeuGuru:
If Err <> 3021 Then
    MsgBox Error$(Err)
End If
Resume Sai
End Sub

