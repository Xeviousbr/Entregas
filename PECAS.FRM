VERSION 4.00
Begin VB.Form frmPecas 
   Caption         =   "Pecas"
   ClientHeight    =   1350
   ClientLeft      =   2355
   ClientTop       =   3015
   ClientWidth     =   5415
   Height          =   1755
   Icon            =   "PECAS.frx":0000
   Left            =   2295
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   5415
   Top             =   2670
   Width           =   5535
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   300
      Left            =   3360
      TabIndex        =   7
      Top             =   705
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Gravar"
      Height          =   300
      Left            =   2280
      TabIndex        =   6
      Top             =   705
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "A&pagar"
      Height          =   300
      Left            =   1200
      TabIndex        =   5
      Top             =   700
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Adicionar"
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   975
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   ""
      DatabaseName    =   "TROJAN.MDB"
      Exclusive       =   -1  'True
      Height          =   270
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Pecas"
      Top             =   1080
      Width           =   5415
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Valor"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Peca"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   720
      MaxLength       =   50
      TabIndex        =   1
      Top             =   45
      Width           =   4455
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Valor:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   375
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "Pe&ça:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   495
   End
End
Attribute VB_Name = "frmPecas"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
If txtFields(0).Enabled = False Then
   txtFields(0).Enabled = True
   txtFields(1).Enabled = True
End If
Data1.Recordset.AddNew
txtFields(0).SetFocus
cmdAdd.Enabled = False
cmdDelete.Enabled = False
End Sub

Private Sub cmdDelete_Click()
  'this may produce an error if you delete the last
  'record or the only record in the recordset
  Data1.Recordset.Delete
  Data1.Recordset.MoveNext
End Sub

Private Sub cmdUpdate_Click()
If txtFields(0).Text > "" Then
    Data1.UpdateRecord
    Data1.Recordset.Bookmark = Data1.Recordset.LastModified
End If
cmdAdd.Enabled = True
cmdDelete.Enabled = True
cmdAdd.SetFocus
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Data1_Error(DataErr As Integer, Response As Integer)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Error$(DataErr)
  Response = 0  'throw away the error
End Sub

Private Sub Data1_Reposition()
  'Screen.MousePointer = vbDefault
  On Error Resume Next
  'This will display the current record position
  'for dynasets and snapshots
  
  Posicao% = (Data1.Recordset.AbsolutePosition + 1)
  Data1.Caption = "Peça: " & Posicao% & " de" & str$(Data1.Recordset.RecordCount)
  
  'Data1.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1)
  'for the table object you must set the index property when
  'the recordset gets created and use the following line
  'Data1.Caption = "Record: " & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
  End Select
  'Screen.MousePointer = vbHourglass
End Sub

Private Sub Form_Activate()
On Error GoTo ComErro
Data1.Recordset.MoveFirst
Data1.Recordset.MoveLast

Sai:
If Data1.Recordset.RecordCount = 0 Then
   txtFields(0).Enabled = False
   txtFields(1).Enabled = False
End If
Exit Sub

ComErro:
Resume Sai
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
Seleciona
End Sub

Private Sub txtFields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
End Sub

Public Property Let NomePeca(vNewValue)
vNewValue = Chr$(34) + vNewValue + Chr$(34)
Data1.Recordset.FindFirst "[Peca] = " & vNewValue
End Property
